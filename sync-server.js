#!/usr/bin/env node
/**
 * Локальный синх-сервер Meister для турниров по Wi-Fi.
 * CLI: node sync-server.js
 * Модуль: const { startSyncServer } = require('./sync-server');
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const dgram = require('dgram');
const os = require('os');

const DEFAULT_PORT = 41235;
const DISCOVERY_PORT = 41234;

const MIME = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'application/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.ico': 'image/x-icon'
};

function getLocalIps() {
    try {
        const ips = [];
        const nets = os.networkInterfaces();
        for (const name of Object.keys(nets)) {
            for (const net of nets[name] || []) {
                if (net.family === 'IPv4' && !net.internal) {
                    ips.push(net.address);
                }
            }
        }
        return ips;
    } catch (e) {
        return [];
    }
}

function buildServerUrls(port, ips) {
    const lanUrls = (ips || []).map(function(ip) {
        return { url: 'http://' + ip + ':' + port + '/', local: false };
    });
    return lanUrls.concat([{ url: 'http://127.0.0.1:' + port + '/', local: true }]);
}

function createSyncServer(options) {
    options = options || {};
    const port = parseInt(options.port || process.env.MEISTER_PORT || process.env.GLADIAGON_PORT || DEFAULT_PORT, 10);
    const root = options.root || __dirname;
    const silent = !!options.silent;
    const onLog = typeof options.onLog === 'function' ? options.onLog : null;
    const maxLogs = options.maxLogs || 500;
    const logs = [];

    function log(level, message, meta) {
        const entry = {
            time: new Date().toISOString(),
            level: level,
            message: message,
            meta: meta || null
        };
        logs.push(entry);
        if (logs.length > maxLogs) logs.shift();
        const line = '[' + entry.time + '] [' + level + '] ' + message +
            (meta ? ' ' + JSON.stringify(meta) : '');
        if (!silent) console.log(line);
        if (onLog) onLog(entry, line);
    }

    function getClientIp(req) {
        const forwarded = req.headers['x-forwarded-for'];
        if (forwarded) return String(forwarded).split(',')[0].trim();
        return req.socket && req.socket.remoteAddress
            ? String(req.socket.remoteAddress).replace(/^::ffff:/, '')
            : 'unknown';
    }

    let state = {
        version: 0,
        arenaCount: 0,
        hostDeviceId: null,
        tournament: null,
        devices: {},
        matchLocks: {}
    };

    const sseClients = new Set();
    let discoverySocket = null;
    let discoveryTimer = null;
    let httpServer = null;
    let running = false;

    function getPublicState() {
        return {
            version: state.version,
            arenaCount: state.arenaCount,
            hostDeviceId: state.hostDeviceId,
            tournament: state.tournament,
            devices: state.devices,
            matchLocks: state.matchLocks
        };
    }

    function broadcastState() {
        const payload = JSON.stringify(getPublicState());
        for (const res of sseClients) {
            res.write('event: state\n');
            res.write('data: ' + payload + '\n\n');
        }
    }

    function bumpVersion() {
        state.version += 1;
        broadcastState();
    }

    function poolMatchKey(poolId, matchId) {
        return 'pool:' + poolId + ':' + matchId;
    }

    function bracketMatchKey(loc) {
        return 'bracket:' + loc.side + ':' + loc.roundIndex + ':' + loc.matchIndex;
    }

    function readJsonBody(req) {
        return new Promise(function(resolve, reject) {
            let body = '';
            req.on('data', function(chunk) {
                body += chunk;
                if (body.length > 2 * 1024 * 1024) {
                    reject(new Error('Body too large'));
                    req.destroy();
                }
            });
            req.on('end', function() {
                if (!body) {
                    resolve({});
                    return;
                }
                try {
                    resolve(JSON.parse(body));
                } catch (e) {
                    reject(e);
                }
            });
            req.on('error', reject);
        });
    }

    function sendJson(res, status, data) {
        const body = JSON.stringify(data);
        res.writeHead(status, {
            'Content-Type': 'application/json; charset=utf-8',
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type'
        });
        res.end(body);
    }

    function getDevice(body) {
        if (!body || !body.deviceId) return null;
        return state.devices[body.deviceId] || null;
    }

    function isHost(body) {
        return body && body.deviceId && body.deviceId === state.hostDeviceId;
    }

    function findPoolMatch(tournament, poolId, matchId) {
        const matches = tournament && tournament.poolMatches && tournament.poolMatches[poolId];
        if (!matches) return null;
        for (let i = 0; i < matches.length; i++) {
            if (matches[i].id === matchId) return matches[i];
        }
        return null;
    }

    function getBracketMatch(tournament, loc) {
        const bracket = tournament && tournament.bracket;
        if (!bracket) return null;
        if (loc.side === 'final') return bracket.final;
        if (loc.side === 'bronze') return bracket.thirdPlace;
        const half = bracket[loc.side];
        if (!half || !half.rounds || !half.rounds[loc.roundIndex]) return null;
        return half.rounds[loc.roundIndex][loc.matchIndex] || null;
    }

    const PLAYOFF_PHASE_ORDER = { side: 0, bronze: 1, final: 2, complete: 3 };

    function mergeMatchPreservingExisting(existingMatch, incomingMatch) {
        if (!existingMatch) return incomingMatch;
        if (!incomingMatch) return existingMatch;
        if (existingMatch.status === 'done') return existingMatch;
        if (existingMatch.status === 'in_progress' &&
            (incomingMatch.status === 'pending' || !incomingMatch.status)) {
            return existingMatch;
        }
        return incomingMatch;
    }

    function mergePoolMatchesPreservingExisting(existingPm, incomingPm) {
        const merged = incomingPm ? JSON.parse(JSON.stringify(incomingPm)) : {};
        if (!existingPm) return merged;
        for (const poolId of Object.keys(existingPm)) {
            const existingMatches = existingPm[poolId];
            if (!existingMatches) continue;
            if (!merged[poolId]) {
                merged[poolId] = JSON.parse(JSON.stringify(existingMatches));
                continue;
            }
            const byId = {};
            for (let i = 0; i < merged[poolId].length; i++) {
                byId[merged[poolId][i].id] = merged[poolId][i];
            }
            for (let j = 0; j < existingMatches.length; j++) {
                const existingMatch = existingMatches[j];
                const incomingMatch = byId[existingMatch.id];
                if (!incomingMatch) {
                    merged[poolId].push(JSON.parse(JSON.stringify(existingMatch)));
                    continue;
                }
                if (existingMatch.status === 'done' && incomingMatch.status !== 'done') {
                    const idx = merged[poolId].indexOf(incomingMatch);
                    merged[poolId][idx] = JSON.parse(JSON.stringify(existingMatch));
                } else if (existingMatch.status === 'in_progress' &&
                    (incomingMatch.status === 'pending' || !incomingMatch.status)) {
                    const idx = merged[poolId].indexOf(incomingMatch);
                    merged[poolId][idx] = JSON.parse(JSON.stringify(existingMatch));
                }
            }
        }
        return merged;
    }

    function mergeBracketPreservingExisting(existingBracket, incomingBracket) {
        if (!incomingBracket) {
            return existingBracket ? JSON.parse(JSON.stringify(existingBracket)) : null;
        }
        if (!existingBracket) return JSON.parse(JSON.stringify(incomingBracket));

        const merged = JSON.parse(JSON.stringify(incomingBracket));

        function mergeHalf(sideName) {
            const existingHalf = existingBracket[sideName];
            const incomingHalf = merged[sideName];
            if (!existingHalf || !incomingHalf || !incomingHalf.rounds) return;
            for (let r = 0; r < incomingHalf.rounds.length; r++) {
                if (!existingHalf.rounds[r]) continue;
                for (let m = 0; m < incomingHalf.rounds[r].length; m++) {
                    incomingHalf.rounds[r][m] = mergeMatchPreservingExisting(
                        existingHalf.rounds[r][m],
                        incomingHalf.rounds[r][m]
                    );
                }
            }
        }

        mergeHalf('left');
        mergeHalf('right');
        merged.thirdPlace = mergeMatchPreservingExisting(
            existingBracket.thirdPlace,
            merged.thirdPlace
        );
        merged.final = mergeMatchPreservingExisting(existingBracket.final, merged.final);

        if ((existingBracket.activeRoundIndex || 0) > (merged.activeRoundIndex || 0)) {
            merged.activeRoundIndex = existingBracket.activeRoundIndex;
        }
        const existingPhase = PLAYOFF_PHASE_ORDER[existingBracket.playoffPhase || 'side'] || 0;
        const incomingPhase = PLAYOFF_PHASE_ORDER[merged.playoffPhase || 'side'] || 0;
        if (existingPhase > incomingPhase) {
            merged.playoffPhase = existingBracket.playoffPhase;
        }

        return merged;
    }

    function mergeTournamentPreservingExisting(existing, incoming) {
        if (!existing) return incoming;
        if (!incoming) return existing;
        const merged = JSON.parse(JSON.stringify(incoming));
        merged.poolMatches = mergePoolMatchesPreservingExisting(
            existing.poolMatches,
            incoming.poolMatches || {}
        );
        if (incoming.bracket || existing.bracket) {
            merged.bracket = mergeBracketPreservingExisting(existing.bracket, incoming.bracket);
        }
        if (existing.tournamentFightHistory && existing.tournamentFightHistory.length &&
            (!incoming.tournamentFightHistory || incoming.tournamentFightHistory.length <
                existing.tournamentFightHistory.length)) {
            merged.tournamentFightHistory = existing.tournamentFightHistory.slice();
        }
        return merged;
    }

    function releaseLocksForDevice(deviceId) {
        const keys = Object.keys(state.matchLocks);
        for (let i = 0; i < keys.length; i++) {
            if (state.matchLocks[keys[i]].deviceId === deviceId) {
                delete state.matchLocks[keys[i]];
            }
        }
    }

    function getPoolAssignedArena(tournament, poolId) {
        if (!tournament || !tournament.poolArenaAssignments) return null;
        const val = tournament.poolArenaAssignments[poolId];
        return val ? parseInt(val, 10) : null;
    }

    function isPoolMatchAllowedForArena(tournament, poolId, arenaId, hostDevice) {
        if (hostDevice) return true;
        const assigned = getPoolAssignedArena(tournament, poolId);
        return assigned !== null && assigned === arenaId;
    }

    function isPoolStageTournament(tournament) {
        return !!(tournament && tournament.playoffStarted && !tournament.bracket);
    }

    function areArenaPoolFightsComplete(tournament, arenaId) {
        if (!tournament || !tournament.pools) return false;
        arenaId = parseInt(arenaId, 10);
        let hasAssigned = false;
        for (let i = 0; i < tournament.pools.length; i++) {
            const poolId = tournament.pools[i].id;
            if (getPoolAssignedArena(tournament, poolId) !== arenaId) continue;
            hasAssigned = true;
            const matches = (tournament.poolMatches && tournament.poolMatches[poolId]) || [];
            for (let m = 0; m < matches.length; m++) {
                if (matches[m].status !== 'done') return false;
            }
        }
        return hasAssigned;
    }

    function findNextPendingPoolMatchForArena(tournament, arenaId) {
        if (!tournament || !tournament.pools) return null;
        arenaId = parseInt(arenaId, 10);
        for (let i = 0; i < tournament.pools.length; i++) {
            const poolId = tournament.pools[i].id;
            if (getPoolAssignedArena(tournament, poolId) !== arenaId) continue;
            const matches = (tournament.poolMatches && tournament.poolMatches[poolId]) || [];
            for (let m = 0; m < matches.length; m++) {
                if (matches[m].status === 'pending') {
                    return { poolId: poolId, matchId: matches[m].id, match: matches[m] };
                }
            }
        }
        return null;
    }

    function hostClaimPoolMatchForArena(arenaId, poolId, matchId) {
        if (!state.hostDeviceId || !state.tournament) return null;
        const match = findPoolMatch(state.tournament, poolId, matchId);
        if (!match || match.status !== 'pending') return null;
        const matchKey = poolMatchKey(poolId, matchId);
        if (state.matchLocks[matchKey]) return null;
        match.status = 'in_progress';
        match.arenaId = arenaId;
        state.matchLocks[matchKey] = {
            deviceId: state.hostDeviceId,
            arenaId: arenaId,
            matchType: 'pool',
            poolId: poolId,
            matchId: matchId,
            bracketLoc: null,
            arenaDeviceId: null,
            since: Date.now()
        };
        return matchKey;
    }

    function autoDispatchNextPoolFightForArena(arenaId) {
        if (!state.hostDeviceId || !isPoolStageTournament(state.tournament)) return;
        arenaId = parseInt(arenaId, 10);
        if (!arenaId || areArenaPoolFightsComplete(state.tournament, arenaId)) return;
        const next = findNextPendingPoolMatchForArena(state.tournament, arenaId);
        if (next) {
            hostClaimPoolMatchForArena(arenaId, next.poolId, next.matchId);
        }
    }

    function getBracketAssignedArena(tournament, phaseKey) {
        if (!tournament || !tournament.bracketArenaAssignments) return null;
        const val = tournament.bracketArenaAssignments[phaseKey];
        return val ? parseInt(val, 10) : null;
    }

    function getBracketPhaseKeyFromLoc(loc) {
        if (!loc) return null;
        if (loc.side === 'bronze') return 'bronze';
        if (loc.side === 'final') return 'final';
        if (loc.side === 'left' || loc.side === 'right') {
            return loc.side + ':' + loc.roundIndex;
        }
        return null;
    }

    function resolveBracketArenaAssignment(tournament, phaseKey) {
        if (!tournament || !phaseKey) return null;
        const assignments = tournament.bracketArenaAssignments || {};
        const direct = assignments[phaseKey];
        if (direct) return parseInt(direct, 10);
        if (phaseKey.indexOf('left:') === 0 || phaseKey.indexOf('right:') === 0) {
            const legacyKey = 'side:' + phaseKey.split(':')[1];
            const legacy = assignments[legacyKey];
            if (legacy) return parseInt(legacy, 10);
        }
        return null;
    }

    function isBracketMatchReadyServer(match) {
        if (!match || !match.fighter1 || !match.fighter2) return false;
        if (match.fighter1.isBye || match.fighter2.isBye) return false;
        return match.status === 'pending';
    }

    function findNextPendingBracketMatchForArena(tournament, arenaId) {
        const bracket = tournament && tournament.bracket;
        if (!bracket) return null;
        arenaId = parseInt(arenaId, 10);
        const activeR = bracket.activeRoundIndex != null ? bracket.activeRoundIndex : 0;
        const phase = bracket.playoffPhase || 'side';

        function tryLoc(match, loc) {
            const phaseKey = getBracketPhaseKeyFromLoc(loc);
            if (!phaseKey || resolveBracketArenaAssignment(tournament, phaseKey) !== arenaId) return null;
            if (!isBracketMatchReadyServer(match)) return null;
            if (phase === 'side' && loc.side !== 'bronze' && loc.side !== 'final' &&
                loc.roundIndex !== activeR) {
                return null;
            }
            if (phase === 'bronze' && loc.side !== 'bronze') return null;
            if (phase === 'final' && loc.side !== 'final') return null;
            return { match: match, loc: loc };
        }

        for (const side of ['left', 'right']) {
            const half = bracket[side];
            if (!half || !half.rounds) continue;
            for (let r = 0; r < half.rounds.length; r++) {
                for (let m = 0; m < half.rounds[r].length; m++) {
                    const found = tryLoc(half.rounds[r][m], {
                        side: side,
                        roundIndex: r,
                        matchIndex: m
                    });
                    if (found) return found;
                }
            }
        }
        let found = tryLoc(bracket.thirdPlace, { side: 'bronze', roundIndex: 0, matchIndex: 0 });
        if (found) return found;
        found = tryLoc(bracket.final, { side: 'final', roundIndex: 0, matchIndex: 0 });
        return found;
    }

    function hostClaimBracketMatchForArena(arenaId, loc) {
        if (!state.hostDeviceId || !state.tournament || !loc) return null;
        const match = getBracketMatch(state.tournament, loc);
        if (!match || match.status !== 'pending') return null;
        const matchKey = bracketMatchKey(loc);
        if (state.matchLocks[matchKey]) return null;
        match.status = 'in_progress';
        match.arenaId = arenaId;
        state.matchLocks[matchKey] = {
            deviceId: state.hostDeviceId,
            arenaId: arenaId,
            matchType: 'bracket',
            poolId: null,
            matchId: null,
            bracketLoc: loc,
            arenaDeviceId: null,
            since: Date.now()
        };
        return matchKey;
    }

    function autoDispatchNextBracketFightForArena(arenaId) {
        if (!state.hostDeviceId || !state.tournament || !state.tournament.bracket) return;
        arenaId = parseInt(arenaId, 10);
        if (!arenaId) return;
        const next = findNextPendingBracketMatchForArena(state.tournament, arenaId);
        if (next) {
            hostClaimBracketMatchForArena(arenaId, next.loc);
        }
    }

    function reopenMatchForEdit(match) {
        match.status = 'in_progress';
        delete match.winnerId;
        delete match.winner;
        delete match.result;
    }

    function serveStatic(req, res, urlPath) {
        let filePath = urlPath === '/' ? '/secretary_terminal.html' : urlPath;
        filePath = decodeURIComponent(filePath.split('?')[0]);
        const abs = path.normalize(path.join(root, filePath));
        if (!abs.startsWith(root)) {
            sendJson(res, 403, { error: 'Forbidden' });
            return;
        }
        fs.readFile(abs, function(err, data) {
            if (err) {
                sendJson(res, 404, { error: 'Not found' });
                return;
            }
            const ext = path.extname(abs).toLowerCase();
            res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream' });
            res.end(data);
        });
    }

    async function handleApi(req, res, url) {
        if (req.method === 'OPTIONS') {
            res.writeHead(204, {
                'Access-Control-Allow-Origin': '*',
                'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
                'Access-Control-Allow-Headers': 'Content-Type'
            });
            res.end();
            return;
        }

        if (req.method === 'GET' && url.pathname === '/api/ping') {
            sendJson(res, 200, {
                ok: true,
                name: os.hostname(),
                port: port,
                ips: getLocalIps(),
                version: state.version,
                hasTournament: !!state.tournament,
                arenaCount: state.arenaCount
            });
            return;
        }

        if (req.method === 'GET' && url.pathname === '/api/state') {
            sendJson(res, 200, getPublicState());
            return;
        }

        if (req.method === 'GET' && url.pathname === '/api/events') {
            res.writeHead(200, {
                'Content-Type': 'text/event-stream; charset=utf-8',
                'Cache-Control': 'no-cache',
                'Connection': 'keep-alive',
                'Access-Control-Allow-Origin': '*'
            });
            res.write(': connected\n\n');
            sseClients.add(res);
            res.write('event: state\n');
            res.write('data: ' + JSON.stringify(getPublicState()) + '\n\n');
            req.on('close', function() {
                sseClients.delete(res);
            });
            return;
        }

        if (req.method !== 'POST') {
            sendJson(res, 405, { error: 'Method not allowed' });
            return;
        }

        let body;
        try {
            body = await readJsonBody(req);
        } catch (e) {
            sendJson(res, 400, { error: 'Invalid JSON' });
            return;
        }

        if (url.pathname === '/api/register') {
            const deviceId = body.deviceId;
            const role = body.role;
            if (!deviceId || (role !== 'host' && role !== 'arena')) {
                sendJson(res, 400, { error: 'deviceId and role required' });
                return;
            }
            if (role === 'host') {
                if (state.hostDeviceId && state.hostDeviceId !== deviceId) {
                    sendJson(res, 409, { error: 'Host already registered' });
                    return;
                }
                state.hostDeviceId = deviceId;
            } else if (role === 'arena') {
                const arenaId = parseInt(body.arenaId, 10);
                if (!arenaId || arenaId < 1) {
                    sendJson(res, 400, { error: 'arenaId required for arena role' });
                    return;
                }
                if (state.arenaCount && arenaId > state.arenaCount) {
                    sendJson(res, 400, { error: 'arenaId exceeds arenaCount' });
                    return;
                }
                for (const id of Object.keys(state.devices)) {
                    const d = state.devices[id];
                    if (d.role === 'arena' && d.arenaId === arenaId && id !== deviceId) {
                        sendJson(res, 409, { error: 'Arena already taken' });
                        return;
                    }
                }
                body.arenaId = arenaId;
            }

            state.devices[deviceId] = {
                role: role,
                name: body.name || (role === 'host' ? 'Главное устройство' : 'Площадка ' + body.arenaId),
                arenaId: role === 'arena' ? body.arenaId : null,
                connectedAt: Date.now()
            };
            bumpVersion();
            sendJson(res, 200, { ok: true, state: getPublicState() });
            return;
        }

        if (url.pathname === '/api/arena-count') {
            if (!isHost(body)) {
                sendJson(res, 403, { error: 'Host only' });
                return;
            }
            const count = parseInt(body.arenaCount, 10);
            if (!count || count < 1 || count > 32) {
                sendJson(res, 400, { error: 'arenaCount must be 1-32' });
                return;
            }
            state.arenaCount = count;
            bumpVersion();
            sendJson(res, 200, { ok: true, state: getPublicState() });
            return;
        }

        if (url.pathname === '/api/tournament') {
            if (!isHost(body)) {
                sendJson(res, 403, { error: 'Host only' });
                return;
            }
            if (!body.tournament || typeof body.tournament !== 'object') {
                sendJson(res, 400, { error: 'tournament object required' });
                return;
            }
            state.tournament = mergeTournamentPreservingExisting(state.tournament, body.tournament);
            bumpVersion();
            sendJson(res, 200, { ok: true, state: getPublicState() });
            return;
        }

        if (url.pathname === '/api/match/claim') {
            const device = getDevice(body);
            if (!device || !isHost(body)) {
                sendJson(res, 403, { error: 'Only the host device can assign fights' });
                return;
            }
            const arenaId = parseInt(body.arenaId, 10) || device.arenaId;
            if (!arenaId) {
                sendJson(res, 400, { error: 'arenaId required' });
                return;
            }
            if (!state.tournament) {
                sendJson(res, 400, { error: 'Tournament not started' });
                return;
            }

            let matchKey;
            let match;

            if (body.matchType === 'pool') {
                match = findPoolMatch(state.tournament, body.poolId, body.matchId);
                if (!match) {
                    sendJson(res, 404, { error: 'Match not found' });
                    return;
                }
                matchKey = poolMatchKey(body.poolId, body.matchId);
            } else if (body.matchType === 'bracket') {
                const loc = body.bracketLoc;
                match = getBracketMatch(state.tournament, loc);
                if (!match) {
                    sendJson(res, 404, { error: 'Bracket match not found' });
                    return;
                }
                matchKey = bracketMatchKey(loc);
            } else {
                sendJson(res, 400, { error: 'matchType required' });
                return;
            }

            const reopen = !!body.reopen;

            if (match.status === 'done') {
                if (!reopen) {
                    sendJson(res, 409, { error: 'Match already done' });
                    return;
                }
                reopenMatchForEdit(match);
            }

            const existing = state.matchLocks[matchKey];
            if (existing && existing.deviceId !== body.deviceId) {
                sendJson(res, 409, { error: 'Match locked by another device', lock: existing });
                return;
            }

            if (match.status !== 'in_progress') {
                match.status = 'in_progress';
            }
            match.arenaId = arenaId;
            state.matchLocks[matchKey] = {
                deviceId: body.deviceId,
                arenaId: arenaId,
                matchType: body.matchType,
                poolId: body.poolId || null,
                matchId: body.matchId || null,
                bracketLoc: body.bracketLoc || null,
                arenaDeviceId: null,
                since: Date.now()
            };
            bumpVersion();
            sendJson(res, 200, { ok: true, matchKey: matchKey, state: getPublicState() });
            return;
        }

        if (url.pathname === '/api/match/join') {
            const device = getDevice(body);
            if (!device || device.role !== 'arena') {
                sendJson(res, 403, { error: 'Arena devices only' });
                return;
            }
            const matchKey = body.matchKey;
            if (!matchKey) {
                sendJson(res, 400, { error: 'matchKey required' });
                return;
            }
            const lock = state.matchLocks[matchKey];
            if (!lock || lock.deviceId !== state.hostDeviceId) {
                sendJson(res, 403, { error: 'Fight not assigned by host' });
                return;
            }
            const lockArenaId = parseInt(lock.arenaId, 10);
            const deviceArenaId = parseInt(device.arenaId, 10);
            if (!lockArenaId || lockArenaId !== deviceArenaId) {
                sendJson(res, 403, { error: 'Fight assigned to another arena' });
                return;
            }
            lock.arenaDeviceId = body.deviceId;
            bumpVersion();
            sendJson(res, 200, { ok: true, matchKey: matchKey, state: getPublicState() });
            return;
        }

        if (url.pathname === '/api/match/complete') {
            const device = getDevice(body);
            if (!device) {
                sendJson(res, 403, { error: 'Unknown device' });
                return;
            }

            const matchKey = body.matchKey;
            const lock = state.matchLocks[matchKey];
            const hostCompleting = isHost(body);

            if (lock && lock.deviceId !== body.deviceId && !hostCompleting) {
                const arenaMayComplete = device.role === 'arena' &&
                    lock.arenaId === device.arenaId &&
                    state.hostDeviceId === lock.deviceId &&
                    (!lock.arenaDeviceId || lock.arenaDeviceId === body.deviceId);
                if (!arenaMayComplete) {
                    sendJson(res, 403, { error: 'Not match owner' });
                    return;
                }
            }

            if (!body.tournament || typeof body.tournament !== 'object') {
                sendJson(res, 400, { error: 'tournament snapshot required' });
                return;
            }

            state.tournament = body.tournament;
            if (matchKey && state.matchLocks[matchKey]) {
                delete state.matchLocks[matchKey];
            }
            if (device.role === 'arena' && device.arenaId) {
                if (isPoolStageTournament(state.tournament)) {
                    autoDispatchNextPoolFightForArena(device.arenaId);
                } else if (state.tournament && state.tournament.bracket) {
                    autoDispatchNextBracketFightForArena(device.arenaId);
                }
            }
            bumpVersion();
            sendJson(res, 200, { ok: true, state: getPublicState() });
            return;
        }

        if (url.pathname === '/api/match/release') {
            const lock = state.matchLocks[body.matchKey];
            if (lock && lock.deviceId !== body.deviceId && !isHost(body)) {
                sendJson(res, 403, { error: 'Not match owner' });
                return;
            }
            if (body.matchKey) {
                delete state.matchLocks[body.matchKey];
            }
            bumpVersion();
            sendJson(res, 200, { ok: true, state: getPublicState() });
            return;
        }

        if (url.pathname === '/api/disconnect') {
            if (body.deviceId && state.devices[body.deviceId]) {
                releaseLocksForDevice(body.deviceId);
                delete state.devices[body.deviceId];
                if (state.hostDeviceId === body.deviceId) {
                    state.hostDeviceId = null;
                }
                bumpVersion();
            }
            sendJson(res, 200, { ok: true, state: getPublicState() });
            return;
        }

        sendJson(res, 404, { error: 'Unknown API route' });
    }

    function startDiscoveryBroadcast() {
        discoverySocket = dgram.createSocket('udp4');
        discoverySocket.bind(function() {
            discoverySocket.setBroadcast(true);
        });

        const ips = getLocalIps();
        const primaryIp = ips[0] || '127.0.0.1';
        const message = Buffer.from('MEISTER|' + port + '|' + primaryIp + '|' + os.hostname());

        discoveryTimer = setInterval(function() {
            try {
                discoverySocket.send(message, 0, message.length, DISCOVERY_PORT, '255.255.255.255');
            } catch (e) {
                // ignore broadcast errors
            }
        }, 3000);
    }

    function getInfo() {
        const ips = getLocalIps();
        const urlEntries = buildServerUrls(port, ips);
        return {
            running: running,
            port: port,
            ips: ips,
            urls: urlEntries.map(function(item) { return item.url; }),
            urlEntries: urlEntries,
            localUrl: 'http://127.0.0.1:' + port + '/',
            lanUrls: urlEntries.filter(function(item) { return !item.local; }).map(function(item) { return item.url; }),
            discoveryPort: DISCOVERY_PORT,
            hostname: os.hostname()
        };
    }

    function stop() {
        return new Promise(function(resolve) {
            if (discoveryTimer) {
                clearInterval(discoveryTimer);
                discoveryTimer = null;
            }
            if (discoverySocket) {
                try { discoverySocket.close(); } catch (e) { /* ignore */ }
                discoverySocket = null;
            }
            for (const res of sseClients) {
                try { res.end(); } catch (e) { /* ignore */ }
            }
            sseClients.clear();
            if (!httpServer) {
                running = false;
                resolve();
                return;
            }
            httpServer.close(function() {
                log('info', 'Server stopped');
                httpServer = null;
                running = false;
                resolve();
            });
        });
    }

    function start() {
        if (running) {
            return Promise.resolve(getInfo());
        }

        httpServer = http.createServer(async function(req, res) {
            const url = new URL(req.url, 'http://localhost');
            const clientIp = getClientIp(req);
            if (url.pathname.startsWith('/api/')) {
                log('info', req.method + ' ' + url.pathname, { ip: clientIp });
                try {
                    await handleApi(req, res, url);
                } catch (e) {
                    log('error', 'API error ' + url.pathname + ': ' + (e.message || e), { ip: clientIp });
                    sendJson(res, 500, { error: e.message || 'Server error' });
                }
                return;
            }
            if (url.pathname === '/' || url.pathname.indexOf('.') !== -1) {
                log('info', 'GET ' + url.pathname, { ip: clientIp });
            }
            serveStatic(req, res, url.pathname);
        });

        return new Promise(function(resolve, reject) {
            httpServer.on('error', function(err) {
                log('error', 'HTTP server error: ' + (err.message || err));
                reject(err);
            });
            httpServer.listen(port, '0.0.0.0', function() {
                running = true;
                startDiscoveryBroadcast();
                const info = getInfo();
                log('info', 'Server started on port ' + port, {
                    ips: info.ips,
                    lanUrls: info.lanUrls
                });
                if (!silent) {
                    console.log('');
                    console.log('  Meister sync-server');
                    console.log('  HTTP port:', port);
                    console.log('  UDP discovery port:', DISCOVERY_PORT);
                    console.log('');
                    info.lanUrls.forEach(function(url) {
                        console.log('  LAN:', url);
                    });
                    console.log('  Local:', info.localUrl);
                    console.log('');
                }
                resolve(info);
            });
        });
    }

    return {
        start: start,
        stop: stop,
        getInfo: getInfo,
        getLogs: function() { return logs.slice(); },
        log: log,
        isRunning: function() { return running; },
        getPort: function() { return port; }
    };
}

function startSyncServer(options) {
    const instance = createSyncServer(options);
    return instance.start().then(function(info) {
        return {
            info: info,
            stop: instance.stop,
            getInfo: instance.getInfo,
            isRunning: instance.isRunning
        };
    });
}

module.exports = {
    createSyncServer: createSyncServer,
    startSyncServer: startSyncServer,
    DEFAULT_PORT: DEFAULT_PORT,
    DISCOVERY_PORT: DISCOVERY_PORT,
    getLocalIps: getLocalIps,
    buildServerUrls: buildServerUrls
};

if (require.main === module) {
    startSyncServer({ root: __dirname, silent: false }).catch(function(err) {
        console.error(err.message || err);
        process.exit(1);
    });
}
