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

    function releaseLocksForDevice(deviceId) {
        const keys = Object.keys(state.matchLocks);
        for (let i = 0; i < keys.length; i++) {
            if (state.matchLocks[keys[i]].deviceId === deviceId) {
                delete state.matchLocks[keys[i]];
            }
        }
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
            state.tournament = body.tournament;
            bumpVersion();
            sendJson(res, 200, { ok: true, state: getPublicState() });
            return;
        }

        if (url.pathname === '/api/match/claim') {
            const device = getDevice(body);
            if (!device || (device.role !== 'arena' && !isHost(body))) {
                sendJson(res, 403, { error: 'Arena or host device only' });
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

            if (match.status === 'done') {
                sendJson(res, 409, { error: 'Match already done' });
                return;
            }

            const existing = state.matchLocks[matchKey];
            if (existing && existing.deviceId !== body.deviceId) {
                sendJson(res, 409, { error: 'Match locked by another device', lock: existing });
                return;
            }

            match.status = 'in_progress';
            match.arenaId = arenaId;
            state.matchLocks[matchKey] = {
                deviceId: body.deviceId,
                arenaId: arenaId,
                matchType: body.matchType,
                poolId: body.poolId || null,
                matchId: body.matchId || null,
                bracketLoc: body.bracketLoc || null,
                since: Date.now()
            };
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
                sendJson(res, 403, { error: 'Not match owner' });
                return;
            }

            if (!body.tournament || typeof body.tournament !== 'object') {
                sendJson(res, 400, { error: 'tournament snapshot required' });
                return;
            }

            state.tournament = body.tournament;
            if (matchKey && state.matchLocks[matchKey]) {
                delete state.matchLocks[matchKey];
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
