(function(global) {
    'use strict';

    var DEFAULT_PORT = 41235;
    var SCAN_TIMEOUT_MS = 800;
    var CONNECT_TIMEOUT_MS = 5000;
    var SCAN_BATCH = 40;

    var networkState = {
        enabled: false,
        serverUrl: null,
        deviceId: null,
        deviceName: null,
        role: null,
        arenaId: null,
        arenaCount: 0,
        connected: false,
        remoteVersion: 0,
        eventSource: null,
        activeMatchKey: null,
        lastRemoteState: null,
        hostDeviceId: null
    };

    var onStateCallbacks = [];
    var onConnectionCallbacks = [];

    var STORAGE_KEYS = {
        deviceId: 'meisterDeviceId',
        deviceName: 'meisterDeviceName',
        serverUrl: 'meisterServerUrl'
    };

    var LEGACY_KEYS = {
        meisterDeviceId: 'gladiagonDeviceId',
        meisterDeviceName: 'gladiagonDeviceName',
        meisterServerUrl: 'gladiagonServerUrl'
    };

    function getStorageItem(key) {
        var val = localStorage.getItem(STORAGE_KEYS[key]);
        if (val !== null && val !== '') return val;
        var legacyKey = LEGACY_KEYS[STORAGE_KEYS[key]];
        if (legacyKey) {
            val = localStorage.getItem(legacyKey);
            if (val !== null && val !== '') {
                localStorage.setItem(STORAGE_KEYS[key], val);
                return val;
            }
        }
        return null;
    }

    function setStorageItem(key, value) {
        localStorage.setItem(STORAGE_KEYS[key], value);
    }

    function generateDeviceId() {
        var stored = getStorageItem('deviceId');
        if (stored) return stored;
        stored = 'dev_' + Date.now() + '_' + Math.floor(Math.random() * 100000);
        setStorageItem('deviceId', stored);
        return stored;
    }

    function getDefaultDeviceName() {
        var stored = getStorageItem('deviceName');
        if (stored) return stored;
        var ua = navigator.userAgent || '';
        if (/iPad|Tablet/i.test(ua)) return 'Планшет';
        if (/Mobile/i.test(ua)) return 'Телефон';
        return 'Компьютер';
    }

    function setDeviceName(name) {
        networkState.deviceName = (name || '').trim() || getDefaultDeviceName();
        setStorageItem('deviceName', networkState.deviceName);
    }

    function normalizeServerUrl(input) {
        var raw = (input || '').trim();
        if (!raw) return null;
        if (!/^https?:\/\//i.test(raw)) {
            raw = 'http://' + raw;
        }
        raw = raw.replace(/\/+$/, '');
        return raw;
    }

    function buildServerUrl(host, port) {
        return 'http://' + host + ':' + (port || DEFAULT_PORT);
    }

    function apiUrl(path) {
        return networkState.serverUrl + path;
    }

    function fetchJson(path, options) {
        return fetch(apiUrl(path), options || {}).then(function(res) {
            return res.json().then(function(data) {
                if (!res.ok) {
                    var err = new Error(data.error || ('HTTP ' + res.status));
                    err.status = res.status;
                    err.data = data;
                    throw err;
                }
                return data;
            });
        });
    }

    function pingServer(baseUrl, timeoutMs) {
        var url = normalizeServerUrl(baseUrl);
        if (!url) return Promise.resolve(null);
        var timeout = timeoutMs || SCAN_TIMEOUT_MS;
        var controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
        var timer = controller ? setTimeout(function() { controller.abort(); }, timeout) : null;
        return fetch(url + '/api/ping', {
            signal: controller ? controller.signal : undefined
        }).then(function(res) {
            if (timer) clearTimeout(timer);
            if (!res.ok) return null;
            return res.json().then(function(data) {
                if (!data.ok) return null;
                return {
                    url: url,
                    name: data.name,
                    ips: data.ips || [],
                    port: data.port || DEFAULT_PORT,
                    arenaCount: data.arenaCount || 0,
                    hasTournament: !!data.hasTournament
                };
            });
        }).catch(function() {
            if (timer) clearTimeout(timer);
            return null;
        });
    }

    function pingServerDetailed(baseUrl, timeoutMs) {
        var url = normalizeServerUrl(baseUrl);
        if (!url) {
            return Promise.resolve({ ok: false, reason: 'empty', message: 'Укажите адрес сервера' });
        }
        var timeout = timeoutMs || CONNECT_TIMEOUT_MS;
        var controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
        var timer = controller ? setTimeout(function() { controller.abort(); }, timeout) : null;
        return fetch(url + '/api/ping', {
            signal: controller ? controller.signal : undefined
        }).then(function(res) {
            if (timer) clearTimeout(timer);
            if (!res.ok) {
                return {
                    ok: false,
                    reason: 'http',
                    message: 'Сервер ответил с ошибкой HTTP ' + res.status + ' (' + url + ')'
                };
            }
            return res.json().then(function(data) {
                if (!data.ok) {
                    return { ok: false, reason: 'bad-response', message: 'Неверный ответ сервера' };
                }
                return {
                    ok: true,
                    info: {
                        url: url,
                        name: data.name,
                        ips: data.ips || [],
                        port: data.port || DEFAULT_PORT,
                        arenaCount: data.arenaCount || 0,
                        hasTournament: !!data.hasTournament
                    }
                };
            });
        }).catch(function(err) {
            if (timer) clearTimeout(timer);
            if (err && err.name === 'AbortError') {
                return {
                    ok: false,
                    reason: 'timeout',
                    message: 'Таймаут подключения к ' + url + '. Проверьте Wi‑Fi, IP и файрвол.'
                };
            }
            return {
                ok: false,
                reason: 'network',
                message: 'Не удалось достучаться до ' + url + '. Проверьте, что сервер запущен, устройства в одной сети, а адрес — не 127.0.0.1 с другого устройства.'
            };
        });
    }

    function guessLocalSubnets() {
        var subnets = [];
        var host = window.location.hostname;
        if (/^\d+\.\d+\.\d+\.\d+$/.test(host)) {
            subnets.push(host.replace(/\.\d+$/, ''));
        }
        subnets.push('192.168.1', '192.168.0', '10.0.0', '172.16.0');
        var unique = [];
        for (var i = 0; i < subnets.length; i++) {
            if (unique.indexOf(subnets[i]) === -1) unique.push(subnets[i]);
        }
        return unique;
    }

    function discoverServers(onProgress) {
        var found = {};
        var results = [];
        var subnets = guessLocalSubnets();
        var tasks = [];

        if (window.location.port === String(DEFAULT_PORT) || window.location.pathname.indexOf('secretary') !== -1) {
            tasks.push(pingServer(buildServerUrl(window.location.hostname, window.location.port || DEFAULT_PORT)));
        }

        for (var s = 0; s < subnets.length; s++) {
            for (var i = 1; i <= 254; i++) {
                tasks.push(pingServer(buildServerUrl(subnets[s] + '.' + i, DEFAULT_PORT)));
            }
        }

        var completed = 0;
        var total = tasks.length;

        function runBatch(start) {
            var batch = tasks.slice(start, start + SCAN_BATCH);
            if (!batch.length) {
                return Promise.resolve(results);
            }
            return Promise.all(batch.map(function(p) {
                return p.then(function(server) {
                    completed++;
                    if (onProgress) onProgress(completed, total);
                    if (server && !found[server.url]) {
                        found[server.url] = true;
                        results.push(server);
                    }
                });
            })).then(function() {
                return runBatch(start + SCAN_BATCH);
            });
        }

        return runBatch(0).then(function() {
            results.sort(function(a, b) { return a.url.localeCompare(b.url); });
            return results;
        });
    }

    function notifyConnection() {
        for (var i = 0; i < onConnectionCallbacks.length; i++) {
            onConnectionCallbacks[i](networkState.connected);
        }
    }

    function applyServerStateSnapshot(remoteState) {
        if (!remoteState) return;
        networkState.remoteVersion = remoteState.version || networkState.remoteVersion || 0;
        networkState.arenaCount = remoteState.arenaCount || 0;
        networkState.hostDeviceId = remoteState.hostDeviceId || null;
        networkState.lastRemoteState = remoteState;
    }

    function notifyState(remoteState) {
        applyServerStateSnapshot(remoteState);
        for (var i = 0; i < onStateCallbacks.length; i++) {
            onStateCallbacks[i](remoteState);
        }
    }

    function refreshServerState() {
        if (!networkState.serverUrl) {
            return Promise.reject(new Error('Сначала укажите адрес сервера и нажмите «Подключиться»'));
        }
        return fetchJson('/api/state').then(function(state) {
            notifyState(state);
            return state;
        });
    }

    function disconnectEventSource() {
        if (networkState.eventSource) {
            networkState.eventSource.close();
            networkState.eventSource = null;
        }
    }

    function connectSSE() {
        disconnectEventSource();
        if (!networkState.serverUrl) return;

        var es = new EventSource(apiUrl('/api/events'));
        networkState.eventSource = es;

        es.addEventListener('state', function(event) {
            try {
                notifyState(JSON.parse(event.data));
            } catch (e) {
                // ignore parse errors
            }
        });

        es.onerror = function() {
            networkState.connected = false;
            notifyConnection();
        };

        es.onopen = function() {
            networkState.connected = true;
            notifyConnection();
        };
    }

    function connect(serverInput) {
        var url = normalizeServerUrl(serverInput);
        if (!url && window.location.port === String(DEFAULT_PORT)) {
            url = buildServerUrl(window.location.hostname, window.location.port);
        }
        if (!url) {
            return Promise.reject(new Error('Укажите адрес сервера'));
        }

        networkState.serverUrl = url;
        networkState.deviceId = networkState.deviceId || generateDeviceId();
        if (!networkState.deviceName) {
            networkState.deviceName = getDefaultDeviceName();
        }

        return pingServerDetailed(url, CONNECT_TIMEOUT_MS).then(function(result) {
            if (!result.ok) {
                throw new Error(result.message || ('Сервер не отвечает: ' + url));
            }
            var info = result.info;
            networkState.enabled = true;
            networkState.connected = true;
            localStorage.setItem('meisterServerUrl', url);
            connectSSE();
            notifyConnection();
            return fetchJson('/api/state').then(function(state) {
                notifyState(state);
                return state;
            });
        });
    }

    function autoConnectIfSaved() {
        var saved = getStorageItem('serverUrl');
        if (saved) {
            return connect(saved).catch(function() { return null; });
        }
        if (window.location.port === String(DEFAULT_PORT)) {
            return connect(buildServerUrl(window.location.hostname, window.location.port)).catch(function() { return null; });
        }
        return Promise.resolve(null);
    }

    function registerHost(arenaCount) {
        var count = parseInt(arenaCount, 10);
        if (!count || count < 1) {
            count = 2;
        }
        return fetchJson('/api/register', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                deviceId: networkState.deviceId,
                role: 'host',
                name: networkState.deviceName
            })
        }).then(function(regData) {
            networkState.role = 'host';
            if (regData && regData.state) {
                applyServerStateSnapshot(regData.state);
            }
            return fetchJson('/api/arena-count', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    deviceId: networkState.deviceId,
                    arenaCount: count
                })
            });
        }).then(function(data) {
            networkState.arenaCount = count;
            networkState.role = 'host';
            if (data && data.state) {
                notifyState(data.state);
            }
            return data;
        });
    }

    /**
     * Регистрирует главное устройство на сервере, если ещё не зарегистрировано.
     * Нужно перед push турнира — хост часто пропускает шаг «Продолжить настройку турнира» в сети.
     */
    function ensureHostRegistered(arenaCount) {
        if (!isServerLinked()) {
            return Promise.reject(new Error('Сначала подключитесь к серверу (кнопка «Подключиться»).'));
        }
        if (networkState.role === 'arena') {
            return Promise.reject(new Error('Это устройство уже подключено как терминал площадки.'));
        }
        var count = parseInt(arenaCount, 10);
        if (!count || count < 1) {
            count = networkState.arenaCount > 0 ? networkState.arenaCount : 2;
        }
        return refreshServerState().then(function(state) {
            if (state.hostDeviceId && state.hostDeviceId !== networkState.deviceId) {
                return Promise.reject(new Error(
                    'На сервере уже зарегистрировано другое главное устройство. Перезапустите сервер турнира на хосте.'
                ));
            }
            if (state.hostDeviceId === networkState.deviceId && state.arenaCount > 0) {
                networkState.role = 'host';
                networkState.arenaCount = state.arenaCount;
                return state;
            }
            return registerHost(count).then(function() {
                return refreshServerState();
            });
        });
    }

    function registerArena(arenaId) {
        return fetchJson('/api/register', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                deviceId: networkState.deviceId,
                role: 'arena',
                arenaId: arenaId,
                name: networkState.deviceName || ('Площадка ' + arenaId)
            })
        }).then(function(data) {
            networkState.role = 'arena';
            networkState.arenaId = arenaId;
            if (data.state) notifyState(data.state);
            return data;
        });
    }

    function pushTournament(tournament) {
        if (networkState.role !== 'host') {
            return Promise.reject(new Error('Только главное устройство'));
        }
        return fetchJson('/api/tournament', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                deviceId: networkState.deviceId,
                tournament: tournament
            })
        }).then(function(data) {
            if (data && data.state) {
                networkState.remoteVersion = data.state.version || networkState.remoteVersion;
                notifyState(data.state);
            }
            return data;
        });
    }

    function claimMatch(matchType, refs, arenaId, options) {
        options = options || {};
        var body = {
            deviceId: networkState.deviceId,
            matchType: matchType,
            arenaId: arenaId || networkState.arenaId,
            reopen: !!options.reopen
        };
        if (matchType === 'pool') {
            body.poolId = refs.poolId;
            body.matchId = refs.matchId;
            networkState.activeMatchKey = 'pool:' + refs.poolId + ':' + refs.matchId;
        } else {
            body.bracketLoc = refs.bracketLoc;
            var loc = refs.bracketLoc;
            networkState.activeMatchKey = 'bracket:' + loc.side + ':' + loc.roundIndex + ':' + loc.matchIndex;
        }
        return fetchJson('/api/match/claim', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        }).then(function(data) {
            if (data && data.state) {
                networkState.remoteVersion = data.state.version || networkState.remoteVersion;
                notifyState(data.state);
            }
            return data;
        });
    }

    function completeMatch(tournament) {
        return fetchJson('/api/match/complete', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                deviceId: networkState.deviceId,
                matchKey: networkState.activeMatchKey,
                tournament: tournament
            })
        }).then(function(data) {
            networkState.activeMatchKey = null;
            if (data && data.state) {
                networkState.remoteVersion = data.state.version || networkState.remoteVersion;
                notifyState(data.state);
            }
            return data;
        });
    }

    function joinMatch(matchKey) {
        networkState.activeMatchKey = matchKey;
        return fetchJson('/api/match/join', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                deviceId: networkState.deviceId,
                matchKey: matchKey
            })
        }).then(function(data) {
            if (data && data.state) {
                networkState.remoteVersion = data.state.version || networkState.remoteVersion;
                notifyState(data.state);
            }
            return data;
        });
    }

    function releaseMatch(matchKey) {
        return fetchJson('/api/match/release', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                deviceId: networkState.deviceId,
                matchKey: matchKey || networkState.activeMatchKey
            })
        }).then(function() {
            networkState.activeMatchKey = null;
        });
    }

    function getTournamentSnapshot(gs) {
        return JSON.parse(JSON.stringify({
            sessionMode: gs.sessionMode,
            ruleset: gs.ruleset,
            tournamentSystem: gs.tournamentSystem,
            participants: gs.participants,
            tournamentStage: gs.tournamentStage,
            pools: gs.pools,
            poolMatches: gs.poolMatches,
            bracket: gs.bracket,
            playoffStarted: gs.playoffStarted,
            qualifyingAdvancersCount: gs.qualifyingAdvancersCount,
            tournamentFightHistory: gs.tournamentFightHistory,
            hostSelectedArenaId: gs.hostSelectedArenaId || null,
            poolArenaAssignments: gs.poolArenaAssignments || {},
            bracketArenaAssignments: gs.bracketArenaAssignments || {}
        }));
    }

    /** Сервер найден и connect() прошёл (роль ещё не назначена). */
    function isServerLinked() {
        return !!networkState.enabled && !!networkState.serverUrl;
    }

    /** Участие в турнире: связь с сервером + роль host/arena. */
    function isNetworkEnabled() {
        return isServerLinked() && !!networkState.role;
    }

    function isServerReachable() {
        return !!networkState.enabled && !!networkState.connected;
    }

    function isHost() {
        return isNetworkEnabled() && networkState.role === 'host';
    }

    function isArena() {
        return isNetworkEnabled() && networkState.role === 'arena';
    }

    function getState() {
        return networkState;
    }

    function onStateUpdate(callback) {
        onStateCallbacks.push(callback);
    }

    function onConnectionChange(callback) {
        onConnectionCallbacks.push(callback);
    }

    global.NetworkSync = {
        DEFAULT_PORT: DEFAULT_PORT,
        connect: connect,
        refreshServerState: refreshServerState,
        autoConnectIfSaved: autoConnectIfSaved,
        discoverServers: discoverServers,
        pingServer: pingServer,
        pingServerDetailed: pingServerDetailed,
        registerHost: registerHost,
        ensureHostRegistered: ensureHostRegistered,
        registerArena: registerArena,
        pushTournament: pushTournament,
        claimMatch: claimMatch,
        joinMatch: joinMatch,
        completeMatch: completeMatch,
        releaseMatch: releaseMatch,
        getTournamentSnapshot: getTournamentSnapshot,
        setDeviceName: setDeviceName,
        isNetworkEnabled: isNetworkEnabled,
        isServerLinked: isServerLinked,
        isServerReachable: isServerReachable,
        isHost: isHost,
        isArena: isArena,
        getState: getState,
        onStateUpdate: onStateUpdate,
        onConnectionChange: onConnectionChange,
        normalizeServerUrl: normalizeServerUrl,
        buildServerUrl: buildServerUrl
    };
})(window);
