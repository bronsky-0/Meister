/* Плей-офф: пулы, бои этапа пулов, сетка на выбывание */
(function(global) {
    'use strict';

    var draggedParticipantId = null;
    var tournamentSyncPending = false;

    function isPlayoffTournament() {
        return gameState.sessionMode === 'tournament' &&
            gameState.tournamentSystem === 'playoff';
    }

    function getParticipantById(id) {
        for (var i = 0; i < gameState.participants.length; i++) {
            if (gameState.participants[i].id === id) return gameState.participants[i];
        }
        return null;
    }

    function getParticipantName(id) {
        var p = getParticipantById(id);
        return p ? p.name : '—';
    }

    function getPoolById(poolId) {
        for (var i = 0; i < gameState.pools.length; i++) {
            if (gameState.pools[i].id === poolId) return gameState.pools[i];
        }
        return null;
    }

    function getUnassignedParticipantIds() {
        var assigned = {};
        for (var i = 0; i < gameState.pools.length; i++) {
            var pool = gameState.pools[i];
            for (var j = 0; j < pool.participantIds.length; j++) {
                assigned[pool.participantIds[j]] = true;
            }
        }
        return gameState.participants
            .filter(function(p) { return !assigned[p.id]; })
            .map(function(p) { return p.id; });
    }

    function generateParticipantId() {
        return 'p_' + Date.now() + '_' + Math.floor(Math.random() * 10000);
    }

    function parseParticipantNamesFromRows(rows) {
        var nameCol = 0;
        var startRow = 0;

        if (rows.length > 0 && rows[0] && rows[0].length) {
            var header = rows[0];
            for (var c = 0; c < header.length; c++) {
                var h = String(header[c] || '').trim().toLowerCase();
                if (h.indexOf('участник') >= 0 || h === 'participant' ||
                    h.indexOf('имя') >= 0 || h === 'name' || h.indexOf('фио') >= 0) {
                    nameCol = c;
                    startRow = 1;
                    break;
                }
            }
            if (startRow === 0 && header.length >= 2) {
                var firstCell = String(header[0] || '').trim();
                var secondCell = String(header[1] || '').trim();
                if ((/^\d+$/.test(firstCell) || typeof header[0] === 'number') && secondCell) {
                    nameCol = 1;
                }
            }
        }

        var names = [];
        var seen = {};
        for (var r = startRow; r < rows.length; r++) {
            var row = rows[r];
            if (!row || !row.length) continue;

            var raw = row[nameCol];
            if (raw === undefined || raw === null) continue;

            var name = String(raw).trim();
            if (!name) continue;

            var lower = name.toLowerCase();
            if (lower === '№' || lower === '#' || lower === 'участник' || lower === 'participant') {
                continue;
            }

            if (name.length > 40) name = name.slice(0, 40);

            var key = lower;
            if (seen[key]) continue;
            seen[key] = true;
            names.push(name);
        }
        return names;
    }

    function applyImportedParticipants(names) {
        if (gameState.participants.length > 0) {
            var msg = 'Текущий список будет заменён (' + names.length + ' участников). Продолжить?';
            if (!confirm(msg)) return;
        }

        gameState.pools = [];
        gameState.poolMatches = {};

        gameState.participants = names.map(function(name) {
            return { id: generateParticipantId(), name: name };
        });

        if (typeof renderParticipantsList === 'function') {
            renderParticipantsList();
        }
    }

    function importParticipantsFromWorkbook(wb) {
        if (!wb.SheetNames || wb.SheetNames.length === 0) {
            alert('В файле нет листов.');
            return;
        }

        var sheet = wb.Sheets[wb.SheetNames[0]];
        var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        var names = parseParticipantNamesFromRows(rows);

        if (names.length === 0) {
            alert('Не найдено имён участников. Ожидается столбец «Участник» или имена в первом столбце.');
            return;
        }

        applyImportedParticipants(names);
    }

    function handleParticipantsXlsxSelected(event) {
        var input = event.target;
        var file = input.files && input.files[0];
        if (!file) return;

        if (typeof XLSX === 'undefined') {
            alert('Библиотека Excel не загружена. Проверьте подключение к интернету.');
            input.value = '';
            return;
        }

        var reader = new FileReader();
        reader.onload = function(e) {
            try {
                var data = new Uint8Array(e.target.result);
                var wb = XLSX.read(data, { type: 'array' });
                importParticipantsFromWorkbook(wb);
            } catch (err) {
                alert('Не удалось прочитать файл: ' + (err.message || 'неизвестная ошибка'));
            }
            input.value = '';
        };
        reader.onerror = function() {
            alert('Не удалось прочитать файл.');
            input.value = '';
        };
        reader.readAsArrayBuffer(file);
    }

    function triggerImportParticipantsXlsx() {
        var input = document.getElementById('participantsXlsxInput');
        if (!input) return;
        input.value = '';
        input.click();
    }

    function goToPoolsComposition() {
        if (!isPlayoffTournament()) {
            alert('Составление пулов доступно только для системы «Плей-офф».');
            return;
        }
        if (gameState.participants.length < 2) {
            alert('Добавьте минимум двух участников.');
            return;
        }
        document.getElementById('startMenuOverlay').style.display = 'none';
        document.getElementById('poolsOverlay').style.display = 'flex';
        gameState.tournamentStage = 'pools';
        if (gameState.pools.length === 0) {
            addPool();
        }
        var defaultCount = Math.min(BRACKET_MAX_SIZE, gameState.participants.length);
        if (defaultCount >= 2 && !readQualifyingAdvancersCount()) {
            syncQualifyingAdvancersInputs(defaultCount);
        }
        renderPoolsComposition();
        if (typeof global.persistActiveTournamentNominationIfAny === 'function') {
            global.persistActiveTournamentNominationIfAny();
        }
    }

    function backFromPoolsToParticipants() {
        document.getElementById('poolsOverlay').style.display = 'none';
        document.getElementById('startMenuOverlay').style.display = 'flex';
        if (typeof showParticipantsPanel === 'function') {
            showParticipantsPanel();
        }
    }

    function generatePoolId() {
        return 'pool_' + Date.now() + '_' + Math.floor(Math.random() * 10000);
    }

    function addPool() {
        var number = gameState.pools.length + 1;
        gameState.pools.push({
            id: generatePoolId(),
            number: number,
            participantIds: []
        });
        renderPoolsComposition();
    }

    function removePool(poolId) {
        gameState.pools = gameState.pools.filter(function(p) { return p.id !== poolId; });
        for (var i = 0; i < gameState.pools.length; i++) {
            gameState.pools[i].number = i + 1;
        }
        delete gameState.poolMatches[poolId];
        renderPoolsComposition();
    }

    function assignParticipantToPool(participantId, poolId) {
        removeParticipantFromAllPools(participantId);
        var pool = getPoolById(poolId);
        if (pool && pool.participantIds.indexOf(participantId) === -1) {
            pool.participantIds.push(participantId);
        }
        renderPoolsComposition();
    }

    function removeParticipantFromAllPools(participantId) {
        for (var i = 0; i < gameState.pools.length; i++) {
            gameState.pools[i].participantIds = gameState.pools[i].participantIds.filter(function(id) {
                return id !== participantId;
            });
        }
    }

    function removeParticipantFromPool(participantId, poolId) {
        var pool = getPoolById(poolId);
        if (pool) {
            pool.participantIds = pool.participantIds.filter(function(id) { return id !== participantId; });
        }
        renderPoolsComposition();
    }

    function renderPoolsComposition() {
        renderUnassignedList();
        renderPoolsArea();
    }

    function renderUnassignedList() {
        var list = document.getElementById('unassignedParticipantsList');
        list.innerHTML = '';
        var ids = getUnassignedParticipantIds();

        if (ids.length === 0) {
            var empty = document.createElement('li');
            empty.className = 'participants-empty';
            empty.textContent = 'Все участники распределены по пулам';
            list.appendChild(empty);
            return;
        }

        for (var i = 0; i < ids.length; i++) {
            list.appendChild(createDraggableParticipantItem(ids[i], true));
        }
    }

    function createDraggableParticipantItem(participantId, fromUnassigned) {
        var item = document.createElement('li');
        item.className = 'participant-item pool-draggable';
        item.draggable = true;
        item.dataset.participantId = participantId;

        item.innerHTML =
            '<span class="participant-drag-handle">☰</span>' +
            '<span class="participant-name"></span>';

        item.querySelector('.participant-name').textContent = getParticipantName(participantId);

        item.addEventListener('dragstart', function(e) {
            draggedParticipantId = participantId;
            e.dataTransfer.setData('text/participant-id', participantId);
            e.dataTransfer.effectAllowed = 'move';
            item.classList.add('dragging');
        });
        item.addEventListener('dragend', function() {
            item.classList.remove('dragging');
            draggedParticipantId = null;
            clearDragOverStates();
        });

        return item;
    }

    function clearDragOverStates() {
        var els = document.querySelectorAll('.drag-over');
        for (var i = 0; i < els.length; i++) els[i].classList.remove('drag-over');
    }

    function renderPoolsArea() {
        var container = document.getElementById('poolsContainer');
        container.innerHTML = '';

        if (gameState.pools.length === 0) {
            container.innerHTML = '<p class="pools-empty">Нажмите «Добавить пул»</p>';
            return;
        }

        for (var i = 0; i < gameState.pools.length; i++) {
            container.appendChild(createPoolCard(gameState.pools[i]));
        }
    }

    function createPoolCard(pool) {
        var card = document.createElement('div');
        card.className = 'pool-card';
        card.dataset.poolId = pool.id;

        var header = document.createElement('div');
        header.className = 'pool-header';
        header.textContent = pool.number + ' пул';

        var table = document.createElement('table');
        table.className = 'pool-table';
        var tbody = document.createElement('tbody');
        tbody.className = 'pool-drop-zone';
        tbody.dataset.poolId = pool.id;

        tbody.addEventListener('dragover', onPoolDragOver);
        tbody.addEventListener('dragleave', onPoolDragLeave);
        tbody.addEventListener('drop', onPoolDrop);

        if (pool.participantIds.length === 0) {
            var emptyRow = document.createElement('tr');
            emptyRow.className = 'pool-empty-row';
            var td = document.createElement('td');
            td.colSpan = 2;
            td.textContent = 'Перетащите участников сюда';
            emptyRow.appendChild(td);
            tbody.appendChild(emptyRow);
        } else {
            for (var j = 0; j < pool.participantIds.length; j++) {
                tbody.appendChild(createPoolParticipantRow(pool.id, pool.participantIds[j]));
            }
        }

        table.appendChild(tbody);

        var removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'pool-remove-btn';
        removeBtn.textContent = 'Удалить пул';
        removeBtn.onclick = function() {
            if (confirm('Удалить пул ' + pool.number + '? Участники вернутся в общий список.')) {
                removePool(pool.id);
            }
        };

        card.appendChild(header);
        card.appendChild(table);
        card.appendChild(removeBtn);
        return card;
    }

    function createPoolParticipantRow(poolId, participantId) {
        var row = document.createElement('tr');
        row.className = 'pool-participant-row';
        row.draggable = true;
        row.dataset.participantId = participantId;
        row.dataset.poolId = poolId;

        var numTd = document.createElement('td');
        numTd.className = 'pool-row-num';
        numTd.textContent = '☰';

        var nameTd = document.createElement('td');
        nameTd.textContent = getParticipantName(participantId);

        row.appendChild(numTd);
        row.appendChild(nameTd);

        row.addEventListener('dragstart', function(e) {
            draggedParticipantId = participantId;
            e.dataTransfer.setData('text/participant-id', participantId);
            e.dataTransfer.setData('text/from-pool', poolId);
            e.dataTransfer.effectAllowed = 'move';
            row.classList.add('dragging');
        });
        row.addEventListener('dragend', function() {
            row.classList.remove('dragging');
            draggedParticipantId = null;
            clearDragOverStates();
        });

        return row;
    }

    function onPoolDragOver(e) {
        e.preventDefault();
        e.currentTarget.closest('.pool-card').classList.add('drag-over');
    }

    function onPoolDragLeave(e) {
        var card = e.currentTarget.closest('.pool-card');
        if (card) card.classList.remove('drag-over');
    }

    function onPoolDrop(e) {
        e.preventDefault();
        var poolId = e.currentTarget.dataset.poolId;
        var card = e.currentTarget.closest('.pool-card');
        if (card) card.classList.remove('drag-over');

        var participantId = e.dataTransfer.getData('text/participant-id') || draggedParticipantId;
        if (participantId && poolId) {
            assignParticipantToPool(participantId, poolId);
        }
    }

    function setupUnassignedDropZone() {
        var zone = document.getElementById('unassignedDropZone');
        if (!zone || zone.dataset.bound) return;
        zone.dataset.bound = '1';

        zone.addEventListener('dragover', function(e) {
            e.preventDefault();
            zone.classList.add('drag-over');
        });
        zone.addEventListener('dragleave', function() {
            zone.classList.remove('drag-over');
        });
        zone.addEventListener('drop', function(e) {
            e.preventDefault();
            zone.classList.remove('drag-over');
            var participantId = e.dataTransfer.getData('text/participant-id') || draggedParticipantId;
            if (participantId) {
                removeParticipantFromAllPools(participantId);
                renderPoolsComposition();
            }
        });
    }

    function shuffleArray(items) {
        var arr = items.slice();
        for (var i = arr.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var tmp = arr[i];
            arr[i] = arr[j];
            arr[j] = tmp;
        }
        return arr;
    }

    function randomInt(min, max) {
        return min + Math.floor(Math.random() * (max - min + 1));
    }

    function randomizePoolsDistribution() {
        if (!isPlayoffTournament()) {
            alert('Доступно только для турнира «Плей-офф».');
            return;
        }
        if (gameState.participants.length < 2) {
            alert('Добавьте минимум двух участников.');
            return;
        }
        if (gameState.pools.length === 0) {
            alert('Сначала создайте хотя бы один пул.');
            return;
        }
        if (!confirm('Случайно распределить всех участников по пулам? Текущее распределение будет сброшено.')) {
            return;
        }

        gameState.poolMatches = {};
        for (var p = 0; p < gameState.pools.length; p++) {
            gameState.pools[p].participantIds = [];
        }

        var shuffledIds = shuffleArray(gameState.participants.map(function(participant) {
            return participant.id;
        }));

        for (var i = 0; i < shuffledIds.length; i++) {
            var pool = gameState.pools[i % gameState.pools.length];
            pool.participantIds.push(shuffledIds[i]);
        }

        renderPoolsComposition();

        var weakPools = [];
        for (var k = 0; k < gameState.pools.length; k++) {
            if (gameState.pools[k].participantIds.length < 2) {
                weakPools.push(gameState.pools[k].number);
            }
        }
        if (weakPools.length > 0) {
            alert(
                'Участники распределены. В пулах ' + weakPools.join(', ') +
                ' меньше 2 человек — добавьте пул или перераспределите вручную.'
            );
        }
    }

    function stripPoolFightHistory(poolId) {
        gameState.tournamentFightHistory = gameState.tournamentFightHistory.filter(function(entry) {
            return !(entry.type === 'pool' && entry.poolId === poolId);
        });
    }

    function completePoolMatchWithRandomResult(match) {
        var pickRed = Math.random() < 0.5;
        var winnerId = pickRed ? match.fighter1Id : match.fighter2Id;
        var winnerSide = pickRed ? 'red' : 'blue';
        var redScore = pickRed ? randomInt(8, 15) : randomInt(0, 7);
        var blueScore = pickRed ? randomInt(0, 7) : randomInt(8, 15);

        match.status = 'done';
        match.winnerId = winnerId;
        match.result = {
            redScore: redScore,
            blueScore: blueScore,
            redBonus: 0,
            blueBonus: 0,
            redRoundsWon: pickRed ? 1 : 0,
            blueRoundsWon: pickRed ? 0 : 1,
            winnerSide: winnerSide,
            date: new Date().toLocaleString('ru-RU'),
            randomTest: true
        };

        gameState.tournamentFightHistory.push({
            type: 'pool',
            poolId: match.poolId,
            matchId: match.id,
            fighter1: getParticipantName(match.fighter1Id),
            fighter2: getParticipantName(match.fighter2Id),
            winnerId: winnerId,
            result: match.result
        });
    }

    function applyRandomResultsToPool(poolId, skipConfirm) {
        var pool = getPoolById(poolId);
        if (!pool || pool.participantIds.length < 2) return false;

        if (!gameState.poolMatches[poolId] || gameState.poolMatches[poolId].length === 0) {
            generatePoolMatches(poolId);
        }

        var matches = getPoolMatches(poolId);
        var hasDone = matches.some(function(m) { return m.status === 'done'; });

        if (!skipConfirm && hasDone && !confirm(
            'Пересчитать все бои пула «' + pool.number + '» случайными результатами?'
        )) {
            return false;
        }

        stripPoolFightHistory(poolId);

        for (var i = 0; i < matches.length; i++) {
            completePoolMatchWithRandomResult(matches[i]);
        }

        if (gameState.activePoolId === poolId) {
            gameState.activePoolId = null;
            gameState.activePoolMatchId = null;
            gameState.activePoolMatchMeta = null;
        }

        return true;
    }

    function randomizePoolFightResults(poolId) {
        if (!isPlayoffTournament()) {
            alert('Доступно только для турнира «Плей-офф».');
            return;
        }

        var pool = getPoolById(poolId);
        if (!pool) return;

        if (pool.participantIds.length < 2) {
            alert('В пуле «' + pool.number + '» меньше 2 участников.');
            return;
        }

        if (!applyRandomResultsToPool(poolId, false)) return;

        updateTournamentBar();

        if (document.getElementById('poolSelectOverlay').style.display === 'flex') {
            showPoolSelectModal();
        }
    }

    function randomizeAllPoolFightResults() {
        if (!isPlayoffTournament()) {
            alert('Доступно только для турнира «Плей-офф».');
            return;
        }
        if (gameState.pools.length === 0) {
            alert('Нет пулов.');
            return;
        }

        if (!confirm('Случайно провести все бои во всех пулах?')) {
            return;
        }

        var applied = 0;
        for (var i = 0; i < gameState.pools.length; i++) {
            if (applyRandomResultsToPool(gameState.pools[i].id, true)) {
                applied++;
            }
        }

        gameState.activePoolId = null;
        gameState.activePoolMatchId = null;
        gameState.activePoolMatchMeta = null;
        updateTournamentBar();

        if (document.getElementById('poolSelectOverlay').style.display === 'flex') {
            showPoolSelectModal();
        }

        if (applied === 0) {
            alert('Ни в одном пуле нет минимум 2 участников.');
            return;
        }

        if (allPoolsComplete()) {
            alert('Все бои пулов проведены (случайные результаты). Можно сформировать сетку.');
        }
    }

    function randomizeCurrentPoolFightResults() {
        if (gameState.activePoolId) {
            randomizePoolFightResults(gameState.activePoolId);
        } else {
            randomizeAllPoolFightResults();
        }
    }

    function validatePoolsForStart() {
        if (gameState.pools.length === 0) {
            alert('Создайте хотя бы один пул.');
            return false;
        }
        var unassigned = getUnassignedParticipantIds();
        if (unassigned.length > 0) {
            alert('Все участники должны быть распределены по пулам.');
            return false;
        }
        for (var i = 0; i < gameState.pools.length; i++) {
            if (gameState.pools[i].participantIds.length < 2) {
                alert('В пуле «' + gameState.pools[i].number + '» должно быть минимум 2 участника.');
                return false;
            }
        }
        return true;
    }

    function generatePoolMatches(poolId) {
        var pool = getPoolById(poolId);
        if (!pool) return [];

        var ids = pool.participantIds.slice();
        var matches = [];
        for (var i = 0; i < ids.length; i++) {
            for (var j = i + 1; j < ids.length; j++) {
                matches.push({
                    id: 'm_' + poolId + '_' + i + '_' + j,
                    poolId: poolId,
                    fighter1Id: ids[i],
                    fighter2Id: ids[j],
                    status: 'pending',
                    winnerId: null,
                    result: null
                });
            }
        }
        gameState.poolMatches[poolId] = matches;
        return matches;
    }

    function getPoolMatches(poolId) {
        return gameState.poolMatches[poolId] || [];
    }

    function isNetworkMode() {
        return typeof NetworkSync !== 'undefined' && NetworkSync.isNetworkEnabled();
    }

    function isNetworkArenaDevice() {
        return isNetworkMode() && NetworkSync.isArena();
    }

    function isNetworkHost() {
        return isNetworkMode() && NetworkSync.isHost();
    }

    function getLocalArenaId() {
        if (gameState.activeArenaId) return gameState.activeArenaId;
        if (isNetworkArenaDevice()) return NetworkSync.getState().arenaId;
        return null;
    }

    function canEditMatchOnArena(match, arenaId) {
        if (!match || !arenaId) return false;
        if (match.status === 'in_progress' && match.arenaId === arenaId) return true;
        if (match.status === 'done' && match.arenaId === arenaId) return true;
        return false;
    }

    function isMatchAvailableForArena(match, arenaId) {
        if (!match) return false;
        if (match.status === 'pending') return true;
        if (match.status === 'in_progress' && arenaId && match.arenaId === arenaId) return true;
        return false;
    }

    function isMatchSelectableForArena(match, arenaId) {
        if (isNetworkArenaDevice()) {
            return match && match.status === 'in_progress' && match.arenaId === arenaId;
        }
        return isMatchAvailableForArena(match, arenaId) || canEditMatchOnArena(match, arenaId);
    }

    function isMatchLockedByHostForArena(matchKey, arenaId) {
        if (!matchKey || !arenaId || typeof NetworkSync === 'undefined') return false;
        var remote = NetworkSync.getState().lastRemoteState;
        if (!remote || !remote.matchLocks || !remote.hostDeviceId) return false;
        var lock = remote.matchLocks[matchKey];
        return !!(lock &&
            lock.deviceId === remote.hostDeviceId &&
            lock.arenaId === arenaId);
    }

    function parseMatchKey(matchKey) {
        if (!matchKey) return null;
        var parts = matchKey.split(':');
        if (parts[0] === 'pool' && parts.length >= 3) {
            return {
                type: 'pool',
                poolId: parts[1],
                matchId: parts.slice(2).join(':')
            };
        }
        if (parts[0] === 'bracket' && parts.length >= 4) {
            return {
                type: 'bracket',
                bracketLoc: {
                    side: parts[1],
                    roundIndex: parseInt(parts[2], 10),
                    matchIndex: parseInt(parts[3], 10)
                }
            };
        }
        return null;
    }

    function findHostAssignedFightForArena(remoteState, arenaId) {
        if (!remoteState || !remoteState.matchLocks || !remoteState.hostDeviceId || !arenaId) {
            return null;
        }
        var locks = remoteState.matchLocks;
        var keys = Object.keys(locks);
        for (var i = 0; i < keys.length; i++) {
            var matchKey = keys[i];
            var lock = locks[matchKey];
            if (!lock || lock.deviceId !== remoteState.hostDeviceId) continue;
            if (lock.arenaId !== arenaId) continue;
            var parsed = parseMatchKey(matchKey);
            if (!parsed) continue;
            return { matchKey: matchKey, lock: lock, parsed: parsed };
        }
        return null;
    }

    function beginHostAssignedArenaFight(assignment) {
        if (!assignment || !assignment.parsed) return;
        var parsed = assignment.parsed;
        if (parsed.type === 'pool') {
            var poolMatch = findPoolMatch(parsed.matchId);
            if (!poolMatch || poolMatch.status !== 'in_progress') return;
            gameState.activePoolId = parsed.poolId;
            gameState.activePoolMatchId = poolMatch.id;
            gameState.tournamentStage = 'pool-fights';
            loadPoolMatchIntoFight(poolMatch);
        } else {
            var loc = parsed.bracketLoc;
            var bracketMatch = getBracketMatch(loc);
            if (!bracketMatch || bracketMatch.status !== 'in_progress') return;
            gameState.activeBracketMatch = loc;
            gameState.tournamentStage = 'bracket-fights';
            loadBracketMatchIntoFight(bracketMatch, loc);
        }
        gameState.sessionStarted = true;
        gameState.tournamentMode = true;
        document.getElementById('secretaryTerminal').classList.remove('hidden');
        var waitEl = document.getElementById('arenaWaitOverlay');
        if (waitEl) waitEl.style.display = 'none';
        document.getElementById('poolSelectOverlay').style.display = 'none';
        hideArenaStageCompleteOverlay();
        updateTournamentBar();
        updatePlayoffTerminalButtons();
        if (typeof updateDisplay === 'function') updateDisplay();
    }

    var arenaAutoJoinInFlight = false;
    var hostDispatchInFlight = false;

    function isPoolStageActive() {
        return gameState.playoffStarted &&
            (!gameState.bracket ||
                gameState.tournamentStage === 'pool-fights' ||
                gameState.tournamentStage === 'pools');
    }

    function isBracketStageActive() {
        return !!gameState.bracket &&
            (gameState.tournamentStage === 'bracket' ||
                gameState.tournamentStage === 'bracket-fights');
    }

    function arenaHasInProgressMatch(arenaId) {
        if (!arenaId) return false;
        var poolIds = gameState.poolMatches || {};
        for (var poolId in poolIds) {
            if (!poolIds.hasOwnProperty(poolId)) continue;
            var matches = poolIds[poolId];
            for (var i = 0; i < matches.length; i++) {
                var m = matches[i];
                if (m.status === 'in_progress' && m.arenaId === arenaId) return true;
            }
        }
        if (!gameState.bracket) return false;
        function checkMatch(match) {
            return match && match.status === 'in_progress' && match.arenaId === arenaId;
        }
        var b = gameState.bracket;
        if (checkMatch(b.final) || checkMatch(b.thirdPlace)) return true;
        for (var si = 0; si < 2; si++) {
            var side = si === 0 ? 'left' : 'right';
            var half = b[side];
            if (!half || !half.rounds) continue;
            for (var r = 0; r < half.rounds.length; r++) {
                for (var mi = 0; mi < half.rounds[r].length; mi++) {
                    if (checkMatch(half.rounds[r][mi])) return true;
                }
            }
        }
        return false;
    }

    function getNextPendingPoolDispatchItem(arenaId) {
        for (var i = 0; i < gameState.pools.length; i++) {
            var poolId = gameState.pools[i].id;
            if (getPoolAssignedArena(poolId) !== arenaId) continue;
            var matches = getPoolMatches(poolId);
            for (var m = 0; m < matches.length; m++) {
                var match = matches[m];
                if (match.status === 'pending') {
                    return {
                        type: 'pool',
                        poolId: poolId,
                        matchId: match.id,
                        match: match
                    };
                }
            }
        }
        return null;
    }

    function getNextBracketDispatchItem(arenaId) {
        if (!gameState.bracket) return null;
        var items = collectAvailableBracketMatchesForHost(arenaId);
        return items.length ? items[0] : null;
    }

    function collectAvailableBracketMatchesForHost(arenaId) {
        var items = [];
        if (!gameState.bracket) return items;

        function tryAdd(match, loc) {
            if (!match || match.status === 'done') return;
            if (match.status === 'in_progress') return;
            if (!canFightBracketMatch(match, loc)) return;
            var f1 = getBracketDisplayFighter(match.fighter1, loc, match) || match.fighter1;
            var f2 = getBracketDisplayFighter(match.fighter2, loc, match) || match.fighter2;
            if (!f1 || !f2) return;
            items.push({
                type: 'bracket',
                bracketLoc: {
                    side: loc.side,
                    roundIndex: loc.roundIndex,
                    matchIndex: loc.matchIndex
                },
                match: match
            });
        }

        ['left', 'right'].forEach(function(side) {
            var half = gameState.bracket[side];
            if (!half || !half.rounds) return;
            for (var r = 0; r < half.rounds.length; r++) {
                for (var m = 0; m < half.rounds[r].length; m++) {
                    tryAdd(half.rounds[r][m], { side: side, roundIndex: r, matchIndex: m });
                }
            }
        });
        tryAdd(gameState.bracket.thirdPlace, { side: 'bronze', roundIndex: 0, matchIndex: 0 });
        tryAdd(gameState.bracket.final, { side: 'final', roundIndex: 0, matchIndex: 0 });
        return items;
    }

    function getNextDispatchItemForArena(arenaId) {
        if (isPoolStageActive()) {
            return getNextPendingPoolDispatchItem(arenaId);
        }
        if (isBracketStageActive()) {
            return getNextBracketDispatchItem(arenaId);
        }
        return null;
    }

    function areArenaPoolFightsComplete(arenaId) {
        var hasAssigned = false;
        for (var i = 0; i < gameState.pools.length; i++) {
            var poolId = gameState.pools[i].id;
            if (getPoolAssignedArena(poolId) !== arenaId) continue;
            hasAssigned = true;
            var matches = getPoolMatches(poolId);
            for (var m = 0; m < matches.length; m++) {
                if (matches[m].status !== 'done') return false;
            }
        }
        return hasAssigned;
    }

    function arenaHadBracketFights(arenaId) {
        if (!gameState.bracket) return false;
        function check(match) {
            return match && (match.status === 'done' || match.status === 'in_progress') &&
                match.arenaId === arenaId;
        }
        var b = gameState.bracket;
        if (check(b.final) || check(b.thirdPlace)) return true;
        for (var si = 0; si < 2; si++) {
            var side = si === 0 ? 'left' : 'right';
            var half = b[side];
            if (!half || !half.rounds) continue;
            for (var r = 0; r < half.rounds.length; r++) {
                for (var mi = 0; mi < half.rounds[r].length; mi++) {
                    if (check(half.rounds[r][mi])) return true;
                }
            }
        }
        return false;
    }

    function areArenaBracketFightsComplete(arenaId) {
        if (!isBracketStageActive() || !arenaId) return false;
        if (arenaHasInProgressMatch(arenaId)) return false;
        var remote = typeof NetworkSync !== 'undefined' ? NetworkSync.getState().lastRemoteState : null;
        if (findHostAssignedFightForArena(remote, arenaId)) return false;
        if (getNextBracketDispatchItem(arenaId)) return false;
        return arenaHadBracketFights(arenaId);
    }

    function isArenaStageCompleteForTerminal() {
        var arenaId = getLocalArenaId();
        if (!arenaId) return false;
        if (isPoolStageActive()) return areArenaPoolFightsComplete(arenaId);
        if (isBracketStageActive()) return areArenaBracketFightsComplete(arenaId);
        return false;
    }

    function showArenaStageCompleteOverlay() {
        if (!isNetworkArenaDevice()) return;
        hideArenaWaitOverlay();
        var overlay = document.getElementById('arenaStageCompleteOverlay');
        var term = document.getElementById('secretaryTerminal');
        if (term) term.classList.add('hidden');
        if (overlay) overlay.style.display = 'flex';
    }

    function hideArenaStageCompleteOverlay() {
        var overlay = document.getElementById('arenaStageCompleteOverlay');
        if (overlay) overlay.style.display = 'none';
    }

    function hideArenaWaitOverlay() {
        var waitEl = document.getElementById('arenaWaitOverlay');
        if (waitEl) waitEl.style.display = 'none';
    }

    function refreshArenaTerminalUi() {
        if (!isNetworkArenaDevice()) return;
        if (gameState.activePoolMatchId || gameState.activeBracketMatch) return;

        if (isArenaStageCompleteForTerminal()) {
            showArenaStageCompleteOverlay();
            return;
        }

        hideArenaStageCompleteOverlay();
        var remote = typeof NetworkSync !== 'undefined' ? NetworkSync.getState().lastRemoteState : null;
        updateArenaWaitOverlay(remote, findHostAssignedFightForArena(remote, getLocalArenaId()));
        tryAutoJoinHostAssignedFight(remote);
    }

    function hostDispatchClaim(item, arenaId, options) {
        if (!item || !arenaId) return Promise.resolve();
        if (item.type === 'pool') {
            return NetworkSync.claimMatch('pool', {
                poolId: item.poolId,
                matchId: item.matchId
            }, arenaId, options || {});
        }
        return NetworkSync.claimMatch('bracket', {
            bracketLoc: item.bracketLoc
        }, arenaId, options || {});
    }

    function dispatchNextFightForArena(arenaId) {
        if (!isNetworkHost() || !arenaId || hostDispatchInFlight) {
            return Promise.resolve();
        }
        if (gameState.activePoolMatchId || gameState.activeBracketMatch) {
            return Promise.resolve();
        }

        var remote = typeof NetworkSync !== 'undefined' ? NetworkSync.getState().lastRemoteState : null;
        if (findHostAssignedFightForArena(remote, arenaId)) {
            return Promise.resolve();
        }
        if (arenaHasInProgressMatch(arenaId)) {
            return Promise.resolve();
        }

        if (isPoolStageActive() && areArenaPoolFightsComplete(arenaId)) {
            return Promise.resolve();
        }
        if (isBracketStageActive() && areArenaBracketFightsComplete(arenaId)) {
            return Promise.resolve();
        }

        var item = getNextDispatchItemForArena(arenaId);
        if (!item) return Promise.resolve();

        hostDispatchInFlight = true;
        return hostDispatchClaim(item, arenaId).then(function(data) {
            if (data && data.state) {
                applyRemoteTournamentState(data.state);
            }
        }).catch(function(err) {
            console.warn('dispatchNextFightForArena', err);
        }).finally(function() {
            hostDispatchInFlight = false;
        });
    }

    function getArenaIdsForAutoDispatch() {
        var ids = {};
        if (isPoolStageActive()) {
            for (var i = 0; i < gameState.pools.length; i++) {
                var a = getPoolAssignedArena(gameState.pools[i].id);
                if (a) ids[a] = true;
            }
        }
        if (isBracketStageActive()) {
            var selected = readSelectedArenaId();
            if (selected) ids[selected] = true;
        }
        return Object.keys(ids).map(function(k) { return parseInt(k, 10); });
    }

    function dispatchForAllActiveArenas() {
        if (!isNetworkHost()) return Promise.resolve();
        var arenaIds = getArenaIdsForAutoDispatch();
        var chain = Promise.resolve();
        for (var i = 0; i < arenaIds.length; i++) {
            (function(aid) {
                chain = chain.then(function() { return dispatchNextFightForArena(aid); });
            })(arenaIds[i]);
        }
        return chain;
    }

    function tryAutoJoinHostAssignedFight(remoteState) {
        if (!isNetworkArenaDevice() || !remoteState || arenaAutoJoinInFlight) return;
        if (gameState.activePoolMatchId || gameState.activeBracketMatch) return;

        var arenaId = getLocalArenaId();
        if (!arenaId) return;

        if (isArenaStageCompleteForTerminal()) {
            showArenaStageCompleteOverlay();
            return;
        }

        var assignment = findHostAssignedFightForArena(remoteState, arenaId);
        if (!assignment) {
            updateArenaWaitOverlay(remoteState, null);
            return;
        }

        hideArenaStageCompleteOverlay();

        var match;
        if (assignment.parsed.type === 'pool') {
            match = findPoolMatch(assignment.parsed.matchId);
        } else {
            match = getBracketMatch(assignment.parsed.bracketLoc);
        }
        if (!match || match.status !== 'in_progress') {
            updateArenaWaitOverlay(remoteState, null);
            return;
        }

        updateArenaWaitOverlay(remoteState, assignment);
        arenaAutoJoinInFlight = true;
        NetworkSync.joinMatch(assignment.matchKey).then(function(data) {
            if (data && data.state) {
                applyRemoteTournamentState(data.state, { skipAutoJoin: true });
            }
            beginHostAssignedArenaFight(assignment);
        }).catch(function() {
            updateArenaWaitOverlay(remoteState, assignment);
        }).finally(function() {
            arenaAutoJoinInFlight = false;
        });
    }

    function updateArenaWaitOverlay(remoteState, assignment) {
        if (!isNetworkArenaDevice()) return;
        var waitEl = document.getElementById('arenaWaitOverlay');
        var hintEl = document.getElementById('arenaWaitHint');
        var titleEl = document.getElementById('arenaWaitTitle');
        if (!waitEl || !hintEl) return;

        if (!remoteState || !remoteState.tournament || !remoteState.tournament.playoffStarted) {
            return;
        }

        if (gameState.activePoolMatchId || gameState.activeBracketMatch) {
            waitEl.style.display = 'none';
            hideArenaStageCompleteOverlay();
            return;
        }

        if (isArenaStageCompleteForTerminal()) {
            showArenaStageCompleteOverlay();
            return;
        }

        hideArenaStageCompleteOverlay();

        if (!assignment) {
            waitEl.style.display = 'flex';
            if (titleEl) {
                titleEl.textContent = 'Площадка ' + getLocalArenaId() + ' — ожидание';
            }
            var stageHint = isPoolStageActive()
                ? 'Ожидание назначения пула на эту площадку…'
                : 'Ожидание следующего боя сетки…';
            hintEl.textContent = stageHint;
            return;
        }

        waitEl.style.display = 'flex';
        if (titleEl) titleEl.textContent = 'Площадка ' + getLocalArenaId() + ' — бой назначен';
        var label = 'Подключение к бою…';
        if (assignment.parsed.type === 'pool') {
            var pm = findPoolMatch(assignment.parsed.matchId);
            if (pm) {
                label = getParticipantName(pm.fighter1Id) + ' vs ' +
                    getParticipantName(pm.fighter2Id);
            }
        } else {
            var bm = getBracketMatch(assignment.parsed.bracketLoc);
            if (bm) {
                var f1 = bm.fighter1 ? bm.fighter1.name : '—';
                var f2 = bm.fighter2 ? bm.fighter2.name : '—';
                label = f1 + ' vs ' + f2;
            }
        }
        hintEl.textContent = 'Главное устройство запустило: ' + label;
    }

    function syncTournamentToServer() {
        if (typeof global.persistActiveTournamentNominationIfAny === 'function') {
            global.persistActiveTournamentNominationIfAny();
        }
        if (!isNetworkHost()) return Promise.resolve();
        return NetworkSync.pushTournament(NetworkSync.getTournamentSnapshot(gameState)).catch(function(err) {
            alert('Ошибка синхронизации: ' + (err.message || err));
        });
    }

    function syncTournamentAfterFight() {
        if (typeof global.persistActiveTournamentNominationIfAny === 'function') {
            global.persistActiveTournamentNominationIfAny();
        }
        if (!isNetworkMode()) return Promise.resolve();
        tournamentSyncPending = true;
        return NetworkSync.completeMatch(NetworkSync.getTournamentSnapshot(gameState)).then(function(data) {
            if (data && data.state) {
                applyRemoteTournamentState(data.state);
            }
            if (isNetworkHost()) {
                return dispatchForAllActiveArenas();
            }
            if (isNetworkArenaDevice()) {
                refreshArenaTerminalUi();
            }
        }).catch(function(err) {
            alert('Ошибка синхронизации результата: ' + (err.message || err));
        }).finally(function() {
            tournamentSyncPending = false;
        });
    }

    function syncHostArenaSelection(arenaId) {
        if (!isNetworkHost()) return Promise.resolve();
        gameState.hostSelectedArenaId = arenaId;
        return syncTournamentToServer().then(function() {
            return dispatchNextFightForArena(arenaId);
        });
    }

    function populateArenaSelectElement(selectEl, selectedId) {
        if (!selectEl) return;
        var arenaCount = (typeof NetworkSync !== 'undefined' && NetworkSync.getState().arenaCount) || 1;
        var current = selectedId || readSelectedArenaId();
        selectEl.innerHTML = '';
        for (var a = 1; a <= arenaCount; a++) {
            var opt = document.createElement('option');
            opt.value = String(a);
            opt.textContent = 'Площадка ' + a;
            if (current === a) opt.selected = true;
            selectEl.appendChild(opt);
        }
    }

    function populateNetworkArenaSelects() {
        if (!isNetworkMode()) return;
        var selected = readSelectedArenaId();
        populateArenaSelectElement(document.getElementById('arenaSelectInput'), selected);
        populateArenaSelectElement(document.getElementById('tournamentBarArenaSelect'), selected);
        populateArenaSelectElement(document.getElementById('bracketArenaSelect'), selected);
    }

    function onAdminArenaChanged() {
        var arenaId = readSelectedArenaId();
        gameState.activeArenaId = arenaId;
        if (isNetworkHost()) {
            syncHostArenaSelection(arenaId);
        }
        populateNetworkArenaSelects();
        if (document.getElementById('poolSelectOverlay').style.display === 'flex' && !isNetworkArenaDevice()) {
            showArenaMatchSelectModal();
        }
        updateTournamentBar();
    }

    function applyRemoteTournamentState(remoteState, options) {
        options = options || {};
        if (!remoteState || !remoteState.tournament) return;

        if (typeof NetworkSync !== 'undefined') {
            var incomingVersion = remoteState.version || 0;
            var currentVersion = NetworkSync.getState().remoteVersion || 0;
            if (incomingVersion > 0 && incomingVersion < currentVersion) {
                return;
            }
            if (incomingVersion > currentVersion) {
                NetworkSync.getState().remoteVersion = incomingVersion;
            }
        }

        if (tournamentSyncPending) {
            return;
        }

        var t = remoteState.tournament;
        var inFight = !!(gameState.activePoolMatchId || gameState.activeBracketMatch);

        gameState.sessionMode = t.sessionMode || 'tournament';
        gameState.ruleset = t.ruleset;
        gameState.tournamentSystem = t.tournamentSystem;
        gameState.participants = t.participants || [];
        gameState.tournamentStage = t.tournamentStage;
        gameState.pools = t.pools || [];
        gameState.poolMatches = t.poolMatches || {};
        gameState.bracket = t.bracket;
        gameState.playoffStarted = !!t.playoffStarted;
        gameState.qualifyingAdvancersCount = t.qualifyingAdvancersCount;
        gameState.tournamentFightHistory = t.tournamentFightHistory || [];
        gameState.tournamentMode = !!t.playoffStarted;
        if (t.hostSelectedArenaId) {
            gameState.hostSelectedArenaId = t.hostSelectedArenaId;
        }
        gameState.poolArenaAssignments = t.poolArenaAssignments || {};

        if (t.playoffStarted && !inFight) {
            document.getElementById('startMenuOverlay').style.display = 'none';
            document.getElementById('poolsOverlay').style.display = 'none';
            if (isNetworkArenaDevice()) {
                var waitEl = document.getElementById('arenaWaitOverlay');
                if (waitEl) waitEl.style.display = 'none';
            }
            showSecretaryTerminal();
        }

        updateTournamentBar();
        updateFormBracketButtons();
        updateOpenBracketButton();
        updatePlayoffTerminalButtons();
        populateNetworkArenaSelects();

        if (document.getElementById('poolSelectOverlay').style.display === 'flex') {
            if (isNetworkArenaDevice()) {
                document.getElementById('poolSelectOverlay').style.display = 'none';
            } else if (isNetworkMode()) {
                showArenaMatchSelectModal();
            } else {
                showPoolSelectModal();
            }
        }

        if (gameState.bracket) {
            renderBracket(true);
        }

        enforceTournamentNavigation();

        if (!options.skipAutoJoin && isNetworkArenaDevice()) {
            tryAutoJoinHostAssignedFight(remoteState);
        }

        if (!options.skipAutoJoin && isNetworkHost()) {
            dispatchForAllActiveArenas();
        }

        if (typeof global.onRemoteTournamentUpdated === 'function') {
            global.onRemoteTournamentUpdated(remoteState);
        }
    }

    function enforceTournamentNavigation() {
        if (!isNetworkMode() || isNetworkHost()) return;

        var poolsOverlay = document.getElementById('poolsOverlay');
        if (poolsOverlay) poolsOverlay.style.display = 'none';

        var stage = gameState.tournamentStage;
        var hasBracket = !!gameState.bracket;

        if (hasBracket && stage !== 'pool-fights' && stage !== 'pools') {
            var poolSelect = document.getElementById('poolSelectOverlay');
            if (poolSelect && poolSelect.style.display === 'flex' && isNetworkArenaDevice()) {
                hidePoolDetailView();
                poolSelect.style.display = 'none';
            }
        }

        if (stage === 'bracket' || stage === 'bracket-fights' || hasBracket) {
            if (poolsOverlay) poolsOverlay.style.display = 'none';
        }
    }

    function getNextPendingPoolMatch(poolId, arenaId) {
        arenaId = arenaId || getLocalArenaId();
        var matches = getPoolMatches(poolId);
        for (var i = 0; i < matches.length; i++) {
            if (isMatchAvailableForArena(matches[i], arenaId)) return matches[i];
        }
        return null;
    }

    function isPoolComplete(poolId) {
        var matches = getPoolMatches(poolId);
        if (matches.length === 0) return false;
        for (var i = 0; i < matches.length; i++) {
            if (matches[i].status !== 'done') return false;
        }
        return true;
    }

    function allPoolsComplete() {
        if (gameState.pools.length === 0) return false;
        for (var i = 0; i < gameState.pools.length; i++) {
            var poolId = gameState.pools[i].id;
            if (!gameState.poolMatches[poolId] || gameState.poolMatches[poolId].length === 0) {
                return false;
            }
            if (!isPoolComplete(poolId)) return false;
        }
        return true;
    }

    function isPlayoffSessionActive() {
        return isPlayoffTournament() && !!gameState.playoffStarted;
    }

    function showSecretaryTerminal() {
        document.getElementById('secretaryTerminal').classList.remove('hidden');
    }

    function isBracketStageActive() {
        return !!gameState.bracket ||
            gameState.tournamentStage === 'bracket' ||
            gameState.tournamentStage === 'bracket-fights';
    }

        function updateFormBracketButtons() {
        var canForm = allPoolsComplete() && !gameState.bracket && (!isNetworkMode() || isNetworkHost());
        var canRank = isPlayoffSessionActive();
        var ids = ['formBracketBtn', 'formBracketBtnPool'];
        for (var i = 0; i < ids.length; i++) {
            var btn = document.getElementById(ids[i]);
            if (!btn) continue;
            if (gameState.bracket) {
                btn.classList.add('hidden');
                continue;
            }
            btn.classList.remove('hidden');
            btn.disabled = !canForm;
            btn.title = canForm ? '' : 'Сначала проведите все бои во всех пулах';
        }
        var rankingsBtn = document.getElementById('showPoolRankingsBtn');
        if (rankingsBtn) {
            if (canRank) rankingsBtn.classList.remove('hidden');
            else rankingsBtn.classList.add('hidden');
        }
        var rankingsBarBtn = document.getElementById('showRankingsBtnBar');
        if (rankingsBarBtn) {
            if (canRank) rankingsBarBtn.classList.remove('hidden');
            else rankingsBarBtn.classList.add('hidden');
        }
    }

    function updatePlayoffTerminalButtons() {
        var bracketStage = isBracketStageActive();
        var playoffOn = isPlayoffSessionActive();

        var poolBtn = document.getElementById('selectPoolBtn');
        if (poolBtn) {
            if (bracketStage && !isNetworkMode()) poolBtn.classList.add('hidden');
            else if (playoffOn) {
                poolBtn.classList.remove('hidden');
                poolBtn.textContent = isNetworkMode() ? 'Выбрать бой' : 'Выбрать пул';
            }
        }

        var formIds = ['formBracketBtn', 'formBracketBtnPool'];
        for (var f = 0; f < formIds.length; f++) {
            var formBtn = document.getElementById(formIds[f]);
            if (!formBtn) continue;
            if (bracketStage) formBtn.classList.add('hidden');
        }

        var randomPoolBtn = document.getElementById('randomPoolResultsBtn');
        if (randomPoolBtn) {
            if (bracketStage) randomPoolBtn.classList.add('hidden');
            else if (playoffOn) randomPoolBtn.classList.remove('hidden');
        }

        var rankingsBarBtn = document.getElementById('showRankingsBtnBar');
        if (rankingsBarBtn) {
            if (playoffOn) rankingsBarBtn.classList.remove('hidden');
            else rankingsBarBtn.classList.add('hidden');
        }

        var arenaWrap = document.getElementById('tournamentBarArenaWrap');
        if (arenaWrap) {
            if (isNetworkMode() && playoffOn) {
                arenaWrap.classList.remove('hidden');
                populateNetworkArenaSelects();
            } else {
                arenaWrap.classList.add('hidden');
            }
        }
    }

    function updateOpenBracketButton() {
        var btn = document.getElementById('openBracketBtn');
        if (!btn) return;
        if (gameState.bracket) {
            btn.classList.remove('hidden');
        } else {
            btn.classList.add('hidden');
        }
    }

    function getPoolStandings(poolId) {
        var standings = {};
        var pool = getPoolById(poolId);
        if (!pool) return standings;

        for (var i = 0; i < pool.participantIds.length; i++) {
            standings[pool.participantIds[i]] = { wins: 0, losses: 0, points: 0 };
        }

        var matches = getPoolMatches(poolId);
        for (var m = 0; m < matches.length; m++) {
            var match = matches[m];
            if (match.status !== 'done' || !match.winnerId) continue;
            standings[match.winnerId].wins++;
            var loser = match.winnerId === match.fighter1Id ? match.fighter2Id : match.fighter1Id;
            if (standings[loser]) standings[loser].losses++;
        }
        return standings;
    }

    function getPoolWinnerId(poolId) {
        var standings = getPoolStandings(poolId);
        var pool = getPoolById(poolId);
        if (!pool || pool.participantIds.length === 0) return null;

        var bestId = null;
        var bestWins = -1;
        for (var i = 0; i < pool.participantIds.length; i++) {
            var id = pool.participantIds[i];
            var w = standings[id] ? standings[id].wins : 0;
            if (w > bestWins) {
                bestWins = w;
                bestId = id;
            }
        }
        return bestId;
    }

    function readQualifyingAdvancersCount() {
        var ids = ['qualifyingAdvancersInput', 'qualifyingAdvancersInputPool'];
        for (var i = 0; i < ids.length; i++) {
            var el = document.getElementById(ids[i]);
            if (el && el.value !== '') {
                var val = parseInt(el.value, 10);
                if (!isNaN(val)) return val;
            }
        }
        if (gameState.qualifyingAdvancersCount) {
            return gameState.qualifyingAdvancersCount;
        }
        return null;
    }

    function syncQualifyingAdvancersInputs(count) {
        var text = count != null && count !== '' ? String(count) : '';
        var ids = ['qualifyingAdvancersInput', 'qualifyingAdvancersInputPool'];
        for (var i = 0; i < ids.length; i++) {
            var el = document.getElementById(ids[i]);
            if (el) el.value = text;
        }
        gameState.qualifyingAdvancersCount = count;
    }

    function getGlobalPoolStandings() {
        if (typeof PoolRankings !== 'undefined' && PoolRankings.buildPoolRankingsList) {
            return PoolRankings.buildPoolRankingsList(gameState.participants, gameState.poolMatches);
        }
        return [];
    }

    function selectAdvancersForBracket(count) {
        var standings = getGlobalPoolStandings();
        if (standings.length < count) return null;

        return standings.slice(0, count).map(function(row, index) {
            return {
                participantId: row.participantId,
                name: row.name,
                fromPool: null,
                seed: index + 1,
                poolWins: row.wins,
                poolLosses: row.losses
            };
        });
    }

    function getLastSideRoundIndex() {
        if (!gameState.bracket || !gameState.bracket.left.rounds.length) return 0;
        return gameState.bracket.left.rounds.length - 1;
    }

    function ensureBracketPlayoffFields() {
        if (!gameState.bracket) return;
        if (!gameState.bracket.thirdPlace) {
            gameState.bracket.thirdPlace = createBracketMatch('br_bronze', null, null);
        }
        if (!gameState.bracket.playoffPhase) {
            var maxR = getLastSideRoundIndex();
            if (gameState.bracket.activeRoundIndex <= maxR) {
                gameState.bracket.playoffPhase = 'side';
            } else if (gameState.bracket.thirdPlace.status !== 'done') {
                gameState.bracket.playoffPhase = 'bronze';
            } else if (gameState.bracket.final.status !== 'done') {
                gameState.bracket.playoffPhase = 'final';
            } else {
                gameState.bracket.playoffPhase = 'complete';
            }
        }
    }

    function isBronzeRoundActive() {
        if (!gameState.bracket) return false;
        ensureBracketPlayoffFields();
        return gameState.bracket.playoffPhase === 'bronze';
    }

    function isFinalRoundActive() {
        if (!gameState.bracket) return false;
        ensureBracketPlayoffFields();
        return gameState.bracket.playoffPhase === 'final';
    }

    function areSemifinalsComplete() {
        if (!gameState.bracket) return false;
        var maxR = getLastSideRoundIndex();
        return isSideRoundComplete('left', maxR) && isSideRoundComplete('right', maxR);
    }

    function getSemiFinalMatch(side) {
        var maxR = getLastSideRoundIndex();
        var round = gameState.bracket[side].rounds[maxR];
        return round && round.length ? round[0] : null;
    }

    function getMatchLoser(match) {
        if (!match || match.status !== 'done' || !match.winner) return null;
        if (match.winner === match.fighter1) return match.fighter2;
        if (match.winner === match.fighter2) return match.fighter1;
        return null;
    }

    function syncBronzeFightersFromSemis() {
        if (!gameState.bracket || !areSemifinalsComplete()) return;

        var leftLoser = getMatchLoser(getSemiFinalMatch('left'));
        var rightLoser = getMatchLoser(getSemiFinalMatch('right'));

        if (leftLoser && !leftLoser.isBye) {
            gameState.bracket.thirdPlace.fighter1 = leftLoser;
        }
        if (rightLoser && !rightLoser.isBye) {
            gameState.bracket.thirdPlace.fighter2 = rightLoser;
        }
    }

    function resolveBronzeByesIfNeeded() {
        if (!gameState.bracket) return false;
        var match = gameState.bracket.thirdPlace;
        if (!match || match.status === 'done') return false;

        var f1 = match.fighter1;
        var f2 = match.fighter2;
        var f1Ok = f1 && !f1.isBye;
        var f2Ok = f2 && !f2.isBye;

        if (f1Ok && f2Ok) return false;

        if (f1Ok && !f2Ok) {
            match.winner = f1;
            match.status = 'done';
            return true;
        }
        if (f2Ok && !f1Ok) {
            match.winner = f2;
            match.status = 'done';
            return true;
        }

        match.status = 'done';
        match.winner = null;
        return true;
    }

    function enterBronzePlayoffPhase() {
        if (!gameState.bracket) return;
        syncBronzeFightersFromSemis();
        gameState.bracket.playoffPhase = 'bronze';
        if (resolveBronzeByesIfNeeded()) {
            tryAdvancePlayoffPhase(false);
        } else {
            resolveBracketByesInMatch(gameState.bracket.thirdPlace, { side: 'bronze' });
        }
    }

    function enterFinalPlayoffPhase() {
        if (!gameState.bracket) return;
        gameState.bracket.playoffPhase = 'final';
        resolveBracketByesInMatch(gameState.bracket.final, { side: 'final' });
    }

    function tryAdvancePlayoffPhase(showAlert) {
        if (!gameState.bracket) {
            return { advanced: false, fromLabel: '', toLabel: '' };
        }

        ensureBracketPlayoffFields();
        var fromLabel = getActiveBracketRoundLabel();
        var advanced = false;

        if (gameState.bracket.playoffPhase === 'bronze' &&
            gameState.bracket.thirdPlace.status === 'done') {
            enterFinalPlayoffPhase();
            advanced = true;
        } else if (gameState.bracket.playoffPhase === 'final' &&
            gameState.bracket.final.status === 'done') {
            gameState.bracket.playoffPhase = 'complete';
            advanced = true;
        }

        return {
            advanced: advanced,
            fromLabel: fromLabel,
            toLabel: getActiveBracketRoundLabel()
        };
    }

    function isSideRoundComplete(side, roundIndex) {
        var round = gameState.bracket[side].rounds[roundIndex];
        if (!round) return true;
        for (var i = 0; i < round.length; i++) {
            if (round[i].status !== 'done') return false;
        }
        return true;
    }

    function isBracketMatchDone(match) {
        return !!match && match.status === 'done';
    }

    function isBracketMatchInProgressOnOtherArena(match, arenaId) {
        return !!match && match.status === 'in_progress' && match.arenaId && match.arenaId !== arenaId;
    }

    function isCurrentBracketRoundComplete() {
        if (!gameState.bracket) return false;
        ensureBracketPlayoffFields();
        if (gameState.bracket.playoffPhase === 'bronze') {
            return gameState.bracket.thirdPlace.status === 'done';
        }
        if (gameState.bracket.playoffPhase === 'final') {
            return gameState.bracket.final.status === 'done';
        }
        if (gameState.bracket.playoffPhase === 'complete') return true;
        var ri = gameState.bracket.activeRoundIndex;
        return isSideRoundComplete('left', ri) && isSideRoundComplete('right', ri);
    }

    function getBracketDisplayFighter(fighter, loc, match) {
        if (!fighter) return null;
        return fighter;
    }

    function getActiveBracketRoundLabel() {
        if (!gameState.bracket) return '';
        ensureBracketPlayoffFields();
        if (gameState.bracket.playoffPhase === 'bronze') return 'Бой за 3-е место';
        if (gameState.bracket.playoffPhase === 'final') return 'Финал';
        if (gameState.bracket.playoffPhase === 'complete') return 'Турнир завершён';
        return getSideRoundTitle(
            gameState.bracket.activeRoundIndex,
            gameState.bracket.size
        );
    }

    function getBracketRoundLabel(roundIndex) {
        if (!gameState.bracket) return '';
        if (roundIndex > getLastSideRoundIndex()) return 'Плей-офф';
        return getSideRoundTitle(roundIndex, gameState.bracket.size);
    }

    function tryAdvanceBracketRound() {
        if (!gameState.bracket) {
            return { advanced: false, fromIndex: 0, fromLabel: '', toLabel: '' };
        }

        var maxSideRound = getLastSideRoundIndex();
        var fromIndex = gameState.bracket.activeRoundIndex;
        var advanced = false;

        resolveBracketByesInRound(gameState.bracket.activeRoundIndex);

        while (gameState.bracket.activeRoundIndex <= maxSideRound) {
            resolveBracketByesInRound(gameState.bracket.activeRoundIndex);

            if (!isSideRoundComplete('left', gameState.bracket.activeRoundIndex) ||
                !isSideRoundComplete('right', gameState.bracket.activeRoundIndex)) {
                break;
            }

            gameState.bracket.activeRoundIndex++;
            advanced = true;
            resolveBracketByesInRound(gameState.bracket.activeRoundIndex);
        }

        if (gameState.bracket.activeRoundIndex > maxSideRound) {
            if (gameState.bracket.playoffPhase === 'side' || !gameState.bracket.playoffPhase) {
                enterBronzePlayoffPhase();
                advanced = true;
            } else {
                tryAdvancePlayoffPhase(false);
            }
        }

        return {
            advanced: advanced,
            fromIndex: fromIndex,
            fromLabel: getBracketRoundLabel(fromIndex),
            toLabel: getActiveBracketRoundLabel()
        };
    }

    function syncBracketRoundProgression() {
        if (!gameState.bracket) return { advanced: false, fromLabel: '', toLabel: '' };
        ensureBracketPlayoffFields();
        if (gameState.bracket.playoffPhase === 'side') {
            resolveBracketByesInRound(gameState.bracket.activeRoundIndex);
            return tryAdvanceBracketRound();
        }
        tryAdvancePlayoffPhase(false);
        return { advanced: false, fromLabel: '', toLabel: getActiveBracketRoundLabel() };
    }

    function getBracketMatchLabel(loc) {
        if (loc.side === 'final') return 'Финал';
        if (loc.side === 'bronze') return 'Бронза';
        return getActiveBracketRoundLabel();
    }

    function collectAvailablePoolMatches(arenaId) {
        var items = [];
        for (var i = 0; i < gameState.pools.length; i++) {
            var pool = gameState.pools[i];
            if (!isPoolOnArena(pool.id, arenaId)) continue;
            var matches = getPoolMatches(pool.id);
            for (var m = 0; m < matches.length; m++) {
                var match = matches[m];
                if (!isMatchSelectableForArena(match, arenaId)) continue;
                items.push({
                    type: 'pool',
                    poolId: pool.id,
                    matchId: match.id,
                    match: match,
                    label: 'Пул ' + pool.number + ': ' +
                        getParticipantName(match.fighter1Id) + ' vs ' +
                        getParticipantName(match.fighter2Id) +
                        (match.status === 'in_progress' ? ' (продолжить)' : '')
                });
            }
        }
        return items;
    }

    function collectAvailableBracketMatches(arenaId) {
        var items = [];
        if (!gameState.bracket) return items;

        function tryAdd(match, loc) {
            if (!match) return;
            if (isNetworkArenaDevice()) {
                if (match.status !== 'in_progress' || match.arenaId !== arenaId) return;
            } else if (isBracketMatchInProgressOnOtherArena(match, arenaId)) {
                return;
            } else if (match.status === 'done') {
                if (!canEditMatchOnArena(match, arenaId)) return;
            } else if (match.status === 'in_progress') {
                if (match.arenaId !== arenaId) return;
            } else if (!canFightBracketMatch(match, loc)) {
                return;
            }
            var f1 = getBracketDisplayFighter(match.fighter1, loc, match) || match.fighter1;
            var f2 = getBracketDisplayFighter(match.fighter2, loc, match) || match.fighter2;
            if (!f1 || !f2) return;
            items.push({
                type: 'bracket',
                bracketLoc: { side: loc.side, roundIndex: loc.roundIndex, matchIndex: loc.matchIndex },
                match: match,
                label: getBracketMatchLabel(loc) + ': ' + f1.name + ' vs ' + f2.name
            });
        }

        ['left', 'right'].forEach(function(side) {
            var half = gameState.bracket[side];
            if (!half || !half.rounds) return;
            for (var r = 0; r < half.rounds.length; r++) {
                for (var m = 0; m < half.rounds[r].length; m++) {
                    tryAdd(half.rounds[r][m], { side: side, roundIndex: r, matchIndex: m });
                }
            }
        });
        tryAdd(gameState.bracket.thirdPlace, { side: 'bronze', roundIndex: 0, matchIndex: 0 });
        tryAdd(gameState.bracket.final, { side: 'final', roundIndex: 0, matchIndex: 0 });
        return items;
    }

    function readSelectedArenaId() {
        var selectIds = ['tournamentBarArenaSelect', 'bracketArenaSelect', 'arenaSelectInput'];
        for (var i = 0; i < selectIds.length; i++) {
            var select = document.getElementById(selectIds[i]);
            if (!select || !select.value) continue;
            var val = parseInt(select.value, 10);
            if (val > 0) return val;
        }
        if (gameState.hostSelectedArenaId) return gameState.hostSelectedArenaId;
        return getLocalArenaId() || 1;
    }

    var currentDetailPoolId = null;

    function ensurePoolArenaAssignments() {
        if (!gameState.poolArenaAssignments) {
            gameState.poolArenaAssignments = {};
        }
    }

    function getPoolAssignedArena(poolId) {
        ensurePoolArenaAssignments();
        var val = gameState.poolArenaAssignments[poolId];
        return val ? parseInt(val, 10) : null;
    }

    function setPoolAssignedArena(poolId, arenaId) {
        ensurePoolArenaAssignments();
        if (arenaId && arenaId > 0) {
            gameState.poolArenaAssignments[poolId] = arenaId;
        } else {
            delete gameState.poolArenaAssignments[poolId];
        }
    }

    function isPoolOnArena(poolId, arenaId) {
        var assigned = getPoolAssignedArena(poolId);
        if (!assigned) {
            return isNetworkHost() || !isNetworkMode();
        }
        return assigned === arenaId;
    }

    function getPoolMatchStats(poolId) {
        var matches = getPoolMatches(poolId);
        var done = 0;
        for (var i = 0; i < matches.length; i++) {
            if (matches[i].status === 'done') done++;
        }
        return { total: matches.length, done: done };
    }

    function getArenaCountForSelect() {
        if (typeof NetworkSync !== 'undefined' && NetworkSync.getState().arenaCount) {
            return NetworkSync.getState().arenaCount;
        }
        return 8;
    }

    function populatePoolArenaSelect(select, selectedArena, includeEmpty) {
        if (!select) return;
        select.innerHTML = '';
        if (includeEmpty) {
            var opt0 = document.createElement('option');
            opt0.value = '';
            opt0.textContent = '— не назначена —';
            select.appendChild(opt0);
        }
        var count = getArenaCountForSelect();
        for (var a = 1; a <= count; a++) {
            var opt = document.createElement('option');
            opt.value = String(a);
            opt.textContent = 'Площадка ' + a;
            if (selectedArena === a) opt.selected = true;
            select.appendChild(opt);
        }
    }

    function hidePoolDetailView() {
        var main = document.getElementById('poolSelectMainView');
        var detail = document.getElementById('poolSelectDetailView');
        if (main) main.style.display = '';
        if (detail) detail.style.display = 'none';
        currentDetailPoolId = null;
    }

    function getArenaIdForPoolFight(poolId) {
        return getPoolAssignedArena(poolId) || readSelectedArenaId();
    }

    function renderPoolFightList(poolId) {
        var list = document.getElementById('poolFightList');
        if (!list) return;
        list.innerHTML = '';

        var arenaId = getArenaIdForPoolFight(poolId);
        var matches = getPoolMatches(poolId);
        if (!matches.length) {
            var empty = document.createElement('li');
            empty.className = 'pool-fight-item pool-fight-empty';
            empty.textContent = 'Бои не сгенерированы';
            list.appendChild(empty);
            return;
        }

        for (var i = 0; i < matches.length; i++) {
            (function(match) {
                var item = {
                    type: 'pool',
                    poolId: poolId,
                    matchId: match.id,
                    match: match
                };
                var li = document.createElement('li');
                li.className = 'pool-fight-item';
                if (match.status === 'done') li.classList.add('done');
                if (match.status === 'in_progress') li.classList.add('in-progress');

                var info = document.createElement('div');
                info.className = 'pool-fight-item-info';
                var mark = match.status === 'done' ? '✓' :
                    match.status === 'in_progress' ? '▶' : '○';
                info.textContent = mark + ' ' +
                    getParticipantName(match.fighter1Id) + ' vs ' +
                    getParticipantName(match.fighter2Id);
                li.appendChild(info);

                if (!isNetworkArenaDevice() &&
                    isPoolOnArena(poolId, arenaId) &&
                    isMatchSelectableForArena(match, arenaId)) {
                    var btn = document.createElement('button');
                    btn.type = 'button';
                    btn.className = 'menu-button';
                    if (match.status === 'done') {
                        btn.textContent = 'Изменить';
                    } else if (match.status === 'in_progress') {
                        btn.textContent = 'Продолжить';
                    } else {
                        btn.textContent = 'Провести';
                    }
                    btn.onclick = function() {
                        hidePoolSelectModal();
                        startNetworkMatch(item, getArenaIdForPoolFight(poolId), {
                            reopen: match.status === 'done'
                        });
                    };
                    li.appendChild(btn);
                }
                list.appendChild(li);
            })(matches[i]);
        }
    }

    function showPoolDetailView(poolId) {
        currentDetailPoolId = poolId;
        var pool = getPoolById(poolId);
        if (!pool) return;

        var main = document.getElementById('poolSelectMainView');
        var detail = document.getElementById('poolSelectDetailView');
        if (main) main.style.display = 'none';
        if (detail) detail.style.display = 'block';

        var title = document.getElementById('poolDetailTitle');
        if (title) title.textContent = 'Пул ' + pool.number;

        var assignRow = document.getElementById('poolArenaAssignRow');
        var assignSelect = document.getElementById('poolArenaAssignSelect');
        if (isNetworkHost()) {
            if (assignRow) assignRow.style.display = 'flex';
            populatePoolArenaSelect(assignSelect, getPoolAssignedArena(poolId), true);
        } else if (assignRow) {
            assignRow.style.display = 'none';
        }

        renderPoolFightList(poolId);
    }

    function onPoolArenaAssignChanged() {
        if (!currentDetailPoolId) return;
        var select = document.getElementById('poolArenaAssignSelect');
        var arenaId = select && select.value ? parseInt(select.value, 10) : null;
        var poolId = currentDetailPoolId;
        setPoolAssignedArena(poolId, arenaId);

        if (!gameState.poolMatches[poolId] || gameState.poolMatches[poolId].length === 0) {
            generatePoolMatches(poolId);
        }

        syncTournamentToServer().then(function() {
            if (arenaId) {
                dispatchNextFightForArena(arenaId);
            }
        });
        if (document.getElementById('poolSelectMainView').style.display !== 'none') {
            renderHostPoolButtons();
        }
        renderPoolFightList(currentDetailPoolId);
    }

    function renderHostPoolButtons() {
        var list = document.getElementById('poolSelectList');
        if (!list) return;
        list.innerHTML = '';

        for (var i = 0; i < gameState.pools.length; i++) {
            (function(pool) {
                var stats = getPoolMatchStats(pool.id);
                var btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'menu-button';
                var label = 'Пул ' + pool.number;
                if (stats.total) label += ' (' + stats.done + '/' + stats.total + ')';
                if (isPoolComplete(pool.id)) label += ' ✓';
                var assigned = getPoolAssignedArena(pool.id);
                if (assigned) label += ' · пл. ' + assigned;
                btn.textContent = label;
                btn.onclick = function() { showPoolDetailView(pool.id); };
                list.appendChild(btn);
            })(gameState.pools[i]);
        }
    }

    function showHostPoolManagementModal() {
        hidePoolDetailView();
        var title = document.getElementById('poolSelectTitle');
        var arenaRow = document.getElementById('arenaSelectRow');
        if (title) title.textContent = 'Пулы турнира';
        if (arenaRow) arenaRow.style.display = 'flex';

        var select = document.getElementById('arenaSelectInput');
        populateArenaSelectElement(select, readSelectedArenaId());
        if (select) {
            select.onchange = function() {
                onAdminArenaChanged();
            };
        }

        renderHostPoolButtons();
        document.getElementById('poolSelectOverlay').style.display = 'flex';
        updateTournamentBar();
    }

    function showArenaPoolSelectModal() {
        hidePoolDetailView();
        var list = document.getElementById('poolSelectList');
        var title = document.getElementById('poolSelectTitle');
        var arenaRow = document.getElementById('arenaSelectRow');
        if (title) title.textContent = 'Пулы вашей площадки';
        if (arenaRow) arenaRow.style.display = 'none';
        list.innerHTML = '';

        var localArenaId = getLocalArenaId();
        var hasPools = false;

        for (var i = 0; i < gameState.pools.length; i++) {
            var pool = gameState.pools[i];
            var assigned = getPoolAssignedArena(pool.id);
            if (!assigned || assigned !== localArenaId) continue;

            hasPools = true;
            (function(pid, pnum) {
                var stats = getPoolMatchStats(pid);
                var btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'menu-button';
                var label = 'Пул ' + pnum;
                if (stats.total) label += ' (' + stats.done + '/' + stats.total + ')';
                if (isPoolComplete(pid)) label += ' ✓';
                btn.textContent = label;
                btn.onclick = function() { showPoolDetailView(pid); };
                list.appendChild(btn);
            })(pool.id, pool.number);
        }

        if (!hasPools) {
            var empty = document.createElement('p');
            empty.className = 'participants-hint';
            empty.textContent = 'Администратор ещё не назначил пулы на вашу площадку.';
            list.appendChild(empty);
        }

        document.getElementById('poolSelectOverlay').style.display = 'flex';
        updateTournamentBar();
    }

    function showBracketMatchSelectModal() {
        var list = document.getElementById('poolSelectList');
        var title = document.getElementById('poolSelectTitle');
        var arenaRow = document.getElementById('arenaSelectRow');
        hidePoolDetailView();
        list.innerHTML = '';

        if (title) title.textContent = 'Выберите бой';
        if (arenaRow) arenaRow.style.display = 'flex';

        var select = document.getElementById('arenaSelectInput');
        populateArenaSelectElement(select, readSelectedArenaId());
        if (select) {
            select.onchange = function() {
                onAdminArenaChanged();
            };
        }

        var arenaId = readSelectedArenaId();
        var items = collectAvailablePoolMatches(arenaId).concat(collectAvailableBracketMatches(arenaId));

        if (items.length === 0) {
            var empty = document.createElement('p');
            empty.className = 'participants-hint';
            empty.textContent = 'Нет доступных боёв на выбранной площадке.';
            list.appendChild(empty);
        }

        for (var i = 0; i < items.length; i++) {
            (function(item) {
                var btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'menu-button';
                btn.textContent = item.label;
                var reopen = item.match && item.match.status === 'done';
                btn.onclick = function() {
                    hidePoolSelectModal();
                    startNetworkMatch(item, readSelectedArenaId(), { reopen: reopen });
                };
                list.appendChild(btn);
            })(items[i]);
        }

        document.getElementById('poolSelectOverlay').style.display = 'flex';
        updateTournamentBar();
    }

    function showArenaMatchSelectModal() {
        if (isNetworkArenaDevice()) {
            var remote = typeof NetworkSync !== 'undefined' ? NetworkSync.getState().lastRemoteState : null;
            updateArenaWaitOverlay(remote, findHostAssignedFightForArena(remote, getLocalArenaId()));
            return;
        }
        if (gameState.tournamentStage === 'pool-fights' ||
            (gameState.playoffStarted && !gameState.bracket)) {
            if (isNetworkHost()) {
                showHostPoolManagementModal();
                return;
            }
        }
        showBracketMatchSelectModal();
    }

    function startNetworkMatch(item, arenaId, options) {
        options = options || {};
        if (isNetworkArenaDevice()) {
            alert('Бой назначает только главное устройство. Дождитесь запуска на вашей площадке.');
            return;
        }
        if (isNetworkMode() && !isNetworkHost()) {
            alert('Только главное устройство может назначать бои.');
            return;
        }
        gameState.activeArenaId = arenaId;

        function beginFight() {
            if (item.type === 'pool') {
                gameState.activePoolId = item.poolId;
                gameState.activePoolMatchId = item.match.id;
                gameState.tournamentStage = 'pool-fights';
                loadPoolMatchIntoFight(item.match, {
                    restoreResult: !!options.reopen
                });
            } else {
                gameState.activeBracketMatch = item.bracketLoc;
                gameState.tournamentStage = 'bracket-fights';
                var bracketMatch = getBracketMatch(item.bracketLoc);
                loadBracketMatchIntoFight(bracketMatch, item.bracketLoc, {
                    restoreResult: !!options.reopen
                });
            }
            gameState.sessionStarted = true;
            gameState.tournamentMode = true;
            document.getElementById('secretaryTerminal').classList.remove('hidden');
            updateTournamentBar();
            if (typeof updateDisplay === 'function') updateDisplay();
        }

        if (!isNetworkMode()) {
            beginFight();
            return;
        }

        var claimPromise;
        if (item.type === 'pool') {
            claimPromise = NetworkSync.claimMatch('pool', {
                poolId: item.poolId,
                matchId: item.matchId
            }, arenaId, options);
        } else {
            claimPromise = NetworkSync.claimMatch('bracket', {
                bracketLoc: item.bracketLoc
            }, arenaId, options);
        }

        claimPromise.then(function(data) {
            if (data.state && data.state.tournament) {
                applyRemoteTournamentState(data.state);
            }
            beginFight();
        }).catch(function(err) {
            alert(err.message || 'Не удалось занять бой');
            showArenaMatchSelectModal();
        });
    }

    function showPoolSelectModal() {
        if (isNetworkMode()) {
            showArenaMatchSelectModal();
            return;
        }

        var list = document.getElementById('poolSelectList');
        var title = document.getElementById('poolSelectTitle');
        var arenaRow = document.getElementById('arenaSelectRow');
        list.innerHTML = '';
        if (title) title.textContent = 'Какой пул провести?';
        if (arenaRow) arenaRow.style.display = 'none';

        for (var i = 0; i < gameState.pools.length; i++) {
            var pool = gameState.pools[i];
            var btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'menu-button';
            var done = isPoolComplete(pool.id);
            var pending = getNextPendingPoolMatch(pool.id);
            var label = pool.number + ' пул';
            if (done) label += ' ✓';
            else if (pending) label += ' — есть бои';
            btn.textContent = label;
            btn.onclick = (function(pid) {
                return function() {
                    hidePoolSelectModal();
                    startPoolFights(pid);
                };
            })(pool.id);
            list.appendChild(btn);
        }

        document.getElementById('poolSelectOverlay').style.display = 'flex';
        updateTournamentBar();
    }

    function hidePoolSelectModal() {
        hidePoolDetailView();
        document.getElementById('poolSelectOverlay').style.display = 'none';
    }

    function startPlayoffFromPools() {
        if (!validatePoolsForStart()) return;

        for (var i = 0; i < gameState.pools.length; i++) {
            var pid = gameState.pools[i].id;
            if (!gameState.poolMatches[pid] || gameState.poolMatches[pid].length === 0) {
                generatePoolMatches(pid);
            }
        }

        gameState.playoffStarted = true;
        gameState.sessionStarted = true;
        gameState.tournamentMode = true;
        gameState.tournamentStage = 'pool-fights';

        if (typeof global.persistActiveTournamentNominationIfAny === 'function') {
            global.persistActiveTournamentNominationIfAny();
        }

        document.getElementById('poolsOverlay').style.display = 'none';
        showSecretaryTerminal();
        syncTournamentToServer().then(function() {
            showPoolSelectModal();
            updateTournamentBar();
        });
    }

    function startPoolFights(poolId) {
        if (!gameState.poolMatches[poolId] || gameState.poolMatches[poolId].length === 0) {
            generatePoolMatches(poolId);
        }

        var match = getNextPendingPoolMatch(poolId);
        if (!match) {
            alert('Все бои этого пула уже проведены.');
            updateTournamentBar();
            if (allPoolsComplete()) {
                alert('Этап пулов завершён. Можно сформировать сетку.');
            }
            return;
        }

        gameState.activePoolId = poolId;
        gameState.activePoolMatchId = match.id;
        gameState.tournamentStage = 'pool-fights';
        gameState.sessionStarted = true;
        gameState.tournamentMode = true;

        loadPoolMatchIntoFight(match);
        document.getElementById('secretaryTerminal').classList.remove('hidden');
        updateTournamentBar();
        if (typeof updateDisplay === 'function') updateDisplay();
    }

    function restoreFightScoresFromMatchResult(result) {
        if (!result) return;
        gameState.redScore = result.redScore || 0;
        gameState.blueScore = result.blueScore || 0;
        gameState.redBonus = result.redBonus || 0;
        gameState.blueBonus = result.blueBonus || 0;
        gameState.redRoundsWon = result.redRoundsWon || 0;
        gameState.blueRoundsWon = result.blueRoundsWon || 0;
    }

    function loadPoolMatchIntoFight(match, options) {
        options = options || {};
        resetFightState();
        gameState.redFighterName = getParticipantName(match.fighter1Id);
        gameState.blueFighterName = getParticipantName(match.fighter2Id);
        gameState.activePoolMatchMeta = {
            matchId: match.id,
            poolId: match.poolId,
            fighter1Id: match.fighter1Id,
            fighter2Id: match.fighter2Id
        };
        if (options.restoreResult && match.result) {
            restoreFightScoresFromMatchResult(match.result);
        }
    }

    function loadBracketMatchIntoFight(match, loc, options) {
        options = options || {};
        var f1 = getBracketDisplayFighter(match.fighter1, loc, match) || match.fighter1;
        var f2 = getBracketDisplayFighter(match.fighter2, loc, match) || match.fighter2;
        resetFightState();
        gameState.redFighterName = f1.name;
        gameState.blueFighterName = f2.name;
        if (options.restoreResult && match.result) {
            restoreFightScoresFromMatchResult(match.result);
        }
    }

    function findPoolMatch(matchId) {
        for (var poolId in gameState.poolMatches) {
            if (!gameState.poolMatches.hasOwnProperty(poolId)) continue;
            var matches = gameState.poolMatches[poolId];
            for (var i = 0; i < matches.length; i++) {
                if (matches[i].id === matchId) return matches[i];
            }
        }
        return null;
    }

    function saveActivePoolMatchResult() {
        if (!gameState.activePoolMatchId || !gameState.activePoolMatchMeta) return false;

        var match = findPoolMatch(gameState.activePoolMatchId);
        if (!match) return false;

        var winnerSide = getRoundWinner();
        var winnerId = null;
        if (winnerSide === 'red') winnerId = match.fighter1Id;
        else if (winnerSide === 'blue') winnerId = match.fighter2Id;

        match.status = 'done';
        match.winnerId = winnerId;
        match.result = {
            redScore: gameState.redScore,
            blueScore: gameState.blueScore,
            redBonus: gameState.redBonus,
            blueBonus: gameState.blueBonus,
            redRoundsWon: gameState.redRoundsWon,
            blueRoundsWon: gameState.blueRoundsWon,
            winnerSide: winnerSide,
            date: new Date().toLocaleString('ru-RU')
        };

        gameState.tournamentFightHistory.push({
            type: 'pool',
            poolId: match.poolId,
            matchId: match.id,
            fighter1: getParticipantName(match.fighter1Id),
            fighter2: getParticipantName(match.fighter2Id),
            winnerId: winnerId,
            result: match.result
        });

        var completedArenaId = match.arenaId || gameState.activeArenaId || readSelectedArenaId();
        gameState.activePoolMatchId = null;
        gameState.activePoolMatchMeta = null;
        updateTournamentBar();

        if (isNetworkArenaDevice()) {
            gameState.activePoolId = null;
            var term = document.getElementById('secretaryTerminal');
            if (term) term.classList.add('hidden');
            syncTournamentAfterFight();
            return true;
        }

        if (isNetworkHost()) {
            gameState.activePoolId = null;
            syncTournamentAfterFight().then(function() {
                if (completedArenaId) dispatchNextFightForArena(completedArenaId);
            });
            return true;
        }

        var next = getNextPendingPoolMatch(gameState.activePoolId);
        if (next) {
            if (confirm('Бой сохранён. Начать следующий бой в этом пуле?')) {
                gameState.activePoolMatchId = next.id;
                loadPoolMatchIntoFight(next);
                if (typeof updateDisplay === 'function') updateDisplay();
            }
        } else {
            gameState.activePoolId = null;
            if (allPoolsComplete() && !gameState.bracket) {
                showPoolRankingsOverlay({ showFormBracket: false });
            } else {
                alert('Все бои выбранного пула проведены.');
            }
        }
        syncTournamentAfterFight();
        return true;
    }

    function updateTournamentBar() {
        var bar = document.getElementById('tournamentBar');
        var info = document.getElementById('tournamentBarInfo');

        if (!bar) return;

        if (!isPlayoffSessionActive()) {
            bar.classList.add('hidden');
            return;
        }

        bar.classList.remove('hidden');
        var text = 'Плей-офф';
        if (gameState.activePoolId) {
            var pool = getPoolById(gameState.activePoolId);
            text += ' · Пул ' + (pool ? pool.number : '?');
        }
        if (gameState.activeArenaId) {
            text += ' · Площадка ' + gameState.activeArenaId;
        }
        if (gameState.activePoolMatchMeta) {
            text += ' · ' + gameState.redFighterName + ' vs ' + gameState.blueFighterName;
        }
        if (gameState.tournamentStage === 'bracket' || gameState.tournamentStage === 'bracket-fights') {
            text += ' · Сетка · ' + getActiveBracketRoundLabel();
        } else if (allPoolsComplete() && !gameState.bracket) {
            text += ' · Пулы завершены';
        }
        info.textContent = text;

        updateFormBracketButtons();
        updateOpenBracketButton();
        updatePlayoffTerminalButtons();
    }

    function nextPowerOfTwo(n) {
        var p = 1;
        while (p < n) p *= 2;
        return p;
    }

    function generateOlympicSeedOrder(size) {
        if (size <= 1) return [1];
        var half = generateOlympicSeedOrder(size / 2);
        var order = [];
        for (var i = 0; i < half.length; i++) {
            order.push(half[i]);
            order.push(size + 1 - half[i]);
        }
        return order;
    }

    var BRACKET_MAX_SIZE = 32;
    var BRACKET_SLOT_HEIGHT = 28;
    var BRACKET_SLOT_GAP = 8;
    var BRACKET_LABEL_OFFSET = 24;

    function getBracketCapacity(count) {
        var size = nextPowerOfTwo(Math.max(count, 2));
        return Math.min(BRACKET_MAX_SIZE, size);
    }

    function getSideRoundTitle(roundIndex, bracketSize) {
        var teamsOnSide = bracketSize / 2;
        var teamsAtRound = teamsOnSide / Math.pow(2, roundIndex);

        if (teamsAtRound >= 16) return '1/16 финала';
        if (teamsAtRound === 8) return '1/8 финала';
        if (teamsAtRound === 4) return '1/4 финала';
        if (teamsAtRound === 2) return 'Полуфинал';
        return 'Раунд ' + (roundIndex + 1);
    }

    function createBracketMatch(id, fighter1, fighter2) {
        return {
            id: id,
            fighter1: fighter1,
            fighter2: fighter2,
            status: 'pending',
            winner: null
        };
    }

    function createEmptyBracketMatch(id) {
        return createBracketMatch(id, null, null);
    }

    function buildHalfBracket(seededSlots, side) {
        var halfSize = seededSlots.length;
        var totalRounds = Math.round(Math.log(halfSize) / Math.log(2));
        var rounds = [];
        var firstRound = [];

        for (var i = 0; i < halfSize / 2; i++) {
            firstRound.push(createBracketMatch(
                side + '_r0_m' + i,
                seededSlots[i * 2],
                seededSlots[i * 2 + 1]
            ));
        }
        rounds.push(firstRound);

        for (var r = 1; r < totalRounds; r++) {
            var matchCount = halfSize / Math.pow(2, r + 1);
            var round = [];
            for (var m = 0; m < matchCount; m++) {
                round.push(createEmptyBracketMatch(side + '_r' + r + '_m' + m));
            }
            rounds.push(round);
        }
        return rounds;
    }

    function seedHalfSlots(fighters, halfSize) {
        var list = fighters.slice();
        while (list.length < halfSize) {
            list.push({
                participantId: null,
                name: 'BYE',
                fromPool: null,
                isBye: true
            });
        }
        var seedOrder = generateOlympicSeedOrder(halfSize);
        return seedOrder.map(function(seed) {
            return list[seed - 1];
        });
    }

    function splitAdvancersToHalves(advancers) {
        var left = [];
        var right = [];
        for (var i = 0; i < advancers.length; i++) {
            if (i % 2 === 0) left.push(advancers[i]);
            else right.push(advancers[i]);
        }
        return { left: left, right: right };
    }

    function getBracketMatch(loc) {
        if (!gameState.bracket) return null;
        if (loc.side === 'final') return gameState.bracket.final;
        if (loc.side === 'bronze') return gameState.bracket.thirdPlace;
        return gameState.bracket[loc.side].rounds[loc.roundIndex][loc.matchIndex];
    }

    function assignSemiLoserToBronze(side, loc) {
        var match = gameState.bracket[side].rounds[loc.roundIndex][loc.matchIndex];
        var loser = getMatchLoser(match);
        if (!loser || loser.isBye) return;
        if (side === 'left') gameState.bracket.thirdPlace.fighter1 = loser;
        else gameState.bracket.thirdPlace.fighter2 = loser;
    }

    function advanceBracketWinnerSilent(loc, winner) {
        if (!winner || !gameState.bracket) return;

        if (loc.side === 'final' || loc.side === 'bronze') return;

        var half = gameState.bracket[loc.side];
        var nextRoundIndex = loc.roundIndex + 1;

        if (nextRoundIndex >= half.rounds.length) {
            if (loc.side === 'left') gameState.bracket.final.fighter1 = winner;
            else gameState.bracket.final.fighter2 = winner;
            assignSemiLoserToBronze(side, loc);
            return;
        }

        var nextMatchIndex = Math.floor(loc.matchIndex / 2);
        var isFirstSlot = loc.matchIndex % 2 === 0;
        var nextMatch = half.rounds[nextRoundIndex][nextMatchIndex];
        if (!nextMatch) return;

        if (isFirstSlot) nextMatch.fighter1 = winner;
        else nextMatch.fighter2 = winner;
    }

    function resolveBracketByesInMatch(match, loc) {
        if (match.status !== 'pending' || !match.fighter1 || !match.fighter2) return false;
        if (!match.fighter1.isBye && !match.fighter2.isBye) return false;

        match.winner = match.fighter1.isBye ? match.fighter2 : match.fighter1;
        match.status = 'done';
        advanceBracketWinnerSilent(loc, match.winner);
        return true;
    }

    function resolveBracketByesInRound(roundIndex) {
        if (!gameState.bracket || roundIndex < 0) return;

        var changed = true;
        var maxSideRound = getLastSideRoundIndex();

        while (changed) {
            changed = false;
            if (roundIndex <= maxSideRound) {
                for (var s = 0; s < 2; s++) {
                    var sideName = s === 0 ? 'left' : 'right';
                    var round = gameState.bracket[sideName].rounds[roundIndex];
                    if (!round) continue;
                    for (var m = 0; m < round.length; m++) {
                        if (resolveBracketByesInMatch(round[m], {
                            side: sideName,
                            roundIndex: roundIndex,
                            matchIndex: m
                        })) {
                            changed = true;
                        }
                    }
                }
            } else if (isBronzeRoundActive()) {
                if (resolveBracketByesInMatch(gameState.bracket.thirdPlace, { side: 'bronze' })) {
                    changed = true;
                }
            } else if (isFinalRoundActive()) {
                if (resolveBracketByesInMatch(gameState.bracket.final, { side: 'final' })) {
                    changed = true;
                }
            }
        }
    }

    function computeBracketMatchOffsets(matchCount, roundIndex) {
        var blockH = getBracketMatchBlockHeight();
        if (roundIndex === 0) {
            var offsets = [];
            var step = blockH + BRACKET_SLOT_GAP;
            for (var i = 0; i < matchCount; i++) {
                offsets.push(i * step);
            }
            return offsets;
        }
        var prev = computeBracketMatchOffsets(matchCount * 2, roundIndex - 1);
        var offsets = [];
        for (var j = 0; j < matchCount; j++) {
            // Верх следующей пары — середина между верхами двух пар предыдущего раунда
            offsets.push((prev[j * 2] + prev[j * 2 + 1]) / 2);
        }
        return offsets;
    }

    function getBracketMatchAbsoluteTop(roundIndex, matchIndex, matchCount) {
        return computeBracketMatchOffsets(matchCount, roundIndex)[matchIndex] || 0;
    }

    function getBracketMatchMarginTop(roundIndex, matchIndex, matchCount) {
        var top = getBracketMatchAbsoluteTop(roundIndex, matchIndex, matchCount);
        if (matchIndex === 0) return top;
        var blockH = getBracketMatchBlockHeight();
        var prevTop = getBracketMatchAbsoluteTop(roundIndex, matchIndex - 1, matchCount);
        return top - prevTop - blockH;
    }

    function openBracketOverlay() {
        document.getElementById('bracketOverlay').style.display = 'flex';
        document.getElementById('poolSelectOverlay').style.display = 'none';
    }

    function buildBracketFromAdvancers(advancers) {
        var bracketSize = getBracketCapacity(advancers.length);
        var halfSize = bracketSize / 2;
        var halves = splitAdvancersToHalves(advancers);
        var leftSlots = seedHalfSlots(halves.left, halfSize);
        var rightSlots = seedHalfSlots(halves.right, halfSize);

        gameState.bracket = {
            size: bracketSize,
            advancers: advancers,
            qualifyingCount: advancers.length,
            activeRoundIndex: 0,
            left: { rounds: buildHalfBracket(leftSlots, 'left') },
            right: { rounds: buildHalfBracket(rightSlots, 'right') },
            thirdPlace: createBracketMatch('br_bronze', null, null),
            final: createBracketMatch('br_final', null, null),
            playoffPhase: 'side'
        };
        gameState.tournamentStage = 'bracket';
        resolveBracketByesInRound(0);
    }

    var poolRankingsContinueCallback = null;

    function renderPoolRankingsTable() {
        var tbody = document.getElementById('poolRankingsTableBody');
        if (!tbody) return;

        var rows = getGlobalPoolStandings();
        tbody.innerHTML = '';

        if (!rows.length) {
            var tr = document.createElement('tr');
            tr.innerHTML = '<td colspan="6" class="pool-rankings-empty">Нет результатов пулов</td>';
            tbody.appendChild(tr);
            return;
        }

        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            var tr = document.createElement('tr');
            tr.innerHTML =
                '<td class="col-place">' + (i + 1) + '</td>' +
                '<td class="col-name">' + escapeHtml(row.name) + '</td>' +
                '<td class="col-num">' + row.wins + '</td>' +
                '<td class="col-num">' + row.losses + '</td>' +
                '<td class="col-num">' + row.pointsScored + '</td>' +
                '<td class="col-num">' + row.pointsReceived + '</td>';
            tbody.appendChild(tr);
        }
    }

    function escapeHtml(text) {
        return String(text)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function showPoolRankingsOverlay(options) {
        options = options || {};
        poolRankingsContinueCallback = typeof options.onContinue === 'function' ? options.onContinue : null;

        renderPoolRankingsTable();

        var continueBtn = document.getElementById('poolRankingsContinueBtn');
        if (continueBtn) {
            if (options.showFormBracket) {
                continueBtn.classList.remove('hidden');
            } else {
                continueBtn.classList.add('hidden');
            }
        }

        var overlay = document.getElementById('poolRankingsOverlay');
        if (overlay) overlay.style.display = 'flex';
    }

    function hidePoolRankingsOverlay() {
        var overlay = document.getElementById('poolRankingsOverlay');
        if (overlay) overlay.style.display = 'none';
        poolRankingsContinueCallback = null;
    }

    function confirmPoolRankingsAndFormBracket() {
        var cb = poolRankingsContinueCallback;
        hidePoolRankingsOverlay();
        if (cb) cb();
    }

    function showPoolRankingsFromToolbar() {
        if (!gameState || !gameState.participants || gameState.participants.length === 0) {
            alert('Нет участников для рейтинга.');
            return;
        }
        showPoolRankingsOverlay({ showFormBracket: false });
    }

    function previewPoolRankingsDemo() {
        gameState.participants = [
            { id: 'p1', name: 'Алексей' },
            { id: 'p2', name: 'Борис' },
            { id: 'p3', name: 'Виктор' },
            { id: 'p4', name: 'Григорий' },
            { id: 'p5', name: 'Дмитрий' },
            { id: 'p6', name: 'Егор' }
        ];
        gameState.poolMatches = {
            demo: [
                { id: 'd1', fighter1Id: 'p1', fighter2Id: 'p2', status: 'done', winnerId: 'p1',
                    result: { redScore: 10, blueScore: 5, redBonus: 0, blueBonus: 0 } },
                { id: 'd2', fighter1Id: 'p3', fighter2Id: 'p4', status: 'done', winnerId: 'p3',
                    result: { redScore: 12, blueScore: 8, redBonus: 0, blueBonus: 0 } },
                { id: 'd3', fighter1Id: 'p1', fighter2Id: 'p3', status: 'done', winnerId: 'p3',
                    result: { redScore: 7, blueScore: 11, redBonus: 0, blueBonus: 0 } },
                { id: 'd4', fighter1Id: 'p2', fighter2Id: 'p4', status: 'done', winnerId: 'p2',
                    result: { redScore: 9, blueScore: 6, redBonus: 0, blueBonus: 0 } },
                { id: 'd5', fighter1Id: 'p5', fighter2Id: 'p6', status: 'done', winnerId: 'p5',
                    result: { redScore: 14, blueScore: 9, redBonus: 0, blueBonus: 0 } }
            ]
        };
        document.getElementById('startMenuOverlay').style.display = 'none';
        showPoolRankingsOverlay({ showFormBracket: false });
    }

    function formBracket() {
        if (!allPoolsComplete()) {
            alert('Сначала проведите все бои во всех пулах.');
            return;
        }
        if (gameState.bracket) {
            openBracketOverlay();
            renderBracket();
            return;
        }

        showPoolRankingsOverlay({
            showFormBracket: true,
            onContinue: formBracketAfterRankings
        });
        hidePoolSelectModal();
    }

    function formBracketAfterRankings() {
        hidePoolRankingsOverlay();

        var count = readQualifyingAdvancersCount();
        if (!count || count < 2) {
            alert('Укажите, сколько участников проходит в сетку (от 2 до ' + BRACKET_MAX_SIZE + ').');
            return;
        }
        if (count > BRACKET_MAX_SIZE) {
            count = BRACKET_MAX_SIZE;
        }

        var advancers = selectAdvancersForBracket(count);
        if (!advancers || advancers.length < 2) {
            alert('Недостаточно участников с результатами пулов для формирования сетки.');
            return;
        }

        syncQualifyingAdvancersInputs(count);
        buildBracketFromAdvancers(advancers);
        renderBracket();
        updateTournamentBar();
        updatePlayoffTerminalButtons();
        openBracketOverlay();

        alert(
            'Сетка сформирована: ' + advancers.length + ' участников. ' +
            'Сейчас открыт этап «' + getActiveBracketRoundLabel() + '».'
        );
        syncTournamentToServer().then(function() {
            var arenaId = readSelectedArenaId();
            if (arenaId) dispatchNextFightForArena(arenaId);
        });
    }

    function renderBracketHeaderControls() {
        var arenaWrap = document.getElementById('bracketArenaWrap');
        if (!arenaWrap) return;
        if (isNetworkMode()) {
            arenaWrap.style.display = 'flex';
            populateArenaSelectElement(document.getElementById('bracketArenaSelect'), readSelectedArenaId());
            var select = document.getElementById('bracketArenaSelect');
            if (select) {
                select.onchange = function() { onAdminArenaChanged(); };
            }
        } else {
            arenaWrap.style.display = 'none';
        }
    }

    function previewBracketDemo(count) {
        count = count || 32;
        if (count < 2) count = 2;
        if (count > BRACKET_MAX_SIZE) count = BRACKET_MAX_SIZE;

        var advancers = [];
        for (var i = 0; i < count; i++) {
            advancers.push({
                participantId: 'demo_' + i,
                name: 'Участник ' + (i + 1),
                fromPool: i + 1,
                seed: i + 1
            });
        }

        gameState.sessionMode = 'tournament';
        gameState.tournamentSystem = 'playoff';
        gameState.playoffStarted = true;
        gameState.sessionStarted = true;
        gameState.tournamentMode = true;
        gameState.participants = advancers.map(function(a, idx) {
            return { id: a.participantId, name: a.name };
        });

        buildBracketFromAdvancers(advancers);
        renderBracket();
        updateTournamentBar();

        document.getElementById('startMenuOverlay').style.display = 'none';
        document.getElementById('poolsOverlay').style.display = 'none';
        document.getElementById('poolSelectOverlay').style.display = 'none';
        document.getElementById('secretaryTerminal').classList.add('hidden');
        openBracketOverlay();
    }

    function maybeRunBracketPreviewFromUrl() {
        var params = new URLSearchParams(window.location.search);
        if (params.get('previewRankings')) {
            previewPoolRankingsDemo();
            return;
        }
        var preview = params.get('previewBracket');
        if (!preview) return;

        var count = parseInt(preview, 10);
        if (isNaN(count)) count = 32;
        previewBracketDemo(count);
    }

    function getBracketMatchBlockHeight() {
        return BRACKET_SLOT_HEIGHT * 2;
    }

    function addBracketJoinsToColumn(col, sideName, roundIndex, matchCount, isLastColumn) {
        if (isLastColumn || matchCount < 2) return;

        var layer = document.createElement('div');
        layer.className = 'bracket-joins';
        var blockH = getBracketMatchBlockHeight();

        for (var m = 0; m < matchCount; m += 2) {
            if (m + 1 >= matchCount) break;
            var topM = BRACKET_LABEL_OFFSET + getBracketMatchAbsoluteTop(roundIndex, m, matchCount);
            var topM1 = BRACKET_LABEL_OFFSET + getBracketMatchAbsoluteTop(roundIndex, m + 1, matchCount);
            var center1 = topM + blockH / 2;
            var center2 = topM1 + blockH / 2;
            var join = document.createElement('div');
            join.className = 'bracket-join' + (sideName === 'right' ? ' is-right' : '');
            join.style.top = center1 + 'px';
            join.style.height = Math.max(2, center2 - center1) + 'px';
            layer.appendChild(join);
        }
        col.appendChild(layer);
    }

    function formatBracketFighter(fighter) {
        if (!fighter) return { text: '\u00a0', className: 'tbd' };
        if (fighter.isBye) return { text: 'BYE', className: 'tbd' };
        return { text: fighter.name, className: '' };
    }

    function formatBracketMatchScore(match) {
        if (!match || !match.result) return '';
        return match.result.redScore + ' : ' + match.result.blueScore;
    }

    function createBracketFighterRowEl(fighter, match, loc, position) {
        var info = formatBracketFighter(fighter);
        var row = document.createElement('div');
        row.className = 'bracket-fighter-row ' + position + ' ' + info.className;
        if (match.winner && match.winner === fighter && fighter && !fighter.isBye) {
            row.classList.add('winner-row');
        }
        row.textContent = info.text;
        row.title = info.text.trim() || '';
        return row;
    }

    function isBracketMatchReady(match) {
        if (!match || match.status !== 'pending') return false;
        if (!match.fighter1 || !match.fighter2) return false;
        if (match.fighter1.isBye || match.fighter2.isBye) return false;
        return true;
    }

    function canFightBracketMatch(match, loc) {
        if (!match || !gameState.bracket) return false;

        if (loc.side === 'bronze') {
            if (!isBronzeRoundActive()) return false;
        } else if (loc.side === 'final') {
            if (!isFinalRoundActive()) return false;
        } else if (loc.roundIndex !== gameState.bracket.activeRoundIndex) {
            return false;
        }

        return isBracketMatchReady(match);
    }

    function createBracketMatchEl(match, loc, matchCount) {
        var card = document.createElement('div');
        card.className = 'bracket-match' + (match.status === 'done' ? ' done' : '');
        card.style.marginTop = getBracketMatchMarginTop(loc.roundIndex, loc.matchIndex, matchCount) + 'px';

        var pairBox = document.createElement('div');
        pairBox.className = 'bracket-match-pair';
        if (loc.side === 'bronze' && isBronzeRoundActive()) {
            pairBox.classList.add('bracket-pair-active');
        } else if (loc.side !== 'final' && loc.roundIndex === gameState.bracket.activeRoundIndex) {
            pairBox.classList.add('bracket-pair-active');
        } else if (loc.side === 'final' && isFinalRoundActive()) {
            pairBox.classList.add('bracket-pair-active');
        }

        var f1 = getBracketDisplayFighter(match.fighter1, loc, match);
        var f2 = getBracketDisplayFighter(match.fighter2, loc, match);
        pairBox.appendChild(createBracketFighterRowEl(f1, match, loc, 'top'));
        pairBox.appendChild(createBracketFighterRowEl(f2, match, loc, 'bottom'));
        card.appendChild(pairBox);

        if (match.status === 'done' && match.result) {
            var scoreEl = document.createElement('div');
            scoreEl.className = 'bracket-match-score';
            scoreEl.textContent = formatBracketMatchScore(match);
            scoreEl.title = 'Счёт боя';
            card.appendChild(scoreEl);
        } else if (match.status === 'in_progress') {
            var progressEl = document.createElement('div');
            progressEl.className = 'bracket-match-score';
            progressEl.textContent = 'Идёт бой · площадка ' + (match.arenaId || '?');
            card.appendChild(progressEl);
        } else if (match.status === 'done') {
            var editBtn = document.createElement('button');
            editBtn.type = 'button';
            editBtn.className = 'menu-button bracket-fight-btn';
            editBtn.textContent = 'Изменить';
            editBtn.onclick = function() {
                startBracketMatch(loc, { reopen: true });
            };
            card.appendChild(editBtn);
        } else if (canFightBracketMatch(match, loc)) {
            var btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'menu-button bracket-fight-btn';
            btn.textContent = 'Бой';
            btn.onclick = function() {
                startBracketMatch(loc);
            };
            card.appendChild(btn);
        }

        return card;
    }

    function renderBracketSide(sideName) {
        var half = gameState.bracket[sideName];
        var halfEl = document.createElement('div');
        halfEl.className = 'bracket-half bracket-half--' + sideName;

        var columnsEl = document.createElement('div');
        columnsEl.className = 'bracket-columns';

        for (var r = 0; r < half.rounds.length; r++) {
            var col = document.createElement('div');
            col.className = 'bracket-column';
            if (r === half.rounds.length - 1) {
                col.classList.add('is-semifinal');
            }
            if (r === gameState.bracket.activeRoundIndex) {
                col.classList.add('bracket-column-active');
            }

            var label = document.createElement('div');
            label.className = 'bracket-column-label';
            label.textContent = getSideRoundTitle(r, gameState.bracket.size);
            col.appendChild(label);

            var matchCount = half.rounds[r].length;
            for (var m = 0; m < matchCount; m++) {
                col.appendChild(createBracketMatchEl(half.rounds[r][m], {
                    side: sideName,
                    roundIndex: r,
                    matchIndex: m
                }, matchCount));
            }
            addBracketJoinsToColumn(col, sideName, r, matchCount, r === half.rounds.length - 1);
            columnsEl.appendChild(col);
        }

        halfEl.appendChild(columnsEl);
        return halfEl;
    }

    function renderBracketCenter() {
        var center = document.createElement('div');
        center.className = 'bracket-center connector-left connector-right';
        var finalLoc = { side: 'final', roundIndex: 0, matchIndex: 0 };
        var bronzeLoc = { side: 'bronze', roundIndex: 0, matchIndex: 0 };

        var championWrap = document.createElement('div');
        championWrap.className = 'bracket-champion-slot bracket-center-finals';
        var championSlot = document.createElement('div');
        var champInfo = { text: '\u00a0', className: 'tbd champion' };
        var champFighter = gameState.bracket.final.winner;
        if (gameState.bracket.final.status === 'done' && champFighter) {
            champInfo = formatBracketFighter(champFighter);
            champInfo.className = 'champion';
        }
        championSlot.className = 'bracket-slot ' + champInfo.className;
        championSlot.textContent = champInfo.text;
        championWrap.appendChild(championSlot);
        center.appendChild(championWrap);

        var bronzeLabel = document.createElement('div');
        bronzeLabel.className = 'bracket-center-label bracket-center-label--spaced';
        bronzeLabel.textContent = 'За 3-е место';
        center.appendChild(bronzeLabel);

        var bronzeWrap = document.createElement('div');
        bronzeWrap.className = 'bracket-bronze-match bracket-center-finals';
        if (isBronzeRoundActive()) {
            bronzeWrap.classList.add('bracket-column-active');
        }
        bronzeWrap.appendChild(createBracketMatchEl(gameState.bracket.thirdPlace, bronzeLoc, 1));
        center.appendChild(bronzeWrap);

        var thirdSlotWrap = document.createElement('div');
        thirdSlotWrap.className = 'bracket-third-slot bracket-center-finals';
        var thirdSlot = document.createElement('div');
        var thirdInfo = { text: '\u00a0', className: 'tbd third-place' };
        var thirdFighter = gameState.bracket.thirdPlace.winner;
        if (gameState.bracket.thirdPlace.status === 'done' && thirdFighter) {
            thirdInfo = formatBracketFighter(thirdFighter);
            thirdInfo.className = 'third-place';
        }
        thirdSlot.className = 'bracket-slot ' + thirdInfo.className;
        thirdSlot.textContent = thirdInfo.text;
        thirdSlotWrap.appendChild(thirdSlot);
        center.appendChild(thirdSlotWrap);

        var finalLabel = document.createElement('div');
        finalLabel.className = 'bracket-center-label bracket-center-label--spaced';
        finalLabel.textContent = 'Финал';
        center.appendChild(finalLabel);

        var finalWrap = document.createElement('div');
        finalWrap.className = 'bracket-final-match bracket-center-finals';
        if (isFinalRoundActive()) {
            finalWrap.classList.add('bracket-column-active');
        }
        finalWrap.appendChild(createBracketMatchEl(gameState.bracket.final, finalLoc, 1));
        center.appendChild(finalWrap);

        return center;
    }

    function renderBracket(skipProgression) {
        renderBracketHeaderControls();

        var container = document.getElementById('bracketContainer');
        container.innerHTML = '';

        if (!gameState.bracket) return;

        if (!skipProgression) {
            syncBracketRoundProgression();
        }

        var tree = document.createElement('div');
        tree.className = 'bracket-tree';
        tree.appendChild(renderBracketSide('left'));
        tree.appendChild(renderBracketCenter());
        tree.appendChild(renderBracketSide('right'));
        container.appendChild(tree);

        var offsets = computeBracketMatchOffsets(gameState.bracket.left.rounds[0].length, 0);
        var lastOffset = offsets.length ? offsets[offsets.length - 1] : 0;
        tree.style.minHeight = (lastOffset + BRACKET_SLOT_HEIGHT * 4 + 80) + 'px';
    }


    function startBracketMatch(loc, options) {
        options = options || {};
        var match = getBracketMatch(loc);
        if (!match) return;

        if (isNetworkMode()) {
            var f1 = getBracketDisplayFighter(match.fighter1, loc, match) || match.fighter1;
            var f2 = getBracketDisplayFighter(match.fighter2, loc, match) || match.fighter2;
            hidePoolSelectModal();
            document.getElementById('bracketOverlay').style.display = 'none';
            startNetworkMatch({
                type: 'bracket',
                bracketLoc: { side: loc.side, roundIndex: loc.roundIndex, matchIndex: loc.matchIndex },
                match: match,
                label: getBracketMatchLabel(loc) + ': ' + f1.name + ' vs ' + f2.name
            }, readSelectedArenaId(), {
                reopen: options.reopen || match.status === 'done'
            });
            return;
        }

        if (match.status === 'done' && !options.reopen) return;
        if (match.status !== 'done' && !canFightBracketMatch(match, loc)) return;

        gameState.tournamentStage = 'bracket-fights';
        gameState.activeBracketMatch = loc;
        document.getElementById('bracketOverlay').style.display = 'none';

        loadBracketMatchIntoFight(match, loc, {
            restoreResult: options.reopen && !!match.result
        });
        gameState.sessionStarted = true;
        document.getElementById('secretaryTerminal').classList.remove('hidden');
        updateTournamentBar();
        updatePlayoffTerminalButtons();
        if (typeof updateDisplay === 'function') updateDisplay();
    }

    function saveActiveBracketMatchResult() {
        if (!gameState.activeBracketMatch || !gameState.bracket) return false;

        var loc = gameState.activeBracketMatch;
        var match = getBracketMatch(loc);

        var winnerSide = getRoundWinner();
        var winner = null;
        if (winnerSide === 'red') winner = match.fighter1;
        else if (winnerSide === 'blue') winner = match.fighter2;

        if (!winner) {
            alert('Определите победителя (счёт не может быть равным).');
            return false;
        }

        match.status = 'done';
        match.winner = winner;
        match.result = {
            redScore: gameState.redScore,
            blueScore: gameState.blueScore,
            winnerSide: winnerSide
        };

        gameState.tournamentFightHistory.push({
            type: 'bracket',
            side: loc.side,
            roundIndex: loc.roundIndex,
            matchIndex: loc.matchIndex,
            winner: winner ? winner.name : 'Ничья',
            result: match.result
        });

        var completedArenaId = match.arenaId || gameState.activeArenaId || readSelectedArenaId();

        if (loc.side === 'bronze' || loc.side === 'final') {
            advanceBracketWinner(loc, winner);
            var playoffProgress = tryAdvancePlayoffPhase(false);
            return finalizeBracketMatchSave(completedArenaId, {
                playoffProgress: playoffProgress
            });
        }

        advanceBracketWinner(loc, winner);
        var progression = tryAdvanceBracketRound();
        return finalizeBracketMatchSave(completedArenaId, {
            progression: progression
        });
    }

    function finalizeBracketMatchSave(completedArenaId, opts) {
        opts = opts || {};
        gameState.activeBracketMatch = null;
        gameState.tournamentStage = 'bracket';
        updateTournamentBar();
        renderBracket(true);

        if (isNetworkArenaDevice()) {
            var term = document.getElementById('secretaryTerminal');
            if (term) term.classList.add('hidden');
            syncTournamentAfterFight();
            return true;
        }

        if (isNetworkHost()) {
            syncTournamentAfterFight().then(function() {
                if (completedArenaId) dispatchNextFightForArena(completedArenaId);
                dispatchForAllActiveArenas();
            });
            return true;
        }

        openBracketOverlay();
        if (isBracketComplete()) {
            var doneMsg = 'Турнир завершён! Победитель: ' + gameState.bracket.final.winner.name;
            if (gameState.bracket.thirdPlace.winner) {
                doneMsg += '. 3-е место: ' + gameState.bracket.thirdPlace.winner.name;
            }
            alert(doneMsg);
        } else if (opts.playoffProgress && opts.playoffProgress.advanced) {
            alert('Этап «' + opts.playoffProgress.fromLabel + '» завершён. Открыт этап «' +
                opts.playoffProgress.toLabel + '».');
        } else if (opts.progression && opts.progression.advanced) {
            alert('Этап «' + opts.progression.fromLabel + '» завершён. Открыт этап «' +
                opts.progression.toLabel + '».');
        }
        syncTournamentAfterFight();
        return true;
    }

    function advanceBracketWinner(loc, winner) {
        if (!winner || !gameState.bracket) return;
        if (loc.side === 'bronze' || loc.side === 'final') return;
        advanceBracketWinnerSilent(loc, winner);
        resolveBracketByesInRound(loc.roundIndex);
        if (loc.roundIndex + 1 <= getLastSideRoundIndex()) {
            resolveBracketByesInRound(loc.roundIndex + 1);
        }
        if (loc.roundIndex === getLastSideRoundIndex()) {
            syncBronzeFightersFromSemis();
        }
    }

    function isBracketComplete() {
        if (!gameState.bracket || !gameState.bracket.final) return false;
        ensureBracketPlayoffFields();
        return gameState.bracket.playoffPhase === 'complete';
    }

    function hideBracketOverlay() {
        document.getElementById('bracketOverlay').style.display = 'none';
    }

    function backFromBracket() {
        hideBracketOverlay();
        showSecretaryTerminal();
    }

    function backFromPoolSelect() {
        hidePoolSelectModal();
        if (isNetworkArenaDevice()) {
            return;
        }
        document.getElementById('poolsOverlay').style.display = 'flex';
    }

    function backFromTerminalToPoolSelect() {
        if (!isPlayoffSessionActive() || gameState.bracket) return;
        if (isNetworkArenaDevice()) {
            showArenaMatchSelectModal();
            return;
        }
        showPoolSelectModal();
    }

    function playoffGoBack() {
        if (!isPlayoffSessionActive()) return;

        var bracketVisible = document.getElementById('bracketOverlay').style.display === 'flex';
        var poolSelectVisible = document.getElementById('poolSelectOverlay').style.display === 'flex';

        if (bracketVisible) {
            backFromBracket();
            return;
        }
        if (poolSelectVisible) {
            if (isNetworkArenaDevice()) {
                hidePoolSelectModal();
                return;
            }
            backFromPoolSelect();
            return;
        }
        if (!gameState.bracket) {
            backFromTerminalToPoolSelect();
        } else {
            openBracketView();
        }
    }

    function resetPlayoffState() {
        gameState.pools = [];
        gameState.poolMatches = {};
        gameState.activePoolId = null;
        gameState.activePoolMatchId = null;
        gameState.activePoolMatchMeta = null;
        gameState.tournamentStage = null;
        gameState.bracket = null;
        gameState.activeBracketMatch = null;
        gameState.tournamentFightHistory = [];
        gameState.playoffStarted = false;
        gameState.qualifyingAdvancersCount = null;
        gameState.poolArenaAssignments = {};
        syncQualifyingAdvancersInputs(null);
        var poolsOverlay = document.getElementById('poolsOverlay');
        if (poolsOverlay) poolsOverlay.style.display = 'none';
        hidePoolSelectModal();
        hideBracketOverlay();
        updateFormBracketButtons();
        updateOpenBracketButton();
        updatePlayoffTerminalButtons();
    }

    function onSaveFightHook() {
        if (!isPlayoffTournament()) return false;
        if (gameState.tournamentStage === 'pool-fights' && gameState.activePoolMatchId) {
            return saveActivePoolMatchResult();
        }
        if (gameState.tournamentStage === 'bracket-fights' && gameState.activeBracketMatch) {
            return saveActiveBracketMatchResult();
        }
        return false;
    }

    function initPlayoffUI() {
        setupUnassignedDropZone();
    }

    global.TournamentPlayoff = {
        init: initPlayoffUI,
        isPlayoff: isPlayoffTournament,
        triggerImportParticipantsXlsx: triggerImportParticipantsXlsx,
        handleParticipantsXlsxSelected: handleParticipantsXlsxSelected,
        previewBracketDemo: previewBracketDemo,
        randomizePoolsDistribution: randomizePoolsDistribution,
        randomizePoolFightResults: randomizePoolFightResults,
        randomizeAllPoolFightResults: randomizeAllPoolFightResults,
        randomizeCurrentPoolFightResults: randomizeCurrentPoolFightResults,
        goToPoolsComposition: goToPoolsComposition,
        backFromPoolsToParticipants: backFromPoolsToParticipants,
        addPool: addPool,
        startPlayoffFromPools: startPlayoffFromPools,
        showPoolSelectModal: showPoolSelectModal,
        hidePoolSelectModal: hidePoolSelectModal,
        formBracket: formBracket,
        showPoolRankingsOverlay: showPoolRankingsOverlay,
        hidePoolRankingsOverlay: hidePoolRankingsOverlay,
        previewPoolRankingsDemo: previewPoolRankingsDemo,
        hideBracketOverlay: hideBracketOverlay,
        backFromBracket: backFromBracket,
        backFromPoolSelect: backFromPoolSelect,
        playoffGoBack: playoffGoBack,
        updateTournamentBar: updateTournamentBar,
        resetPlayoffState: resetPlayoffState,
        onSaveFightHook: onSaveFightHook,
        applyRemoteTournamentState: applyRemoteTournamentState,
        showArenaMatchSelectModal: showArenaMatchSelectModal,
        syncTournamentToServer: syncTournamentToServer,
        populateNetworkArenaSelects: populateNetworkArenaSelects,
        onAdminArenaChanged: onAdminArenaChanged,
        isNetworkMode: isNetworkMode
    };

    global.triggerImportParticipantsXlsx = triggerImportParticipantsXlsx;
    global.handleParticipantsXlsxSelected = handleParticipantsXlsxSelected;
    global.previewBracketDemo = previewBracketDemo;
    global.randomizePoolsDistribution = randomizePoolsDistribution;
    global.randomizePoolFightResults = randomizePoolFightResults;
    global.randomizeAllPoolFightResults = randomizeAllPoolFightResults;
    global.randomizeCurrentPoolFightResults = randomizeCurrentPoolFightResults;
    global.goToPoolsComposition = goToPoolsComposition;
    global.backFromPoolsToParticipants = backFromPoolsToParticipants;
    global.addPool = addPool;
    global.startPlayoffFromPools = startPlayoffFromPools;
    global.showPoolSelectModal = showPoolSelectModal;
    global.hidePoolDetailView = hidePoolDetailView;
    global.onPoolArenaAssignChanged = onPoolArenaAssignChanged;
    global.hidePoolSelectModal = hidePoolSelectModal;
    global.formBracket = formBracket;
    global.showPoolRankingsFromToolbar = showPoolRankingsFromToolbar;
    global.confirmPoolRankingsAndFormBracket = confirmPoolRankingsAndFormBracket;
    global.hidePoolRankingsOverlay = hidePoolRankingsOverlay;
    global.hideBracketOverlay = hideBracketOverlay;
    global.backFromBracket = backFromBracket;
    global.backFromPoolSelect = backFromPoolSelect;
    global.playoffGoBack = playoffGoBack;

    function openBracketView() {
        if (!gameState.bracket) {
            alert('Сетка ещё не сформирована.');
            return;
        }
        if (gameState.tournamentStage === 'pool-fights' || gameState.tournamentStage === 'pools') {
            gameState.tournamentStage = 'bracket';
            if (isNetworkHost()) {
                syncTournamentToServer().then(function() {
                    dispatchNextFightForArena(readSelectedArenaId());
                });
            }
        }
        renderBracket();
        openBracketOverlay();
    }

    function removeParticipantFromTournament(participantId) {
        removeParticipantFromAllPools(participantId);
        gameState.participants = gameState.participants.filter(function(p) {
            return p.id !== participantId;
        });
    }

    global.openBracketView = openBracketView;
    global.onAdminArenaChanged = onAdminArenaChanged;
    global.TournamentPlayoff.openBracketView = openBracketView;
    global.TournamentPlayoff.removeParticipantFromTournament = removeParticipantFromTournament;
    global.TournamentPlayoff.renderBracket = renderBracket;

})(window);
