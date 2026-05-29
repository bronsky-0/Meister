/* Плей-офф: пулы, бои этапа пулов, сетка на выбывание */
(function(global) {
    'use strict';

    var draggedParticipantId = null;

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

    function isMatchAvailableForArena(match, arenaId) {
        if (!match || match.status === 'done') return false;
        if (match.status === 'pending') return true;
        if (match.status === 'in_progress' && arenaId && match.arenaId === arenaId) return true;
        return false;
    }

    function syncTournamentToServer() {
        if (!isNetworkHost()) return Promise.resolve();
        return NetworkSync.pushTournament(NetworkSync.getTournamentSnapshot(gameState)).catch(function(err) {
            alert('Ошибка синхронизации: ' + (err.message || err));
        });
    }

    function syncTournamentAfterFight() {
        if (!isNetworkMode()) return Promise.resolve();
        return NetworkSync.completeMatch(NetworkSync.getTournamentSnapshot(gameState)).catch(function(err) {
            alert('Ошибка синхронизации результата: ' + (err.message || err));
        });
    }

    function applyRemoteTournamentState(remoteState) {
        if (!remoteState || !remoteState.tournament) return;
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

        if (document.getElementById('poolSelectOverlay').style.display === 'flex') {
            if (isNetworkMode()) showArenaMatchSelectModal();
            else showPoolSelectModal();
        }
        if (gameState.bracket && document.getElementById('bracketOverlay').style.display === 'flex') {
            renderBracket();
        }

        if (typeof global.onRemoteTournamentUpdated === 'function') {
            global.onRemoteTournamentUpdated(remoteState);
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
        var stats = {};
        for (var i = 0; i < gameState.participants.length; i++) {
            var participant = gameState.participants[i];
            stats[participant.id] = {
                participantId: participant.id,
                name: participant.name,
                wins: 0,
                losses: 0
            };
        }

        for (var poolId in gameState.poolMatches) {
            if (!gameState.poolMatches.hasOwnProperty(poolId)) continue;
            var matches = gameState.poolMatches[poolId];
            for (var m = 0; m < matches.length; m++) {
                var match = matches[m];
                if (match.status !== 'done' || !match.winnerId) continue;
                if (stats[match.winnerId]) stats[match.winnerId].wins++;
                var loserId = match.winnerId === match.fighter1Id ?
                    match.fighter2Id : match.fighter1Id;
                if (stats[loserId]) stats[loserId].losses++;
            }
        }

        var list = [];
        for (var id in stats) {
            if (stats.hasOwnProperty(id)) list.push(stats[id]);
        }

        list.sort(function(a, b) {
            if (b.wins !== a.wins) return b.wins - a.wins;
            if (a.losses !== b.losses) return a.losses - b.losses;
            return a.name.localeCompare(b.name, 'ru');
        });
        return list;
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
            var matches = getPoolMatches(pool.id);
            for (var m = 0; m < matches.length; m++) {
                var match = matches[m];
                if (!isMatchAvailableForArena(match, arenaId)) continue;
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
            if (!match || match.status === 'done') return;
            if (match.status === 'in_progress') {
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
        var select = document.getElementById('arenaSelectInput');
        if (select) {
            var val = parseInt(select.value, 10);
            if (val > 0) return val;
        }
        return getLocalArenaId() || 1;
    }

    function showArenaMatchSelectModal() {
        var list = document.getElementById('poolSelectList');
        var title = document.getElementById('poolSelectTitle');
        var arenaRow = document.getElementById('arenaSelectRow');
        list.innerHTML = '';

        if (title) title.textContent = 'Выберите бой';
        if (arenaRow) arenaRow.style.display = 'flex';

        var arenaCount = NetworkSync.getState().arenaCount || 1;
        var select = document.getElementById('arenaSelectInput');
        if (select) {
            select.innerHTML = '';
            for (var a = 1; a <= arenaCount; a++) {
                var opt = document.createElement('option');
                opt.value = String(a);
                opt.textContent = 'Площадка ' + a;
                if (getLocalArenaId() === a) opt.selected = true;
                select.appendChild(opt);
            }
            select.onchange = function() {
                showArenaMatchSelectModal();
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
                btn.onclick = function() {
                    hidePoolSelectModal();
                    startNetworkMatch(item, readSelectedArenaId());
                };
                list.appendChild(btn);
            })(items[i]);
        }

        document.getElementById('poolSelectOverlay').style.display = 'flex';
        updateTournamentBar();
    }

    function startNetworkMatch(item, arenaId) {
        gameState.activeArenaId = arenaId;

        function beginFight() {
            if (item.type === 'pool') {
                gameState.activePoolId = item.poolId;
                gameState.activePoolMatchId = item.match.id;
                gameState.tournamentStage = 'pool-fights';
                loadPoolMatchIntoFight(item.match);
            } else {
                gameState.activeBracketMatch = item.bracketLoc;
                gameState.tournamentStage = 'bracket-fights';
                var match = getBracketMatch(item.bracketLoc);
                var f1 = getBracketDisplayFighter(match.fighter1, item.bracketLoc, match) || match.fighter1;
                var f2 = getBracketDisplayFighter(match.fighter2, item.bracketLoc, match) || match.fighter2;
                resetFightState();
                gameState.redFighterName = f1.name;
                gameState.blueFighterName = f2.name;
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
            }, arenaId);
        } else {
            claimPromise = NetworkSync.claimMatch('bracket', {
                bracketLoc: item.bracketLoc
            }, arenaId);
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

    function loadPoolMatchIntoFight(match) {
        resetFightState();
        gameState.redFighterName = getParticipantName(match.fighter1Id);
        gameState.blueFighterName = getParticipantName(match.fighter2Id);
        gameState.activePoolMatchMeta = {
            matchId: match.id,
            poolId: match.poolId,
            fighter1Id: match.fighter1Id,
            fighter2Id: match.fighter2Id
        };
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

        gameState.activePoolMatchId = null;
        gameState.activePoolMatchMeta = null;

        var next = getNextPendingPoolMatch(gameState.activePoolId);
        updateTournamentBar();

        if (next) {
            if (confirm('Бой сохранён. Начать следующий бой в этом пуле?')) {
                gameState.activePoolMatchId = next.id;
                loadPoolMatchIntoFight(next);
                if (typeof updateDisplay === 'function') updateDisplay();
            }
        } else {
            alert('Все бои выбранного пула проведены.');
            gameState.activePoolId = null;
            if (allPoolsComplete()) {
                alert('Этап пулов завершён. Нажмите «Сформировать сетку».');
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
        syncTournamentToServer();
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

    function renderBracket() {
        var container = document.getElementById('bracketContainer');
        container.innerHTML = '';

        if (!gameState.bracket) return;

        syncBracketRoundProgression();

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


    function startBracketMatch(loc) {
        if (isNetworkMode()) {
            var match = getBracketMatch(loc);
            if (!match) return;
            var f1 = getBracketDisplayFighter(match.fighter1, loc, match) || match.fighter1;
            var f2 = getBracketDisplayFighter(match.fighter2, loc, match) || match.fighter2;
            hidePoolSelectModal();
            document.getElementById('bracketOverlay').style.display = 'none';
            startNetworkMatch({
                type: 'bracket',
                bracketLoc: { side: loc.side, roundIndex: loc.roundIndex, matchIndex: loc.matchIndex },
                match: match,
                label: getBracketMatchLabel(loc) + ': ' + f1.name + ' vs ' + f2.name
            }, readSelectedArenaId());
            return;
        }

        var match = getBracketMatch(loc);
        if (!match || match.status === 'done' || !canFightBracketMatch(match, loc)) return;

        var f1 = getBracketDisplayFighter(match.fighter1, loc, match) || match.fighter1;
        var f2 = getBracketDisplayFighter(match.fighter2, loc, match) || match.fighter2;

        gameState.tournamentStage = 'bracket-fights';
        gameState.activeBracketMatch = loc;
        document.getElementById('bracketOverlay').style.display = 'none';

        resetFightState();
        gameState.redFighterName = f1.name;
        gameState.blueFighterName = f2.name;
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

        if (loc.side === 'bronze' || loc.side === 'final') {
            advanceBracketWinner(loc, winner);
            var playoffProgress = tryAdvancePlayoffPhase(false);
            gameState.activeBracketMatch = null;
            gameState.tournamentStage = 'bracket';
            updateTournamentBar();
            renderBracket();
            openBracketOverlay();

            if (isBracketComplete()) {
                var msg = 'Турнир завершён! Победитель: ' + gameState.bracket.final.winner.name;
                if (gameState.bracket.thirdPlace.winner) {
                    msg += '. 3-е место: ' + gameState.bracket.thirdPlace.winner.name;
                }
                alert(msg);
            } else if (playoffProgress.advanced) {
                alert('Этап «' + playoffProgress.fromLabel + '» завершён. Открыт этап «' + playoffProgress.toLabel + '».');
            }
            syncTournamentAfterFight();
            return true;
        }

        advanceBracketWinner(loc, winner);
        var progression = tryAdvanceBracketRound();
        gameState.activeBracketMatch = null;
        gameState.tournamentStage = 'bracket';
        updateTournamentBar();
        renderBracket();
        openBracketOverlay();

        if (isBracketComplete()) {
            var doneMsg = 'Турнир завершён! Победитель: ' + gameState.bracket.final.winner.name;
            if (gameState.bracket.thirdPlace.winner) {
                doneMsg += '. 3-е место: ' + gameState.bracket.thirdPlace.winner.name;
            }
            alert(doneMsg);
        } else if (progression.advanced) {
            alert('Этап «' + progression.fromLabel + '» завершён. Открыт этап «' + progression.toLabel + '».');
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
        document.getElementById('poolsOverlay').style.display = 'flex';
    }

    function backFromTerminalToPoolSelect() {
        if (!isPlayoffSessionActive() || gameState.bracket) return;
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
    global.hidePoolSelectModal = hidePoolSelectModal;
    global.formBracket = formBracket;
    global.hideBracketOverlay = hideBracketOverlay;
    global.backFromBracket = backFromBracket;
    global.backFromPoolSelect = backFromPoolSelect;
    global.playoffGoBack = playoffGoBack;

    function openBracketView() {
        if (!gameState.bracket) {
            alert('Сетка ещё не сформирована.');
            return;
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
    global.TournamentPlayoff.openBracketView = openBracketView;
    global.TournamentPlayoff.removeParticipantFromTournament = removeParticipantFromTournament;
    global.TournamentPlayoff.renderBracket = renderBracket;

})(window);
