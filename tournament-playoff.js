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

    function getNextPendingPoolMatch(poolId) {
        var matches = getPoolMatches(poolId);
        for (var i = 0; i < matches.length; i++) {
            if (matches[i].status === 'pending') return matches[i];
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

    function updateFormBracketButtons() {
        var canForm = allPoolsComplete() && !gameState.bracket;
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

    function showPoolSelectModal() {
        var list = document.getElementById('poolSelectList');
        list.innerHTML = '';

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
        showPoolSelectModal();
        updateTournamentBar();
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
        if (gameState.activePoolMatchMeta) {
            text += ' · ' + gameState.redFighterName + ' vs ' + gameState.blueFighterName;
        }
        if (gameState.tournamentStage === 'bracket' || gameState.tournamentStage === 'bracket-fights') {
            text += ' · Сетка на выбывание';
        } else if (allPoolsComplete() && !gameState.bracket) {
            text += ' · Пулы завершены';
        }
        info.textContent = text;

        updateFormBracketButtons();
        updateOpenBracketButton();
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
    var BRACKET_SLOT_HEIGHT = 33;
    var BRACKET_SLOT_GAP = 10;

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
        return gameState.bracket[loc.side].rounds[loc.roundIndex][loc.matchIndex];
    }

    function advanceBracketWinnerSilent(loc, winner) {
        if (!winner || !gameState.bracket) return;

        if (loc.side === 'final') return;

        var half = gameState.bracket[loc.side];
        var nextRoundIndex = loc.roundIndex + 1;

        if (nextRoundIndex >= half.rounds.length) {
            if (loc.side === 'left') gameState.bracket.final.fighter1 = winner;
            else gameState.bracket.final.fighter2 = winner;
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

    function resolveBracketByes() {
        if (!gameState.bracket) return;

        var changed = true;
        while (changed) {
            changed = false;
            for (var s = 0; s < 2; s++) {
                var sideName = s === 0 ? 'left' : 'right';
                var rounds = gameState.bracket[sideName].rounds;
                for (var r = 0; r < rounds.length; r++) {
                    for (var m = 0; m < rounds[r].length; m++) {
                        if (resolveBracketByesInMatch(rounds[r][m], {
                            side: sideName,
                            roundIndex: r,
                            matchIndex: m
                        })) {
                            changed = true;
                        }
                    }
                }
            }
            if (resolveBracketByesInMatch(gameState.bracket.final, { side: 'final' })) {
                changed = true;
            }
        }
    }

    function computeBracketMatchOffsets(matchCount, roundIndex) {
        if (roundIndex === 0) {
            var offsets = [];
            var step = (BRACKET_SLOT_HEIGHT * 2 + BRACKET_SLOT_GAP) * 2;
            for (var i = 0; i < matchCount; i++) {
                offsets.push(i * step);
            }
            return offsets;
        }
        var prev = computeBracketMatchOffsets(matchCount * 2, roundIndex - 1);
        var offsets = [];
        var block = BRACKET_SLOT_HEIGHT * 2 + BRACKET_SLOT_GAP;
        for (var j = 0; j < matchCount; j++) {
            offsets.push((prev[j * 2] + prev[j * 2 + 1] + block) / 2);
        }
        return offsets;
    }

    function getBracketMatchMarginTop(roundIndex, matchIndex, matchCount) {
        return computeBracketMatchOffsets(matchCount, roundIndex)[matchIndex] || 0;
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
            left: { rounds: buildHalfBracket(leftSlots, 'left') },
            right: { rounds: buildHalfBracket(rightSlots, 'right') },
            final: createBracketMatch('br_final', null, null)
        };
        gameState.tournamentStage = 'bracket';
        resolveBracketByes();
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

        var sortedPools = gameState.pools.slice().sort(function(a, b) {
            return a.number - b.number;
        });
        var advancers = [];
        for (var i = 0; i < sortedPools.length; i++) {
            var winnerId = getPoolWinnerId(sortedPools[i].id);
            if (winnerId) {
                advancers.push({
                    participantId: winnerId,
                    name: getParticipantName(winnerId),
                    fromPool: sortedPools[i].number,
                    seed: i + 1
                });
            }
        }

        if (advancers.length < 2) {
            alert('Недостаточно победителей пулов для сетки.');
            return;
        }

        buildBracketFromAdvancers(advancers);
        renderBracket();
        updateTournamentBar();
        openBracketOverlay();
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

    function formatBracketFighter(fighter) {
        if (!fighter) return { text: 'Ожидание', className: 'tbd' };
        if (fighter.isBye) return { text: 'BYE', className: 'tbd' };
        return { text: fighter.name, className: '' };
    }

    function createBracketSlotEl(fighter, match) {
        var info = formatBracketFighter(fighter);
        var slot = document.createElement('div');
        slot.className = 'bracket-slot ' + info.className;
        if (match.winner && match.winner === fighter && fighter && !fighter.isBye) {
            slot.classList.add('winner-slot');
        }
        slot.textContent = info.text;
        slot.title = info.text;
        return slot;
    }

    function canFightBracketMatch(match) {
        return match.status === 'pending' && match.fighter1 && match.fighter2 &&
            !match.fighter1.isBye && !match.fighter2.isBye;
    }

    function createBracketMatchEl(match, loc, matchCount) {
        var card = document.createElement('div');
        card.className = 'bracket-match' + (match.status === 'done' ? ' done' : '');
        card.style.marginTop = getBracketMatchMarginTop(loc.roundIndex, loc.matchIndex, matchCount) + 'px';

        card.appendChild(createBracketSlotEl(match.fighter1, match));
        card.appendChild(createBracketSlotEl(match.fighter2, match));

        if (canFightBracketMatch(match)) {
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
            columnsEl.appendChild(col);
        }

        halfEl.appendChild(columnsEl);
        return halfEl;
    }

    function renderBracketCenter() {
        var center = document.createElement('div');
        center.className = 'bracket-center';

        var label = document.createElement('div');
        label.className = 'bracket-center-label';
        label.textContent = 'Финал';
        center.appendChild(label);

        var championWrap = document.createElement('div');
        championWrap.className = 'bracket-champion-slot';
        var championSlot = document.createElement('div');
        var champInfo = { text: 'Победитель', className: 'tbd champion' };
        if (gameState.bracket.final.status === 'done' && gameState.bracket.final.winner) {
            champInfo = formatBracketFighter(gameState.bracket.final.winner);
            champInfo.className += ' champion';
        }
        championSlot.className = 'bracket-slot ' + champInfo.className;
        championSlot.textContent = champInfo.text;
        championWrap.appendChild(championSlot);
        center.appendChild(championWrap);

        var finalWrap = document.createElement('div');
        finalWrap.className = 'bracket-final-match';
        finalWrap.appendChild(createBracketMatchEl(gameState.bracket.final, {
            side: 'final',
            roundIndex: 0,
            matchIndex: 0
        }, 1));
        center.appendChild(finalWrap);

        return center;
    }

    function renderBracket() {
        var container = document.getElementById('bracketContainer');
        container.innerHTML = '';

        if (!gameState.bracket) return;

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
        var match = getBracketMatch(loc);
        if (!match || match.status === 'done' || !canFightBracketMatch(match)) return;

        gameState.tournamentStage = 'bracket-fights';
        gameState.activeBracketMatch = loc;
        document.getElementById('bracketOverlay').style.display = 'none';

        resetFightState();
        gameState.redFighterName = match.fighter1.name;
        gameState.blueFighterName = match.fighter2.name;
        gameState.sessionStarted = true;
        document.getElementById('secretaryTerminal').classList.remove('hidden');
        updateTournamentBar();
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

        advanceBracketWinner(loc, winner);
        gameState.activeBracketMatch = null;
        gameState.tournamentStage = 'bracket';
        updateTournamentBar();
        renderBracket();
        openBracketOverlay();

        if (isBracketComplete()) {
            alert('Турнир завершён! Победитель: ' + gameState.bracket.final.winner.name);
        }
        return true;
    }

    function advanceBracketWinner(loc, winner) {
        if (!winner || !gameState.bracket) return;
        advanceBracketWinnerSilent(loc, winner);
        resolveBracketByes();
    }

    function isBracketComplete() {
        if (!gameState.bracket || !gameState.bracket.final) return false;
        return gameState.bracket.final.status === 'done';
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
        var poolsOverlay = document.getElementById('poolsOverlay');
        if (poolsOverlay) poolsOverlay.style.display = 'none';
        hidePoolSelectModal();
        hideBracketOverlay();
        updateFormBracketButtons();
        updateOpenBracketButton();
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
        onSaveFightHook: onSaveFightHook
    };

    global.triggerImportParticipantsXlsx = triggerImportParticipantsXlsx;
    global.handleParticipantsXlsxSelected = handleParticipantsXlsxSelected;
    global.previewBracketDemo = previewBracketDemo;
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
