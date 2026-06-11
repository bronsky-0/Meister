/* Швейцарская система: круги с ручным составлением пар */
(function(global) {
    'use strict';

    var draggedSwissParticipantId = null;
    var gameState = function() { return global.gameState; };

    function gs() {
        return gameState();
    }

    function isSwissTournament() {
        return gs().sessionMode === 'tournament' && gs().tournamentSystem === 'swiss';
    }

    function isSwissSessionActive() {
        return isSwissTournament() && !!gs().swissStarted;
    }

    function getParticipantById(id) {
        var list = gs().participants || [];
        for (var i = 0; i < list.length; i++) {
            if (list[i].id === id) return list[i];
        }
        return null;
    }

    function getParticipantName(id) {
        var p = getParticipantById(id);
        return p ? p.name : '—';
    }

    function generateMatchId() {
        return 'sw_m_' + Date.now() + '_' + Math.floor(Math.random() * 10000);
    }

    function ensureSwissRounds() {
        if (!gs().swissRounds) gs().swissRounds = [];
    }

    function getCurrentSwissRoundNumber() {
        return gs().currentSwissRound || 1;
    }

    function findSwissRound(number) {
        ensureSwissRounds();
        var rounds = gs().swissRounds;
        for (var i = 0; i < rounds.length; i++) {
            if (rounds[i].number === number) return rounds[i];
        }
        return null;
    }

    function ensureSwissRound(number) {
        var round = findSwissRound(number);
        if (round) return round;
        round = { number: number, status: 'setup', matches: [] };
        gs().swissRounds.push(round);
        gs().swissRounds.sort(function(a, b) { return a.number - b.number; });
        return round;
    }

    function getCurrentSwissRound() {
        return ensureSwissRound(getCurrentSwissRoundNumber());
    }

    function getAssignedParticipantIdsInRound(round) {
        var ids = {};
        var matches = round.matches || [];
        for (var i = 0; i < matches.length; i++) {
            var m = matches[i];
            if (m.fighter1Id) ids[m.fighter1Id] = true;
            if (m.fighter2Id) ids[m.fighter2Id] = true;
        }
        return ids;
    }

    function getUnassignedParticipantIdsForRound(round) {
        var assigned = getAssignedParticipantIdsInRound(round);
        return (gs().participants || [])
            .filter(function(p) { return !assigned[p.id]; })
            .map(function(p) { return p.id; });
    }

    function findSwissMatch(matchId) {
        ensureSwissRounds();
        var rounds = gs().swissRounds;
        for (var r = 0; r < rounds.length; r++) {
            var matches = rounds[r].matches || [];
            for (var m = 0; m < matches.length; m++) {
                if (matches[m].id === matchId) return matches[m];
            }
        }
        return null;
    }

    function getAllSwissMatchesFlat() {
        var out = [];
        ensureSwissRounds();
        var rounds = gs().swissRounds;
        for (var r = 0; r < rounds.length; r++) {
            var matches = rounds[r].matches || [];
            for (var m = 0; m < matches.length; m++) {
                if (matches[m].status === 'done') out.push(matches[m]);
            }
        }
        return out;
    }

    function getSwissStandings() {
        if (typeof PoolRankings === 'undefined' || !PoolRankings.buildPoolRankingsList) {
            return [];
        }
        var poolMatches = { swiss: [] };
        var done = getAllSwissMatchesFlat();
        for (var i = 0; i < done.length; i++) {
            poolMatches.swiss.push(done[i]);
        }
        return PoolRankings.buildPoolRankingsList(gs().participants || [], poolMatches);
    }

    function isCurrentSwissRoundComplete() {
        var round = getCurrentSwissRound();
        var matches = (round.matches || []).filter(function(m) {
            return m.status !== 'draft' && m.fighter1Id && m.fighter2Id;
        });
        if (!matches.length) return false;
        for (var i = 0; i < matches.length; i++) {
            if (matches[i].status !== 'done') return false;
        }
        return true;
    }

    function getPendingSwissMatchesInCurrentRound() {
        var round = getCurrentSwissRound();
        return (round.matches || []).filter(function(m) {
            return m.status === 'pending' && m.fighter1Id && m.fighter2Id;
        });
    }

    function persistIfAny() {
        if (typeof global.persistActiveTournamentNominationIfAny === 'function') {
            global.persistActiveTournamentNominationIfAny();
        }
    }

    function escapeHtml(text) {
        return String(text)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function updateSwissRoundTitle() {
        var title = document.getElementById('swissRoundTitle');
        if (title) title.textContent = getCurrentSwissRoundNumber() + ' круг';
        var fightsTitle = document.getElementById('swissFightsRoundTitle');
        if (fightsTitle) fightsTitle.textContent = getCurrentSwissRoundNumber() + ' круг — бои';
    }

    function goToSwissRoundSetup() {
        if (!isSwissTournament()) {
            alert('Доступно только для системы «Швейцарская система».');
            return;
        }
        if (gs().participants.length < 2) {
            alert('Добавьте минимум двух участников.');
            return;
        }

        if (gs().swissStarted && gs().tournamentStage === 'swiss-fights') {
            showSwissFightsOverlay();
            return;
        }

        gs().tournamentStage = 'swiss-setup';
        if (!gs().currentSwissRound) gs().currentSwissRound = 1;
        var round = getCurrentSwissRound();
        if (!round.matches.length) {
            addSwissPair();
        }

        document.getElementById('startMenuOverlay').style.display = 'none';
        document.getElementById('swissFightsOverlay').style.display = 'none';
        document.getElementById('swissOverlay').style.display = 'flex';
        updateSwissRoundTitle();
        renderSwissSetup();
        updateSwissFooterButtons();
        var startBtn = document.getElementById('swissStartRoundBtn');
        if (startBtn) {
            startBtn.textContent = gs().swissStarted ? 'Начать круг' : 'Начать турнир';
        }
        persistIfAny();
    }

    function backFromSwissToParticipants() {
        if (gs().swissStarted) {
            if (!confirm('Вернуться к списку участников? Текущий прогресс круга сохранён.')) return;
        }
        document.getElementById('swissOverlay').style.display = 'none';
        document.getElementById('swissFightsOverlay').style.display = 'none';
        document.getElementById('startMenuOverlay').style.display = 'flex';
        if (typeof showParticipantsPanel === 'function') showParticipantsPanel();
    }

    function addSwissPair() {
        var round = getCurrentSwissRound();
        round.matches.push({
            id: generateMatchId(),
            roundNumber: round.number,
            fighter1Id: null,
            fighter2Id: null,
            status: 'draft',
            winnerId: null,
            result: null,
            arenaId: null
        });
        renderSwissSetup();
        persistIfAny();
    }

    function removeSwissPair(matchId) {
        var round = getCurrentSwissRound();
        round.matches = (round.matches || []).filter(function(m) { return m.id !== matchId; });
        renderSwissSetup();
        persistIfAny();
    }

    function assignParticipantToSwissSlot(matchId, slot, participantId) {
        var match = findSwissMatch(matchId);
        if (!match || match.status !== 'draft') return;

        var round = getCurrentSwissRound();
        var assigned = getAssignedParticipantIdsInRound(round);
        if (participantId && assigned[participantId]) {
            var matches = round.matches || [];
            for (var i = 0; i < matches.length; i++) {
                var m = matches[i];
                if (m.id === matchId) continue;
                if (m.fighter1Id === participantId) m.fighter1Id = null;
                if (m.fighter2Id === participantId) m.fighter2Id = null;
            }
        }

        if (slot === 1) match.fighter1Id = participantId || null;
        else match.fighter2Id = participantId || null;

        renderSwissSetup();
        persistIfAny();
    }

    function renderSwissUnassignedList() {
        var list = document.getElementById('swissUnassignedList');
        if (!list) return;
        list.innerHTML = '';

        var round = getCurrentSwissRound();
        var unassigned = getUnassignedParticipantIdsForRound(round);

        if (!unassigned.length) {
            var empty = document.createElement('p');
            empty.className = 'participants-hint';
            empty.textContent = 'Все участники распределены по парам';
            list.appendChild(empty);
            return;
        }

        var ul = document.createElement('ul');
        ul.className = 'participants-list';
        ul.style.maxHeight = 'none';
        ul.style.margin = '0';

        for (var i = 0; i < unassigned.length; i++) {
            (function(pid) {
                var p = getParticipantById(pid);
                if (!p) return;
                var li = document.createElement('li');
                li.className = 'participant-item';
                li.draggable = true;
                li.dataset.id = pid;
                li.innerHTML =
                    '<span class="participant-drag-handle">☰</span>' +
                    '<span class="participant-name">' + escapeHtml(p.name) + '</span>';
                li.addEventListener('dragstart', onSwissParticipantDragStart);
                li.addEventListener('dragend', onSwissParticipantDragEnd);
                ul.appendChild(li);
            })(unassigned[i]);
        }
        list.appendChild(ul);
    }

    function renderSwissSetup() {
        var container = document.getElementById('swissPairsContainer');
        if (!container) return;
        container.innerHTML = '';

        var round = getCurrentSwissRound();
        var matches = round.matches || [];

        if (!matches.length) {
            container.innerHTML = '<p class="pools-empty">Добавьте пары для этого круга</p>';
            renderSwissUnassignedList();
            return;
        }

        for (var i = 0; i < matches.length; i++) {
            (function(match, index) {
                var card = document.createElement('div');
                card.className = 'pool-card swiss-pair-card';

                var header = document.createElement('div');
                header.className = 'pool-header';
                header.textContent = 'Пара ' + (index + 1);
                card.appendChild(header);

                var table = document.createElement('table');
                table.className = 'pool-table';

                for (var slot = 1; slot <= 2; slot++) {
                    var fid = slot === 1 ? match.fighter1Id : match.fighter2Id;
                    var tr = document.createElement('tr');
                    tr.className = 'pool-participant-row swiss-pair-slot';
                    tr.dataset.matchId = match.id;
                    tr.dataset.slot = String(slot);

                    var name = fid ? getParticipantName(fid) : 'Перетащите участника';
                    tr.innerHTML =
                        '<td class="pool-row-num">' + slot + '</td>' +
                        '<td class="' + (fid ? '' : 'pool-empty-row') + '">' + escapeHtml(name) + '</td>';

                    tr.addEventListener('dragover', onSwissSlotDragOver);
                    tr.addEventListener('dragleave', onSwissSlotDragLeave);
                    tr.addEventListener('drop', onSwissSlotDrop);
                    table.appendChild(tr);
                }
                card.appendChild(table);

                var removeBtn = document.createElement('button');
                removeBtn.type = 'button';
                removeBtn.className = 'pool-remove-btn';
                removeBtn.textContent = 'Удалить пару';
                removeBtn.onclick = function() { removeSwissPair(match.id); };
                card.appendChild(removeBtn);

                container.appendChild(card);
            })(matches[i], i);
        }

        renderSwissUnassignedList();
    }

    function onSwissParticipantDragStart(event) {
        draggedSwissParticipantId = event.currentTarget.dataset.id;
        event.currentTarget.classList.add('dragging');
        event.dataTransfer.effectAllowed = 'move';
    }

    function onSwissParticipantDragEnd(event) {
        event.currentTarget.classList.remove('dragging');
        draggedSwissParticipantId = null;
        var slots = document.querySelectorAll('.swiss-pair-slot');
        for (var i = 0; i < slots.length; i++) slots[i].classList.remove('drag-over');
    }

    function onSwissSlotDragOver(event) {
        event.preventDefault();
        event.currentTarget.classList.add('drag-over');
    }

    function onSwissSlotDragLeave(event) {
        event.currentTarget.classList.remove('drag-over');
    }

    function onSwissSlotDrop(event) {
        event.preventDefault();
        event.currentTarget.classList.remove('drag-over');
        if (!draggedSwissParticipantId) return;
        var matchId = event.currentTarget.dataset.matchId;
        var slot = parseInt(event.currentTarget.dataset.slot, 10);
        assignParticipantToSwissSlot(matchId, slot, draggedSwissParticipantId);
        draggedSwissParticipantId = null;
    }

    function validateSwissRoundForStart() {
        var round = getCurrentSwissRound();
        var valid = (round.matches || []).filter(function(m) {
            return m.fighter1Id && m.fighter2Id;
        });
        if (!valid.length) {
            alert('Составьте хотя бы одну пару из двух участников.');
            return false;
        }

        var assigned = {};
        for (var i = 0; i < valid.length; i++) {
            var m = valid[i];
            if (assigned[m.fighter1Id] || assigned[m.fighter2Id]) {
                alert('Один участник не может быть в нескольких парах одного круга.');
                return false;
            }
            assigned[m.fighter1Id] = true;
            assigned[m.fighter2Id] = true;
        }
        return true;
    }

    function startSwissRoundOrTournament() {
        if (!validateSwissRoundForStart()) return;

        var round = getCurrentSwissRound();
        round.matches = (round.matches || []).filter(function(m) {
            return m.fighter1Id && m.fighter2Id;
        });
        for (var i = 0; i < round.matches.length; i++) {
            round.matches[i].status = 'pending';
            round.matches[i].winnerId = null;
            round.matches[i].result = null;
        }
        round.status = 'active';

        gs().swissStarted = true;
        gs().sessionStarted = true;
        gs().tournamentMode = true;
        gs().tournamentStage = 'swiss-fights';
        gs().networkActive = true;

        if (gs().activeTournamentId && gs().activeNominationId &&
            typeof MeisterTournaments !== 'undefined' &&
            MeisterTournaments.setNominationNetworkActive) {
            MeisterTournaments.setNominationNetworkActive(
                gs().activeTournamentId,
                gs().activeNominationId,
                true
            );
        }

        persistIfAny();
        document.getElementById('swissOverlay').style.display = 'none';
        showSwissFightsOverlay();
    }

    function showSwissFightsOverlay() {
        document.getElementById('swissFightsOverlay').style.display = 'flex';
        updateSwissRoundTitle();
        renderSwissFightList();
        updateSwissFooterButtons();
        updateSwissTournamentBar();
    }

    function hideSwissFightsOverlay() {
        document.getElementById('swissFightsOverlay').style.display = 'none';
    }

    function getArenaOptions() {
        if (typeof NetworkSync !== 'undefined' && NetworkSync.getState) {
            var st = NetworkSync.getState();
            if (st.arenas && st.arenas.length) return st.arenas;
        }
        var count = 4;
        var arenas = [];
        for (var i = 1; i <= count; i++) arenas.push({ id: i, name: 'Площадка ' + i });
        return arenas;
    }

    function populateSwissArenaSelect(selectEl, selectedId) {
        if (!selectEl) return;
        var arenas = getArenaOptions();
        selectEl.innerHTML = '<option value="">—</option>';
        for (var i = 0; i < arenas.length; i++) {
            var opt = document.createElement('option');
            opt.value = String(arenas[i].id);
            opt.textContent = arenas[i].name || ('Площадка ' + arenas[i].id);
            if (selectedId && parseInt(selectedId, 10) === arenas[i].id) opt.selected = true;
            selectEl.appendChild(opt);
        }
    }

    function onSwissMatchArenaChanged(matchId, arenaId) {
        var match = findSwissMatch(matchId);
        if (!match) return;
        match.arenaId = arenaId ? parseInt(arenaId, 10) : null;
        persistIfAny();
    }

    function renderSwissFightList() {
        var list = document.getElementById('swissFightList');
        if (!list) return;
        list.innerHTML = '';

        var round = getCurrentSwissRound();
        var matches = round.matches || [];

        if (!matches.length) {
            list.innerHTML = '<li class="participants-hint">Нет боёв в этом круге</li>';
            return;
        }

        for (var i = 0; i < matches.length; i++) {
            (function(match) {
                var li = document.createElement('li');
                li.className = 'pool-fight-item';
                if (match.status === 'done') li.classList.add('done');
                if (match.status === 'in_progress') li.classList.add('in-progress');

                var row = document.createElement('div');
                row.className = 'swiss-fight-row';

                var info = document.createElement('div');
                info.className = 'pool-fight-item-info';
                var mark = match.status === 'done' ? '✓' :
                    match.status === 'in_progress' ? '▶' : '○';
                info.textContent = mark + ' ' +
                    getParticipantName(match.fighter1Id) + ' vs ' +
                    getParticipantName(match.fighter2Id);
                row.appendChild(info);

                var arenaWrap = document.createElement('div');
                arenaWrap.className = 'swiss-fight-arena';
                var label = document.createElement('label');
                label.textContent = 'Площадка';
                var select = document.createElement('select');
                select.className = 'swiss-arena-select';
                populateSwissArenaSelect(select, match.arenaId);
                select.onchange = function() {
                    onSwissMatchArenaChanged(match.id, select.value);
                };
                arenaWrap.appendChild(label);
                arenaWrap.appendChild(select);
                row.appendChild(arenaWrap);

                if (match.status !== 'done') {
                    var btn = document.createElement('button');
                    btn.type = 'button';
                    btn.className = 'menu-button';
                    btn.textContent = match.status === 'in_progress' ? 'Продолжить' : 'Провести';
                    btn.onclick = function() {
                        startSwissFight(match.id);
                    };
                    row.appendChild(btn);
                } else {
                    var editBtn = document.createElement('button');
                    editBtn.type = 'button';
                    editBtn.className = 'menu-button';
                    editBtn.textContent = 'Изменить';
                    editBtn.onclick = function() {
                        startSwissFight(match.id, { reopen: true });
                    };
                    row.appendChild(editBtn);
                }

                li.appendChild(row);
                list.appendChild(li);
            })(matches[i]);
        }
    }

    function loadSwissMatchIntoFight(match, options) {
        options = options || {};
        if (typeof resetFightState === 'function') resetFightState();
        gs().redFighterName = getParticipantName(match.fighter1Id);
        gs().blueFighterName = getParticipantName(match.fighter2Id);
        gs().activeSwissMatchId = match.id;
        gs().activeSwissMatchMeta = {
            matchId: match.id,
            roundNumber: match.roundNumber,
            fighter1Id: match.fighter1Id,
            fighter2Id: match.fighter2Id
        };
        if (options.restoreResult && match.result && typeof restoreFightScoresFromMatchResult === 'undefined') {
            gs().redScore = match.result.redScore || 0;
            gs().blueScore = match.result.blueScore || 0;
            gs().redBonus = match.result.redBonus || 0;
            gs().blueBonus = match.result.blueBonus || 0;
            gs().redRoundsWon = match.result.redRoundsWon || 0;
            gs().blueRoundsWon = match.result.blueRoundsWon || 0;
        } else if (options.restoreResult && match.result) {
            gs().redScore = match.result.redScore || 0;
            gs().blueScore = match.result.blueScore || 0;
            gs().redBonus = match.result.redBonus || 0;
            gs().blueBonus = match.result.blueBonus || 0;
            gs().redRoundsWon = match.result.redRoundsWon || 0;
            gs().blueRoundsWon = match.result.blueRoundsWon || 0;
        }
        if (!options.reopen) {
            match.status = 'in_progress';
        }
    }

    function startSwissFight(matchId, options) {
        options = options || {};
        var match = findSwissMatch(matchId);
        if (!match) return;

        gs().tournamentStage = 'swiss-fights';
        loadSwissMatchIntoFight(match, options);
        hideSwissFightsOverlay();
        document.getElementById('secretaryTerminal').classList.remove('hidden');
        updateSwissTournamentBar();
        if (typeof updateDisplay === 'function') updateDisplay();
        persistIfAny();
    }

    function saveActiveSwissMatchResult() {
        if (!gs().activeSwissMatchId || !gs().activeSwissMatchMeta) return false;

        var match = findSwissMatch(gs().activeSwissMatchId);
        if (!match) return false;

        var winnerSide = typeof getRoundWinner === 'function' ? getRoundWinner() : 'draw';
        var winnerId = null;
        if (winnerSide === 'red') winnerId = match.fighter1Id;
        else if (winnerSide === 'blue') winnerId = match.fighter2Id;

        match.status = 'done';
        match.winnerId = winnerId;
        match.result = {
            redScore: gs().redScore,
            blueScore: gs().blueScore,
            redBonus: gs().redBonus,
            blueBonus: gs().blueBonus,
            redRoundsWon: gs().redRoundsWon,
            blueRoundsWon: gs().blueRoundsWon,
            winnerSide: winnerSide,
            date: new Date().toLocaleString('ru-RU')
        };

        gs().tournamentFightHistory.push({
            type: 'swiss',
            roundNumber: match.roundNumber,
            matchId: match.id,
            fighter1: getParticipantName(match.fighter1Id),
            fighter2: getParticipantName(match.fighter2Id),
            winnerId: winnerId,
            result: match.result
        });

        gs().activeSwissMatchId = null;
        gs().activeSwissMatchMeta = null;

        var round = findSwissRound(match.roundNumber);
        if (round && isCurrentSwissRoundComplete()) {
            round.status = 'complete';
        }

        updateSwissTournamentBar();
        persistIfAny();
        if (document.getElementById('swissRankingsOverlay') &&
            document.getElementById('swissRankingsOverlay').style.display === 'flex') {
            renderSwissRankingsTable();
        }

        if (isCurrentSwissRoundComplete()) {
            showSwissFightsOverlay();
            alert('Все бои круга завершены. Можно перейти к следующему кругу или подвести итоги.');
        } else {
            if (confirm('Бой сохранён. Вернуться к списку боёв круга?')) {
                showSwissFightsOverlay();
            }
        }
        return true;
    }

    function updateSwissFooterButtons() {
        var nextBtn = document.getElementById('swissNextRoundBtn');
        var finishBtn = document.getElementById('swissFinishBtn');
        var startBtn = document.getElementById('swissStartRoundBtn');
        var complete = isSwissSessionActive() && isCurrentSwissRoundComplete();

        if (nextBtn) {
            if (complete && gs().tournamentStage !== 'swiss-finished') {
                nextBtn.classList.remove('hidden');
            } else {
                nextBtn.classList.add('hidden');
            }
        }
        if (finishBtn) {
            if (isSwissSessionActive()) finishBtn.classList.remove('hidden');
            else finishBtn.classList.add('hidden');
        }
        if (startBtn) {
            startBtn.textContent = gs().swissStarted ? 'Начать круг' : 'Начать турнир';
        }
    }

    function goToNextSwissRound() {
        if (!isCurrentSwissRoundComplete()) {
            alert('Сначала проведите все бои текущего круга.');
            return;
        }
        gs().currentSwissRound = getCurrentSwissRoundNumber() + 1;
        gs().tournamentStage = 'swiss-setup';
        var round = ensureSwissRound(gs().currentSwissRound);
        round.status = 'setup';
        round.matches = [];
        addSwissPair();

        document.getElementById('swissFightsOverlay').style.display = 'none';
        document.getElementById('swissOverlay').style.display = 'flex';
        updateSwissRoundTitle();
        renderSwissSetup();
        updateSwissFooterButtons();
        var startBtn = document.getElementById('swissStartRoundBtn');
        if (startBtn) startBtn.textContent = 'Начать круг';
        persistIfAny();
    }

    function finishSwissTournament() {
        if (!confirm('Завершить номинацию и показать итоговый рейтинг?')) return;
        gs().tournamentStage = 'swiss-finished';
        var round = getCurrentSwissRound();
        if (round && isCurrentSwissRoundComplete()) round.status = 'complete';
        hideSwissFightsOverlay();
        document.getElementById('swissOverlay').style.display = 'none';
        showSwissRankingsOverlay();
        updateSwissTournamentBar();
        persistIfAny();
    }

    function renderSwissRankingsTable() {
        var tbody = document.getElementById('swissRankingsTableBody');
        if (!tbody) return;
        tbody.innerHTML = '';

        var rows = getSwissStandings();
        if (!rows.length) {
            tbody.innerHTML = '<tr><td colspan="6" class="pool-rankings-empty">Нет результатов</td></tr>';
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

    function showSwissRankingsOverlay() {
        renderSwissRankingsTable();
        var overlay = document.getElementById('swissRankingsOverlay');
        if (overlay) overlay.style.display = 'flex';
    }

    function hideSwissRankingsOverlay() {
        var overlay = document.getElementById('swissRankingsOverlay');
        if (overlay) overlay.style.display = 'none';
    }

    function showSwissRankingsFromToolbar() {
        renderSwissRankingsTable();
        showSwissRankingsOverlay();
    }

    function updateSwissTournamentBar() {
        var bar = document.getElementById('tournamentBar');
        var info = document.getElementById('tournamentBarInfo');
        if (!bar || !info) return;

        if (!isSwissSessionActive()) {
            if (!isSwissTournament() || gs().tournamentStage !== 'swiss-finished') {
                return;
            }
        }

        bar.classList.remove('hidden');

        var text = gs().activeNominationName || 'Швейцарская система';
        text += ' · ' + getCurrentSwissRoundNumber() + ' круг';
        if (gs().activeSwissMatchMeta) {
            text += ' · ' + gs().redFighterName + ' vs ' + gs().blueFighterName;
        } else if (isCurrentSwissRoundComplete()) {
            text += ' · Круг завершён';
        }
        info.textContent = text;

        var poolBtn = document.getElementById('selectPoolBtn');
        if (poolBtn && isSwissSessionActive()) {
            poolBtn.classList.remove('hidden');
            poolBtn.textContent = 'Бои круга';
            poolBtn.onclick = function() { showSwissFightsOverlay(); };
        }

        var rankBtn = document.getElementById('showRankingsBtnBar');
        if (rankBtn && isSwissSessionActive()) {
            rankBtn.classList.remove('hidden');
            rankBtn.onclick = function() { showSwissRankingsFromToolbar(); };
        }

        var playoffOnly = ['formBracketBtn', 'randomPoolResultsBtn', 'openBracketBtn'];
        for (var i = 0; i < playoffOnly.length; i++) {
            var el = document.getElementById(playoffOnly[i]);
            if (el) el.classList.add('hidden');
        }

        updateSwissFooterButtons();
    }

    function swissGoBack() {
        if (gs().activeSwissMatchId) {
            if (!confirm('Выйти из боя? Несохранённый результат будет потерян.')) return;
            var match = findSwissMatch(gs().activeSwissMatchId);
            if (match && match.status === 'in_progress' && !match.result) {
                match.status = 'pending';
            }
            gs().activeSwissMatchId = null;
            gs().activeSwissMatchMeta = null;
        }
        showSwissFightsOverlay();
    }

    function resetSwissState() {
        gs().swissRounds = [];
        gs().currentSwissRound = 1;
        gs().swissStarted = false;
        gs().activeSwissMatchId = null;
        gs().activeSwissMatchMeta = null;
        if (gs().tournamentSystem === 'swiss') {
            gs().tournamentStage = null;
        }
        var overlay = document.getElementById('swissOverlay');
        if (overlay) overlay.style.display = 'none';
        hideSwissFightsOverlay();
        hideSwissRankingsOverlay();
    }

    function applySwissFromSnapshot(snapshot) {
        if (!snapshot) return;
        gs().swissRounds = snapshot.swissRounds
            ? JSON.parse(JSON.stringify(snapshot.swissRounds))
            : [];
        gs().currentSwissRound = snapshot.currentSwissRound || 1;
        gs().swissStarted = !!snapshot.swissStarted;
    }

    function onSaveFightHook() {
        if (!isSwissTournament()) return false;
        if (gs().activeSwissMatchId && gs().activeSwissMatchMeta) {
            return saveActiveSwissMatchResult();
        }
        return false;
    }

    function updateParticipantsPanelForSystem() {
        var btn = document.querySelector('#participantsPanel .participants-actions-row .menu-button.primary');
        if (!btn) return;
        if (gs().tournamentSystem === 'swiss') {
            btn.textContent = 'Перейти к 1 кругу';
            btn.onclick = function() { goToSwissRoundSetup(); };
        } else {
            btn.textContent = 'Перейти к составлению пулов';
            btn.onclick = function() {
                if (typeof goToPoolsComposition === 'function') goToPoolsComposition();
            };
        }
    }

    function initSwissUI() {
        updateParticipantsPanelForSystem();
    }

    global.TournamentSwiss = {
        init: initSwissUI,
        isSwiss: isSwissTournament,
        isSwissSessionActive: isSwissSessionActive,
        goToSwissRoundSetup: goToSwissRoundSetup,
        backFromSwissToParticipants: backFromSwissToParticipants,
        addSwissPair: addSwissPair,
        startSwissRoundOrTournament: startSwissRoundOrTournament,
        showSwissFightsOverlay: showSwissFightsOverlay,
        hideSwissFightsOverlay: hideSwissFightsOverlay,
        goToNextSwissRound: goToNextSwissRound,
        finishSwissTournament: finishSwissTournament,
        showSwissRankingsFromToolbar: showSwissRankingsFromToolbar,
        hideSwissRankingsOverlay: hideSwissRankingsOverlay,
        showSwissRankingsOverlay: showSwissRankingsOverlay,
        updateSwissTournamentBar: updateSwissTournamentBar,
        updateParticipantsPanelForSystem: updateParticipantsPanelForSystem,
        swissGoBack: swissGoBack,
        resetSwissState: resetSwissState,
        onSaveFightHook: onSaveFightHook,
        applySwissFromSnapshot: applySwissFromSnapshot,
        getSwissStandings: getSwissStandings
    };

    global.goToSwissRoundSetup = goToSwissRoundSetup;
    global.backFromSwissToParticipants = backFromSwissToParticipants;
    global.addSwissPair = addSwissPair;
    global.startSwissRoundOrTournament = startSwissRoundOrTournament;
    global.goToNextSwissRound = goToNextSwissRound;
    global.finishSwissTournament = finishSwissTournament;
    global.showSwissRankingsFromToolbar = showSwissRankingsFromToolbar;
    global.hideSwissRankingsOverlay = hideSwissRankingsOverlay;
    global.swissGoBack = swissGoBack;
    global.hideSwissFightsOverlay = hideSwissFightsOverlay;
})(window);
