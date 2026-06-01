/* Сохранённые турниры и номинации (localStorage) */
(function(global) {
    'use strict';

    var STORAGE_KEY = 'meisterSavedTournaments';
    var APP_VERSION = '1.0.0';

    function generateId(prefix) {
        return prefix + '_' + Date.now() + '_' + Math.floor(Math.random() * 10000);
    }

    function loadAll() {
        try {
            var raw = localStorage.getItem(STORAGE_KEY);
            if (!raw) return [];
            var list = JSON.parse(raw);
            return Array.isArray(list) ? list : [];
        } catch (e) {
            return [];
        }
    }

    function saveAll(list) {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(list));
    }

    function findTournament(tournamentId) {
        var list = loadAll();
        for (var i = 0; i < list.length; i++) {
            if (list[i].id === tournamentId) return { tournament: list[i], index: i, list: list };
        }
        return null;
    }

    function findNomination(tournament, nominationId) {
        if (!tournament || !tournament.nominations) return null;
        for (var i = 0; i < tournament.nominations.length; i++) {
            if (tournament.nominations[i].id === nominationId) {
                return tournament.nominations[i];
            }
        }
        return null;
    }

    function createEmptyNominationState(name, ruleset) {
        return {
            id: generateId('nom'),
            name: name,
            ruleset: ruleset,
            tournamentSystem: null,
            participants: [],
            tournamentStage: null,
            pools: [],
            poolMatches: {},
            bracket: null,
            playoffStarted: false,
            qualifyingAdvancersCount: null,
            tournamentFightHistory: []
        };
    }

    function createTournament(name, nominationRows) {
        var now = new Date().toISOString();
        var nominations = (nominationRows || []).map(function(row) {
            return createEmptyNominationState(row.name, row.ruleset);
        });
        return {
            id: generateId('t'),
            name: name,
            createdAt: now,
            updatedAt: now,
            nominations: nominations
        };
    }

    function addTournament(tournament) {
        var list = loadAll();
        list.unshift(tournament);
        saveAll(list);
        return tournament;
    }

    function updateTournament(tournament) {
        var found = findTournament(tournament.id);
        if (!found) return false;
        tournament.updatedAt = new Date().toISOString();
        found.list[found.index] = tournament;
        saveAll(found.list);
        return true;
    }

    function extractNominationState(gs) {
        return {
            tournamentSystem: gs.tournamentSystem || null,
            participants: gs.participants ? gs.participants.slice() : [],
            tournamentStage: gs.tournamentStage || null,
            pools: gs.pools ? JSON.parse(JSON.stringify(gs.pools)) : [],
            poolMatches: gs.poolMatches ? JSON.parse(JSON.stringify(gs.poolMatches)) : {},
            bracket: gs.bracket ? JSON.parse(JSON.stringify(gs.bracket)) : null,
            playoffStarted: !!gs.playoffStarted,
            qualifyingAdvancersCount: gs.qualifyingAdvancersCount != null ? gs.qualifyingAdvancersCount : null,
            tournamentFightHistory: gs.tournamentFightHistory ? gs.tournamentFightHistory.slice() : [],
            poolArenaAssignments: gs.poolArenaAssignments ? JSON.parse(JSON.stringify(gs.poolArenaAssignments)) : {}
        };
    }

    function applyNominationToGameState(gs, nomination) {
        if (!nomination) return;
        gs.ruleset = nomination.ruleset;
        gs.tournamentSystem = nomination.tournamentSystem || null;
        gs.participants = nomination.participants ? JSON.parse(JSON.stringify(nomination.participants)) : [];
        gs.tournamentStage = nomination.tournamentStage || null;
        gs.pools = nomination.pools ? JSON.parse(JSON.stringify(nomination.pools)) : [];
        gs.poolMatches = nomination.poolMatches ? JSON.parse(JSON.stringify(nomination.poolMatches)) : {};
        gs.bracket = nomination.bracket ? JSON.parse(JSON.stringify(nomination.bracket)) : null;
        gs.playoffStarted = !!nomination.playoffStarted;
        gs.qualifyingAdvancersCount = nomination.qualifyingAdvancersCount != null
            ? nomination.qualifyingAdvancersCount
            : null;
        gs.tournamentFightHistory = nomination.tournamentFightHistory
            ? nomination.tournamentFightHistory.slice()
            : [];
        gs.poolArenaAssignments = nomination.poolArenaAssignments
            ? JSON.parse(JSON.stringify(nomination.poolArenaAssignments))
            : {};
        gs.tournamentMode = !!nomination.playoffStarted;
    }

    function persistActiveNominationFromGameState(gs) {
        if (!gs.activeTournamentId || !gs.activeNominationId) return false;
        var found = findTournament(gs.activeTournamentId);
        if (!found) return false;

        var nomination = findNomination(found.tournament, gs.activeNominationId);
        if (!nomination) return false;

        var patch = extractNominationState(gs);
        nomination.tournamentSystem = patch.tournamentSystem;
        nomination.participants = patch.participants;
        nomination.tournamentStage = patch.tournamentStage;
        nomination.pools = patch.pools;
        nomination.poolMatches = patch.poolMatches;
        nomination.bracket = patch.bracket;
        nomination.playoffStarted = patch.playoffStarted;
        nomination.qualifyingAdvancersCount = patch.qualifyingAdvancersCount;
        nomination.tournamentFightHistory = patch.tournamentFightHistory;
        nomination.poolArenaAssignments = patch.poolArenaAssignments
            ? JSON.parse(JSON.stringify(patch.poolArenaAssignments))
            : {};

        updateTournament(found.tournament);
        return true;
    }

    function getNominationProgressLabel(nomination) {
        if (!nomination.tournamentSystem) return 'Не настроена';
        if (nomination.playoffStarted && nomination.bracket && nomination.bracket.final &&
            nomination.bracket.final.winner) {
            return 'Завершена';
        }
        if (nomination.playoffStarted) return 'Идёт турнир';
        if (nomination.pools && nomination.pools.length) return 'Пулы';
        if (nomination.participants && nomination.participants.length >= 2) return 'Участники';
        return 'Настройка';
    }

    function getResumeAction(nomination) {
        if (!nomination.tournamentSystem) return 'system';
        if (nomination.playoffStarted) {
            if (nomination.bracket) return 'bracket';
            return 'pool-select';
        }
        if (nomination.pools && nomination.pools.length) return 'pools';
        if (nomination.participants && nomination.participants.length >= 2) return 'participants';
        return 'network';
    }

    global.MeisterTournaments = {
        APP_VERSION: APP_VERSION,
        loadAll: loadAll,
        createTournament: createTournament,
        addTournament: addTournament,
        findTournament: function(id) {
            var found = findTournament(id);
            return found ? found.tournament : null;
        },
        updateTournament: updateTournament,
        findNomination: findNomination,
        applyNominationToGameState: applyNominationToGameState,
        persistActiveNominationFromGameState: persistActiveNominationFromGameState,
        getNominationProgressLabel: getNominationProgressLabel,
        getResumeAction: getResumeAction,
        createEmptyNominationState: createEmptyNominationState
    };

    global.persistActiveTournamentNominationIfAny = function() {
        if (typeof MeisterTournaments !== 'undefined') {
            MeisterTournaments.persistActiveNominationFromGameState(global.gameState);
        }
    };
})(window);
