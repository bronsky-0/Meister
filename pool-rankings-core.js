/**
 * Рейтинг участников после этапа пулов.
 * Сортировка: победы → поражения → разница очков (нанёс − получил).
 */
(function(root, factory) {
    var api = factory();
    if (typeof module === 'object' && module.exports) {
        module.exports = api;
    }
    if (root) {
        root.PoolRankings = api;
    }
}(typeof globalThis !== 'undefined' ? globalThis : this, function() {
    'use strict';

    function fighterPointsFromResult(result, isFighter1) {
        if (!result) return { scored: 0, received: 0 };
        var red = (result.redScore || 0) + (result.redBonus || 0);
        var blue = (result.blueScore || 0) + (result.blueBonus || 0);
        if (isFighter1) {
            return { scored: red, received: blue };
        }
        return { scored: blue, received: red };
    }

    function buildPoolRankingsList(participants, poolMatches) {
        var stats = {};
        var list = participants || [];

        for (var i = 0; i < list.length; i++) {
            var p = list[i];
            stats[p.id] = {
                participantId: p.id,
                name: p.name,
                wins: 0,
                losses: 0,
                pointsScored: 0,
                pointsReceived: 0
            };
        }

        var pools = poolMatches || {};
        for (var poolId in pools) {
            if (!Object.prototype.hasOwnProperty.call(pools, poolId)) continue;
            var matches = pools[poolId] || [];
            for (var m = 0; m < matches.length; m++) {
                var match = matches[m];
                if (match.status !== 'done' || !match.winnerId || !match.result) continue;

                var f1 = stats[match.fighter1Id];
                var f2 = stats[match.fighter2Id];
                if (!f1 || !f2) continue;

                if (match.winnerId === match.fighter1Id) {
                    f1.wins++;
                    f2.losses++;
                } else if (match.winnerId === match.fighter2Id) {
                    f2.wins++;
                    f1.losses++;
                }

                var pts1 = fighterPointsFromResult(match.result, true);
                var pts2 = fighterPointsFromResult(match.result, false);
                f1.pointsScored += pts1.scored;
                f1.pointsReceived += pts1.received;
                f2.pointsScored += pts2.scored;
                f2.pointsReceived += pts2.received;
            }
        }

        var rows = [];
        for (var id in stats) {
            if (Object.prototype.hasOwnProperty.call(stats, id)) {
                rows.push(stats[id]);
            }
        }

        rows.sort(comparePoolRankings);
        return rows;
    }

    function comparePoolRankings(a, b) {
        if (b.wins !== a.wins) return b.wins - a.wins;
        if (a.losses !== b.losses) return a.losses - b.losses;
        var diffA = a.pointsScored - a.pointsReceived;
        var diffB = b.pointsScored - b.pointsReceived;
        if (diffB !== diffA) return diffB - diffA;
        if (b.pointsScored !== a.pointsScored) return b.pointsScored - a.pointsScored;
        return String(a.name).localeCompare(String(b.name), 'ru');
    }

    return {
        buildPoolRankingsList: buildPoolRankingsList,
        comparePoolRankings: comparePoolRankings,
        fighterPointsFromResult: fighterPointsFromResult
    };
}));
