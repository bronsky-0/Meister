#!/usr/bin/env node
'use strict';

var assert = require('assert');
var PoolRankings = require('../pool-rankings-core.js');

var participants = [
    { id: 'a', name: 'Алексей' },
    { id: 'b', name: 'Борис' },
    { id: 'c', name: 'Виктор' },
    { id: 'd', name: 'Григорий' }
];

var poolMatches = {
    pool1: [
        {
            id: 'm1',
            fighter1Id: 'a',
            fighter2Id: 'b',
            status: 'done',
            winnerId: 'a',
            result: { redScore: 10, blueScore: 5, redBonus: 0, blueBonus: 0 }
        },
        {
            id: 'm2',
            fighter1Id: 'c',
            fighter2Id: 'd',
            status: 'done',
            winnerId: 'c',
            result: { redScore: 8, blueScore: 12, redBonus: 0, blueBonus: 0 }
        }
    ],
    pool2: [
        {
            id: 'm3',
            fighter1Id: 'a',
            fighter2Id: 'c',
            status: 'done',
            winnerId: 'c',
            result: { redScore: 6, blueScore: 9, redBonus: 1, blueBonus: 0 }
        },
        {
            id: 'm4',
            fighter1Id: 'b',
            fighter2Id: 'd',
            status: 'done',
            winnerId: 'b',
            result: { redScore: 11, blueScore: 7, redBonus: 0, blueBonus: 0 }
        }
    ]
};

var rankings = PoolRankings.buildPoolRankingsList(participants, poolMatches);

assert.strictEqual(rankings.length, 4, 'four participants in ranking');

assert.strictEqual(rankings[0].name, 'Виктор', 'leader: 2 wins');
assert.strictEqual(rankings[0].wins, 2);
assert.strictEqual(rankings[0].losses, 0);
assert.strictEqual(rankings[0].pointsScored, 8 + 9);
assert.strictEqual(rankings[0].pointsReceived, 12 + 7);

assert.strictEqual(rankings[1].name, 'Алексей', 'second: 1W1L better point diff');
assert.strictEqual(rankings[2].name, 'Борис', 'third: 1W1L worse point diff');
assert.strictEqual(rankings[3].name, 'Григорий', 'last: 0 wins');

console.log('OK — pool rankings tests passed\n');
console.log('Пример таблицы рейтинга:\n');
console.log(
    pad('#', 3) +
    pad('Участник', 14) +
    pad('Побед', 7) +
    pad('Пораж.', 8) +
    pad('Нанес', 8) +
    pad('Получил', 9)
);
console.log('-'.repeat(57));

for (var i = 0; i < rankings.length; i++) {
    var r = rankings[i];
    console.log(
        pad(String(i + 1), 3) +
        pad(r.name, 14) +
        pad(String(r.wins), 7) +
        pad(String(r.losses), 8) +
        pad(String(r.pointsScored), 8) +
        pad(String(r.pointsReceived), 9)
    );
}

function pad(str, width) {
    str = String(str);
    while (str.length < width) str += ' ';
    return str;
}
