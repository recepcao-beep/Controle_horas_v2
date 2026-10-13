import { test } from 'node:test';
import assert from 'node:assert/strict';
import { diff, rebase, hasChanges, type Snapshot } from './sync.js';
import { saveChanges, validateChanges } from './server/sync.js';

const empty = (): Snapshot => ({ sectors: [], employees: [], requests: [], overtimeMultiplier: 1.25, valeTransporteValue: 12 });

test('two devices preserve independent launches when refreshing', () => {
  const base = empty();
  const mobile = { ...base, requests: [{ id: 'mobile', status: 'PENDENTE' }] };
  const desktop = { ...base, requests: [{ id: 'desktop', status: 'PENDENTE' }] };
  assert.deepEqual(rebase(base, mobile, desktop).requests.map(r => r.id), ['desktop', 'mobile']);
  assert.deepEqual(diff(base, mobile).requests, [{ id: 'mobile', fields: { status: 'PENDENTE' } }]);
});

test('remote deletions and empty tables propagate without resurrecting cached data', () => {
  const base = { ...empty(), sectors: [{ id: 's', name: 'Old' }], requests: [{ id: 'a' }] };
  assert.deepEqual(rebase(base, base, empty()), empty());
});

test('pending edits survive reload and remote refresh', () => {
  const base = { ...empty(), requests: [{ id: 'a', status: 'PENDENTE', records: [] }] };
  const local = { ...base, requests: [{ ...base.requests[0], records: ['local hours'] }] };
  const disk = JSON.parse(JSON.stringify({ local, base }));
  const remote = { ...base, requests: [{ ...base.requests[0], status: 'APROVADO' }] };
  assert.deepEqual(rebase(disk.base, disk.local, remote).requests, [{ id: 'a', status: 'APROVADO', records: ['local hours'] }]);
  assert.equal(hasChanges(diff(disk.base, disk.local)), true);
});

test('acknowledging an in-flight save does not acknowledge later edits', () => {
  const sent = { ...empty(), requests: [{ id: 'a', records: ['first'] }] };
  const latest = { ...sent, requests: [{ id: 'a', records: ['second'] }] };
  assert.deepEqual(diff(sent, latest).requests, [{ id: 'a', fields: { records: ['second'] } }]);
});

test('deletions and zero configuration values are explicit changes', () => {
  const base = { ...empty(), employees: [{ id: 'e' }] };
  const next = { ...empty(), valeTransporteValue: 0 };
  assert.deepEqual(diff(base, next).employees, [{ id: 'e', deleted: true }]);
  assert.deepEqual(diff(base, next).config, { valeTransporteValue: 0 });
});

// Small Sheets API fake: concurrent callers read snapshots, cell writes and append
// then operate on the current shared table. No Google credentials/production data.
function fakeSheets() {
  const tables: Record<string, any[][]> = {
    Solicitacoes: [['id', 'employeeName', 'employeeType', 'sectorId', 'sectorName', 'weekStarting', 'records', 'status', 'calculatedValue', 'totalTimeDecimal', 'createdAt', 'editJustification', 'syncDeleted']],
  };
  let failAfterAppend = false;
  const sheetName = (range: string) => range.split('!')[0].replace(/^'|'$/g, '');
  const sheets = { spreadsheets: {
    get: async () => ({ data: { sheets: Object.keys(tables).map(title => ({ properties: { title } })) } }),
    batchUpdate: async ({ requestBody }: any) => {
      for (const req of requestBody.requests) {
        const title = req.addSheet.properties.title;
        if (tables[title]) throw new Error('already exists');
        tables[title] = [];
      }
    },
    values: {
      get: async ({ range }: any) => ({ data: { values: structuredClone(tables[sheetName(range)] || []) } }),
      batchUpdate: async ({ requestBody }: any) => {
        for (const update of requestBody.data) {
          const matrix = tables[sheetName(update.range)];
          const [, col, row] = update.range.match(/!([A-Z])(\d+)$/);
          const r = Number(row) - 1;
          matrix[r] ||= [];
          update.values[0].forEach((value: any, i: number) => matrix[r][col.charCodeAt(0) - 65 + i] = value);
        }
      },
      append: async ({ range, requestBody }: any) => {
        tables[sheetName(range)].push(...structuredClone(requestBody.values));
        if (failAfterAppend) { failAfterAppend = false; throw new Error('response lost'); }
      },
    },
  } };
  return { sheets, tables, loseNextResponse: () => { failAfterAppend = true; } };
}

test('concurrent mobile and desktop saves append both IDs instead of clearing the table', async () => {
  const { sheets, tables } = fakeSheets();
  await Promise.all(['mobile', 'desktop'].map(id => saveChanges(sheets, 'test', diff(empty(), { ...empty(), requests: [{ id, employeeName: id, status: 'PENDENTE' }] }))));
  assert.deepEqual(tables.Solicitacoes.slice(1).map(row => row[0]).sort(), ['desktop', 'mobile']);
});

test('lost response is retryable without duplicating a launch', async () => {
  const { sheets, tables, loseNextResponse } = fakeSheets();
  const changes = diff(empty(), { ...empty(), requests: [{ id: 'mobile', status: 'PENDENTE' }] });
  loseNextResponse();
  await assert.rejects(saveChanges(sheets, 'test', changes), /response lost/);
  await saveChanges(sheets, 'test', changes);
  assert.equal(tables.Solicitacoes.length, 2);
});

test('concurrent edits to different fields of the same record survive', async () => {
  const { sheets, tables } = fakeSheets();
  const base = { ...empty(), requests: [{ id: 'a', employeeName: 'Ana', status: 'PENDENTE' }] };
  await saveChanges(sheets, 'test', diff(empty(), base));
  await Promise.all([
    saveChanges(sheets, 'test', diff(base, { ...base, requests: [{ ...base.requests[0], employeeName: 'Ana Maria' }] })),
    saveChanges(sheets, 'test', diff(base, { ...base, requests: [{ ...base.requests[0], status: 'APROVADO' }] })),
  ]);
  const row = Object.fromEntries(tables.Solicitacoes[0].map((key, i) => [key, tables.Solicitacoes[1][i]]));
  assert.equal(row.employeeName, 'Ana Maria');
  assert.equal(row.status, 'APROVADO');
});

test('deleting one launch keeps the other device launch and records a tombstone', async () => {
  const { sheets, tables } = fakeSheets();
  const base = { ...empty(), requests: [{ id: 'a', status: 'PENDENTE' }, { id: 'b', status: 'PENDENTE' }] };
  await saveChanges(sheets, 'test', diff(empty(), base));
  await saveChanges(sheets, 'test', diff(base, { ...base, requests: [base.requests[1]] }));
  const headers = tables.Solicitacoes[0];
  assert.equal(tables.Solicitacoes[1][headers.indexOf('syncDeleted')], true);
  assert.equal(tables.Solicitacoes[1][headers.indexOf('status')], 'DELETADO');
  assert.equal(tables.Solicitacoes[2][headers.indexOf('status')], 'PENDENTE');
});

test('old whole-database payloads are rejected', () => {
  assert.equal(Boolean(validateChanges(empty())), false);
  assert.equal(Boolean(validateChanges(diff(empty(), empty()))), true);
  assert.equal(Boolean(validateChanges({ ...diff(empty(), empty()), config: { valeTransporteValue: 'bad' } })), false);
});
