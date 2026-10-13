import type { Changes } from '../sync.js';

const names = { sectors: 'Setores', employees: 'Funcionarios', requests: 'Solicitacoes', config: 'Config' };
const keyOf = (key: string) => key.normalize('NFD').replace(/[\u0300-\u036f\s]/g, '').toLowerCase();
const quote = (name: string) => `'${name.replace(/'/g, "''")}'`;

export function validateChanges(value: any): value is Changes {
  return value && ['sectors', 'employees', 'requests'].every(table =>
    Array.isArray(value[table]) && value[table].length <= 500 && value[table].every((c: any) =>
      c && typeof c.id === 'string' && c.id.length > 0 && c.id.length < 200 &&
      (c.deleted === true || (c.fields && typeof c.fields === 'object' && !Array.isArray(c.fields)))
    )) && value.config && Object.entries(value.config).every(([key, val]) =>
      ['overtimeMultiplier', 'valeTransporteValue'].includes(key) && typeof val === 'number' && Number.isFinite(val));
}

// Existing records are updated in place; new IDs are appended by Sheets itself.
// No read/merge/replace of the whole table, and no process-local lock (Vercel has multiple instances).
export async function saveChanges(sheets: any, spreadsheetId: string, changes: Changes) {
  const info = await sheets.spreadsheets.get({ spreadsheetId });
  const titles: string[] = (info.data.sheets || []).map((s: any) => s.properties.title);
  for (const table of Object.keys(names) as (keyof typeof names)[]) {
    const edits = table === 'config'
      ? Object.entries(changes.config).map(([id, value]) => ({ id, fields: { value }, deleted: false }))
      : changes[table];
    if (!edits.length) continue;
    const title = titles.find(title => keyOf(title) === keyOf(names[table])) || names[table];
    if (!titles.includes(title)) {
      try {
        await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [{ addSheet: { properties: { title } } }] } });
      } catch (error) {
        // Another instance may have created the same sheet concurrently. Verify before continuing.
        const refreshed = await sheets.spreadsheets.get({ spreadsheetId });
        if (!refreshed.data.sheets?.some((s: any) => s.properties.title === title)) throw error;
      }
    }
    const response = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${quote(title)}!A:Z`, valueRenderOption: 'UNFORMATTED_VALUE' });
    const rows: any[][] = response.data.values || [];
    const defaults: Record<string, string[]> = {
      sectors: ['id', 'name', 'fixedRate', 'syncDeleted'],
      employees: ['id', 'name', 'sectorId', 'type', 'salary', 'monthlyHours', 'fixedDayOff', 'workRegime', 'parityRef', 'syncDeleted'],
      requests: ['id', 'employeeId', 'employeeName', 'employeeType', 'sectorId', 'sectorName', 'weekStarting', 'records', 'status', 'calculatedValue', 'totalTimeDecimal', 'createdAt', 'editJustification', 'syncDeleted'],
      config: ['key', 'value'],
    };
    const headers = rows[0]?.length ? rows[0].map(String) : defaults[table];
    // Add the same fixed schema columns in the same order on every instance.
    for (const field of defaults[table]) if (!headers.some(h => keyOf(h) === keyOf(field))) headers.push(field);
    if (headers.length > 26) throw new Error(`A aba ${title} excede as 26 colunas suportadas.`);
    const updates: any[] = [{ range: `${quote(title)}!A1`, values: [headers] }];
    const idColumn = headers.findIndex(h => keyOf(h) === (table === 'config' ? 'key' : 'id'));
    const append: any[][] = [];
    const encode = (value: any) => value == null ? '' : typeof value === 'object' ? JSON.stringify(value) : value;
    const columnName = (index: number) => String.fromCharCode(65 + index);
    for (const edit of edits) {
      const indexes = rows.flatMap((row, i) => i > 0 && String(row[idColumn]) === edit.id ? [i] : []);
      const fields: Record<string, any> = edit.deleted ? { syncDeleted: true } : { ...edit.fields, ...(table === 'config' ? {} : { syncDeleted: false }) };
      if (table === 'requests' && edit.deleted) fields.status = 'DELETADO';
      if (indexes.length) {
        // Update only changed cells: concurrent changes to another record/field survive.
        for (const index of indexes) for (const [field, value] of Object.entries(fields)) {
          const column = headers.findIndex(h => keyOf(h) === keyOf(field));
          if (column >= 0) updates.push({ range: `${quote(title)}!${columnName(column)}${index + 1}`, values: [[encode(value)]] });
        }
      } else if (!edit.deleted) {
        const row = headers.map(h => keyOf(h) === (table === 'config' ? 'key' : 'id') ? edit.id : encode(Object.entries(fields).find(([key]) => keyOf(key) === keyOf(h))?.[1]));
        append.push(row);
      }
    }
    await sheets.spreadsheets.values.batchUpdate({ spreadsheetId, requestBody: { valueInputOption: 'RAW', data: updates } });
    if (append.length) await sheets.spreadsheets.values.append({ spreadsheetId, range: `${quote(title)}!A:Z`, valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS', requestBody: { values: append } });
  }
}
