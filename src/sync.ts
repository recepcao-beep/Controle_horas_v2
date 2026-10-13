export type Row = { id: string; [key: string]: any };
export type Snapshot = {
  sectors: Row[]; employees: Row[]; requests: Row[];
  overtimeMultiplier: number; valeTransporteValue: number;
};
export type Change = { id: string; fields?: Record<string, any>; deleted?: boolean };
export type Changes = { sectors: Change[]; employees: Change[]; requests: Change[]; config: Record<string, number> };
const tables = ['sectors', 'employees', 'requests'] as const;
const equal = (a: any, b: any) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);

// Send only fields actually edited here, never the device's entire cached database.
export function diff(base: Snapshot, next: Snapshot): Changes {
  const result: Changes = { sectors: [], employees: [], requests: [], config: {} };
  for (const table of tables) {
    const before = new Map(base[table].map(row => [row.id, row]));
    for (const row of next[table]) {
      const old = before.get(row.id);
      const fields: Record<string, any> = {};
      for (const key of new Set([...Object.keys(old || {}), ...Object.keys(row)])) {
        if (key !== 'id' && (!old || !equal(old[key], row[key]))) fields[key] = row[key] ?? null;
      }
      if (!old || Object.keys(fields).length) result[table].push({ id: row.id, fields });
      before.delete(row.id);
    }
    for (const id of before.keys()) result[table].push({ id, deleted: true });
  }
  for (const key of ['overtimeMultiplier', 'valeTransporteValue'] as const) {
    if (base[key] !== next[key]) result.config[key] = next[key];
  }
  return result;
}

export function hasChanges(changes: Changes): boolean {
  return tables.some(table => changes[table].length > 0) || Object.keys(changes.config).length > 0;
}

// Rebase unsent local edits onto a fresh remote snapshot, including edits made during a request.
export function rebase(base: Snapshot, local: Snapshot, remote: Snapshot): Snapshot {
  const changes = diff(base, local);
  const merged = { ...remote, ...changes.config };
  for (const table of tables) {
    const rows = new Map(remote[table].map(row => [row.id, row]));
    for (const change of changes[table]) {
      if (change.deleted) rows.delete(change.id);
      else rows.set(change.id, { ...(rows.get(change.id) || local[table].find(row => row.id === change.id)), ...change.fields, id: change.id });
    }
    merged[table] = Array.from(rows.values());
  }
  return merged;
}
