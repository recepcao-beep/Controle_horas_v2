import { google } from 'googleapis';

export async function distributeData(sheets: any, spreadsheetId: string, requests: any[]) {
  // 1. Fetch HE - REGISTRADO and HE - FIXO
  const response = await sheets.spreadsheets.values.batchGet({
    spreadsheetId,
    ranges: ["'HE - REGISTRADO'!A:Z", "'HE - FIXO'!A:Z"],
    valueRenderOption: 'FORMULA',
  });

  const valueRanges = response.data.valueRanges || [];
  let regMatriz = valueRanges.find((v: any) => {
    const rangeName = (v.range || '').toUpperCase().replace(/\s+/g, '');
    return rangeName.includes('HE-REGISTRADO');
  })?.values || [];
  
  let fixoMatriz = valueRanges.find((v: any) => {
    const rangeName = (v.range || '').toUpperCase().replace(/\s+/g, '');
    return rangeName.includes('HE-FIXO');
  })?.values || [];

  if (regMatriz.length === 0 && fixoMatriz.length === 0) {
    throw new Error("As abas 'HE - REGISTRADO' e 'HE - FIXO' não foram encontradas ou estão vazias na planilha.");
  }

  // Helper functions
  const parseRecords = (recs: any) => {
    if (typeof recs === 'string') { try { return JSON.parse(recs); } catch(e) { return []; } }
    return Array.isArray(recs) ? recs : [];
  };

  const agruparSolicitacoesPorFuncionario = (reqs: any[]) => {
    const agrupado: any = {};
    reqs.forEach(req => {
      const key = (req.employeeName || "").trim().toUpperCase() + "|" + 
                  (req.employeeType || "").trim().toUpperCase() + "|" + 
                  (req.sectorName || "").trim().toUpperCase();
      if (!agrupado[key]) {
        agrupado[key] = {
          employeeName: req.employeeName,
          employeeType: req.employeeType,
          sectorName: req.sectorName,
          records: []
        };
      }
      let recs = parseRecords(req.records);
      agrupado[key].records = agrupado[key].records.concat(recs);
    });

    const resultado: any[] = [];
    Object.keys(agrupado).forEach(key => {
      let grupo = agrupado[key];
      let records = grupo.records;
      records = records.filter((d: any) => d.realEntry || d.realExit || d.punchEntry || d.punchExit);
      records.sort((a: any, b: any) => (a.date > b.date) ? 1 : -1);

      const semanas: any = {};
      records.forEach((rec: any) => {
        let partes = rec.date.split("-");
        let d = new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]), 12, 0, 0);
        let diaSemana = d.getDay();
        let diffParaSegunda = diaSemana === 0 ? -6 : 1 - diaSemana;
        let segunda = new Date(d);
        segunda.setDate(d.getDate() + diffParaSegunda);
        let keySemana = segunda.getFullYear() + "-" + (segunda.getMonth() + 1) + "-" + segunda.getDate();
        
        if (!semanas[keySemana]) semanas[keySemana] = [];
        let idx = semanas[keySemana].findIndex((r: any) => r.date === rec.date);
        if (idx !== -1) semanas[keySemana][idx] = rec;
        else semanas[keySemana].push(rec);
      });
      
      Object.keys(semanas).forEach(keySemana => {
        let recsSemana = semanas[keySemana];
        if ((grupo.employeeType || "").toUpperCase().trim() === "REGISTRADO") {
          resultado.push({ ...grupo, records: recsSemana });
        } else {
          for (let i = 0; i < recsSemana.length; i += 7) {
            resultado.push({ ...grupo, records: recsSemana.slice(i, i + 7) });
          }
        }
      });
    });
    return resultado;
  };

  const limparMatriz = (matriz: any[], tipo: string) => {
    for (let i = 0; i < matriz.length; i++) {
      if (!matriz[i]) matriz[i] = [];
      if (i === 13 || i === 14) continue; 
      let txtA = (matriz[i][0] || "").toString().toUpperCase();
      if (txtA.includes("NOME COMPLETO:")) {
        matriz[i][1] = ""; 
        let start = (tipo === "REGISTRADO") ? 7 : 3;
        let limit = (tipo === "REGISTRADO") ? 8 : 7;
        for (let g = 0; g < limit; g++) {
          let rIdx = i + start + g;
          if (!matriz[rIdx]) matriz[rIdx] = [];
          if (tipo === "REGISTRADO") [0, 2, 3, 6, 7].forEach(c => matriz[rIdx][c] = ""); 
          else [0, 1, 2, 4].forEach(c => matriz[rIdx][c] = ""); 
        }
      }
      if (tipo === "FIXO" && matriz[i][7] && matriz[i][7].toString().toUpperCase().includes("NOME COMPLETO:")) {
        matriz[i][8] = ""; 
        for (let g = 0; g < 7; g++) {
          let rIdx = i + 3 + g;
          if (!matriz[rIdx]) matriz[rIdx] = [];
          [7, 8, 9, 11].forEach(c => matriz[rIdx][c] = "");
        }
      }
    }
  };

  const preencherColunaAERegistros = (matriz: any[], linhaInicio: number, records: any[]) => {
    let horas = parseRecords(records);
    if (horas.length === 0) return;
    horas.sort((a: any, b: any) => (a.date > b.date) ? 1 : -1);
    let partesData = horas[0].date.split("-"); 
    let dataRefOriginal = new Date(parseInt(partesData[0]), parseInt(partesData[1]) - 1, parseInt(partesData[2]), 12, 0, 0);
    let diaSemana = dataRefOriginal.getDay();
    let diffParaSegunda = diaSemana === 0 ? -6 : 1 - diaSemana;
    let dataRef = new Date(dataRefOriginal);
    dataRef.setDate(dataRefOriginal.getDate() + diffParaSegunda);
    
    const diasDaSemana = ["SEGUNDA-FEIRA", "TERÇA-FEIRA", "QUARTA-FEIRA", "QUINTA-FEIRA", "SEXTA-FEIRA", "SÁBADO", "DOMINGO"];
    
    for (let i = 0; i < 7; i++) {
      let r = linhaInicio + i;
      if (!matriz[r]) matriz[r] = [];
      let dataLoop = new Date(dataRef);
      dataLoop.setDate(dataRef.getDate() + i);
      
      matriz[r][0] = diasDaSemana[i];
      
      let sBusca = dataLoop.toISOString().split('T')[0];
      let reg = horas.find((h: any) => h.date === sBusca);
      if (reg) {
        matriz[r][2] = reg.realEntry || ""; matriz[r][3] = reg.punchEntry || "";
        matriz[r][6] = reg.punchExit || ""; matriz[r][7] = reg.realExit || "";
      }
    }
  };

  const preencherHorasNaMatriz = (matriz: any[], linhaInicio: number, records: any[], col: number) => {
    let horas = parseRecords(records);
    horas.sort((a: any, b: any) => (a.date > b.date) ? 1 : -1);
    let preenchidas = 0;
    let diasComDados = horas.filter((d: any) => d.realEntry || d.realExit);
    for (let i = 0; i < diasComDados.length; i++) {
      if (preenchidas < 7) {
        let r = linhaInicio + preenchidas;
        if (!matriz[r]) matriz[r] = [];
        let partes = diasComDados[i].date.split("-");
        let dataFormatada = partes[2] + "/" + partes[1] + "/" + partes[0];
        
        matriz[r][col] = dataFormatada;
        matriz[r][col + 1] = diasComDados[i].realEntry || "";
        matriz[r][col + 2] = diasComDados[i].realExit || "";  
        preenchidas++;
      }
    }
  };

  const localizarVagaNoBlocoSetor = (matriz: any[], linhaSetor: number, col: number) => {
    let contador = 0;
    for (let i = linhaSetor; i < matriz.length; i++) {
      if (!matriz[i]) matriz[i] = [];
      let txt = (matriz[i][col] || "").toString().toUpperCase();
      if (txt.includes("NOME COMPLETO:")) {
        if ((matriz[i][col + 1] || "").toString().trim() === "") return i;
        contador++;
        if (contador >= 5) break; 
      }
      if (i > linhaSetor && (matriz[i][1] || "").toString().toUpperCase().includes("SETOR:")) break;
    }
    return -1;
  };

  // Process data
  const aprovados = requests.filter(r => (r.status || "").toUpperCase().trim() === "APROVADO");
  const agrupados = agruparSolicitacoesPorFuncionario(aprovados);

  // Process REGISTRADO
  if (regMatriz.length > 0) {
    limparMatriz(regMatriz, "REGISTRADO");
    const regEmployees = agrupados.filter(r => r.employeeType.toUpperCase().trim() === "REGISTRADO");
    const sectorsMap: any = {};
    regEmployees.forEach(emp => {
      const sName = (emp.sectorName || "GERAL").toUpperCase().trim();
      if (!sectorsMap[sName]) sectorsMap[sName] = [];
      sectorsMap[sName].push(emp);
    });

    let currentSheetIdx = 0; 
    Object.keys(sectorsMap).sort().forEach(sName => {
      const emps = sectorsMap[sName];
      for (let i = 0; i < emps.length; i += 2) {
        while (currentSheetIdx < regMatriz.length && regMatriz[currentSheetIdx + 4] && (regMatriz[currentSheetIdx + 4][1] || "").toString().trim() !== "") {
           currentSheetIdx += 52;
        }
        if (currentSheetIdx >= regMatriz.length) break;

        const emp1 = emps[i];
        if (!regMatriz[currentSheetIdx + 4]) regMatriz[currentSheetIdx + 4] = [];
        if (!regMatriz[currentSheetIdx + 1]) regMatriz[currentSheetIdx + 1] = [];
        regMatriz[currentSheetIdx + 4][1] = emp1.employeeName;
        regMatriz[currentSheetIdx + 1][1] = emp1.sectorName; 
        preencherColunaAERegistros(regMatriz, currentSheetIdx + 4 + 7, emp1.records);

        if (i + 1 < emps.length) {
          const emp2 = emps[i + 1];
          if (!regMatriz[currentSheetIdx + 28]) regMatriz[currentSheetIdx + 28] = [];
          regMatriz[currentSheetIdx + 28][1] = emp2.employeeName;
          preencherColunaAERegistros(regMatriz, currentSheetIdx + 28 + 7, emp2.records);
        }
        currentSheetIdx += 52;
      }
    });
  }

  // Process FIXO
  if (fixoMatriz.length > 0) {
    limparMatriz(fixoMatriz, "FIXO");
    agrupados.filter(r => r.employeeType.toUpperCase().trim() === "FIXO").forEach(func => {
      let fIdx = -1; let colBase = -1; 
      let nomeSetorAlvo = (func.sectorName || "").toUpperCase().trim();
      for (let i = 0; i < fixoMatriz.length; i++) {
        if (!fixoMatriz[i]) fixoMatriz[i] = [];
        if ((fixoMatriz[i][1] || "").toString().toUpperCase().trim() === nomeSetorAlvo) {
          let buscaEsq = localizarVagaNoBlocoSetor(fixoMatriz, i, 0);
          if (buscaEsq !== -1) { fIdx = buscaEsq; colBase = 0; break; }
          let buscaDir = localizarVagaNoBlocoSetor(fixoMatriz, i, 7);
          if (buscaDir !== -1) { fIdx = buscaDir; colBase = 7; break; }
        }
      }
      if (fIdx !== -1) {
        if (!fixoMatriz[fIdx]) fixoMatriz[fIdx] = [];
        fixoMatriz[fIdx][colBase + 1] = func.employeeName;
        preencherHorasNaMatriz(fixoMatriz, fIdx + 3, func.records, colBase);
      }
    });
  }

  // Update sheets
  const dataToUpdate = [];
  if (regMatriz.length > 0) {
    dataToUpdate.push({
      range: "'HE - REGISTRADO'!A1",
      values: regMatriz
    });
  }
  if (fixoMatriz.length > 0) {
    dataToUpdate.push({
      range: "'HE - FIXO'!A1",
      values: fixoMatriz
    });
  }

  if (dataToUpdate.length > 0) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: dataToUpdate,
      },
    });
  }
}

export async function performClosing(sheets: any, spreadsheetId: string) {
  // 1. Get spreadsheet info to find sheet IDs
  const ssInfo = await sheets.spreadsheets.get({ spreadsheetId });
  const sheetsList = ssInfo.data.sheets || [];
  
  const regSheet = sheetsList.find((s: any) => s.properties.title === 'HE - REGISTRADO');
  const fixoSheet = sheetsList.find((s: any) => s.properties.title === 'HE - FIXO');
  const solSheet = sheetsList.find((s: any) => s.properties.title === 'Solicitacoes');
  let backupSheet = sheetsList.find((s: any) => s.properties.title === 'BACKUP');

  if (!solSheet) throw new Error("Aba 'Solicitacoes' não encontrada.");

  const dateStr = new Date().toLocaleDateString('pt-BR').replace(/\//g, '-');
  const requestsToBatch: any[] = [];

  // Duplicate REGISTRADO and FIXO
  if (regSheet) {
    requestsToBatch.push({
      duplicateSheet: {
        sourceSheetId: regSheet.properties.sheetId,
        insertSheetIndex: sheetsList.length,
        newSheetName: `HE - REGISTRADO (${dateStr})`
      }
    });
  }
  if (fixoSheet) {
    requestsToBatch.push({
      duplicateSheet: {
        sourceSheetId: fixoSheet.properties.sheetId,
        insertSheetIndex: sheetsList.length + 1,
        newSheetName: `HE - FIXO (${dateStr})`
      }
    });
  }

  // Create BACKUP sheet if it doesn't exist
  if (!backupSheet) {
    requestsToBatch.push({
      addSheet: {
        properties: {
          title: 'BACKUP'
        }
      }
    });
  }

  if (requestsToBatch.length > 0) {
    try {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: requestsToBatch }
      });
    } catch (e: any) {
      // Ignore errors if sheet already exists
      console.warn("Error duplicating sheets (might already exist):", e.message);
    }
  }

  // Backup Solicitacoes data
  const solData = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: 'Solicitacoes!A:Z'
  });

  const values = solData.data.values || [];
  if (values.length > 1) {
    // Append to BACKUP
    const headers = values[0];
    const dataRows = values.slice(1);
    
    // If BACKUP was just created, we might need to add headers
    if (!backupSheet) {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: 'BACKUP!A1',
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [headers] }
      });
    }

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: 'BACKUP!A:A',
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: dataRows }
    });

    // Clear Solicitacoes (keep headers)
    await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: 'Solicitacoes!A2:Z'
    });
  }

  // Clear REGISTRADO and FIXO using the distributeData logic with empty requests
  await distributeData(sheets, spreadsheetId, []);
}

// Gera PDFs do modelo oficial do RH no Drive SEM limpar o banco de dados.
// Estrutura criada automaticamente:
// FECHAMENTOS / HE-FIXO|HE-REGISTRADO / ANO / MÊS / DD-MM / SETOR-DD-MM.pdf
export async function generateRhClosings(
  sourceSheets: any,
  targetSheets: any,
  drive: any,
  sourceSpreadsheetId: string,
  templateSpreadsheetId: string,
  destinationFolderId: string
) {
  const { Readable } = await import('stream');

  const solData = await sourceSheets.spreadsheets.values.get({
    spreadsheetId: sourceSpreadsheetId,
    range: 'Solicitacoes!A:Z'
  });

  const values = solData.data.values || [];
  if (values.length <= 1) throw new Error('Não há solicitações para gerar o fechamento.');

  const headers = values[0].map((h: any) => String(h || '').trim());
  const rows = values.slice(1).map((row: any[]) => {
    const obj: any = {};
    headers.forEach((h: string, i: number) => obj[h] = row[i]);
    // Alguns backups antigos salvaram `records` como JSON mais de uma vez.
    // Desembrulha até virar array, sem descartar silenciosamente os registros.
    try {
      let recs: any = obj.records;
      for (let i = 0; i < 3 && typeof recs === 'string'; i++) recs = JSON.parse(recs);
      obj.records = Array.isArray(recs) ? recs : [];
    } catch { obj.records = []; }
    return obj;
  }).filter((r: any) => String(r.status || '').trim().toUpperCase() === 'APROVADO');

  if (!rows.length) throw new Error('Não há solicitações APROVADAS para gerar o fechamento.');

  const clean = (v: any) => String(v || '').trim();

  // O campo weekStarting do app legado nem sempre guarda a segunda-feira.
  // Em alguns registros ele contém o DOMINGO de fechamento (ex.: 13/09),
  // enquanto os lançamentos pertencem à semana 07/09 a 13/09. Se usarmos
  // esse valor diretamente, o gerador procura as horas em 13/09..19/09 e
  // deixa 07/09..12/09 em branco. Normalizamos sempre para a segunda-feira
  // da semana e, quando houver registros, usamos a primeira data real como
  // fonte de verdade.
  const mondayOf = (iso: string) => {
    const value = clean(iso).slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return 'SEM-DATA';
    const [y, m, d] = value.split('-').map(Number);
    const dt = new Date(Date.UTC(y, m - 1, d));
    const day = dt.getUTCDay(); // 0=domingo, 1=segunda...
    const diff = day === 0 ? -6 : 1 - day;
    dt.setUTCDate(dt.getUTCDate() + diff);
    return dt.toISOString().slice(0, 10);
  };

  const normalizeDate = (value: any) => {
    const raw = clean(value);
    if (!raw) return '';
    const iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
    const br = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
    if (br) return `${br[3]}-${br[2]}-${br[1]}`;
    return '';
  };

  const hasHours = (r: any) => !!(clean(r?.realEntry) || clean(r?.punchEntry) || clean(r?.punchExit) || clean(r?.realExit));

  // Mescla registros repetidos do mesmo colaborador/data SEM deixar um registro
  // posterior vazio apagar um horário que já existia no banco.
  const mergeRecord = (oldRec: any, newRec: any, date: string) => ({
    ...(oldRec || {}),
    ...(newRec || {}),
    date,
    realEntry: clean(newRec?.realEntry) || clean(oldRec?.realEntry),
    punchEntry: clean(newRec?.punchEntry) || clean(oldRec?.punchEntry),
    punchExit: clean(newRec?.punchExit) || clean(oldRec?.punchExit),
    realExit: clean(newRec?.realExit) || clean(oldRec?.realExit),
  });

  // Consolida por SETOR + SEMANA usando a DATA DE CADA LANÇAMENTO como fonte
  // de verdade. Solicitação sem horário não gera ficha em branco.
  const groups = new Map<string, any>();
  const ensureEmployee = (sector: string, week: string, r: any) => {
    const key = `${sector.toUpperCase()}|${week}`;
    if (!groups.has(key)) groups.set(key, { sector, week, employees: new Map<string, any>() });
    const g = groups.get(key);
    const empKey = `${clean(r.employeeName).toUpperCase()}|${clean(r.employeeType).toUpperCase()}`;
    if (!g.employees.has(empKey)) g.employees.set(empKey, {
      employeeName: clean(r.employeeName), employeeType: clean(r.employeeType).toUpperCase(), records: []
    });
    return g.employees.get(empKey);
  };

  for (const r of rows) {
    const sector = clean(r.sectorName) || 'GERAL';
    // Só entram no fechamento dias que realmente possuem algum horário.
    const validRecords = (r.records || []).filter((rec: any) => normalizeDate(rec?.date) && hasHours(rec));
    for (const rec of validRecords) {
      const date = normalizeDate(rec.date);
      const week = mondayOf(date);
      const emp = ensureEmployee(sector, week, r);
      const idx = emp.records.findIndex((x: any) => normalizeDate(x.date) === date);
      if (idx >= 0) emp.records[idx] = mergeRecord(emp.records[idx], rec, date);
      else emp.records.push(mergeRecord(null, rec, date));
    }
  }

  if (!groups.size) throw new Error('Há solicitações aprovadas, mas nenhuma possui horários válidos para gerar o fechamento.');

  const addDays = (iso: string, days: number) => {
    const [y,m,d] = iso.split('-').map(Number);
    const dt = new Date(Date.UTC(y, m - 1, d + days));
    return dt.toISOString().slice(0,10);
  };
  const safeName = (s: string) => s.replace(/[\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim();
  const monthNames = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];
  const dateFolder = (iso: string) => {
    if (!iso || iso === 'SEM-DATA') return 'SEM-DATA';
    const [,m,d] = iso.split('-');
    return `${d}-${m}`;
  };

  const driveEscape = (v: string) => v.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
  const ensureFolder = async (parentId: string, name: string) => {
    const found = await drive.files.list({
      q: `'${driveEscape(parentId)}' in parents and mimeType='application/vnd.google-apps.folder' and name='${driveEscape(name)}' and trashed=false`,
      fields: 'files(id,name)',
      pageSize: 10
    });
    const existing = found.data.files?.[0];
    if (existing?.id) return existing.id;
    const created = await drive.files.create({
      requestBody: { name, mimeType: 'application/vnd.google-apps.folder', parents: [parentId] },
      fields: 'id'
    });
    if (!created.data.id) throw new Error(`Não foi possível criar a pasta ${name}.`);
    return created.data.id;
  };

  const ensureClosingPath = async (type: 'HE-FIXO'|'HE-REGISTRADO', week: string) => {
    const typeId = await ensureFolder(destinationFolderId, type);
    if (week === 'SEM-DATA') {
      const undatedId = await ensureFolder(typeId, 'SEM-DATA');
      return { typeId, finalId: undatedId };
    }
    const [year, month] = week.split('-').map(Number);
    const yearId = await ensureFolder(typeId, String(year));
    const monthId = await ensureFolder(yearId, monthNames[month - 1]);
    const finalId = await ensureFolder(monthId, dateFolder(week));
    return { typeId, finalId };
  };

  const removeExistingPdf = async (folderId: string, name: string) => {
    const found = await drive.files.list({
      q: `'${driveEscape(folderId)}' in parents and name='${driveEscape(name)}' and mimeType='application/pdf' and trashed=false`,
      fields: 'files(id)',
      pageSize: 100
    });
    for (const f of found.data.files || []) {
      if (f.id) await drive.files.update({ fileId: f.id, requestBody: { trashed: true } });
    }
  };

  const exportTypePdf = async (
    populatedSpreadsheetId: string,
    type: 'HE-FIXO'|'HE-REGISTRADO',
    sector: string,
    week: string
  ) => {
    const { finalId } = await ensureClosingPath(type, week);
    const baseName = safeName(`${sector.toUpperCase()}-${dateFolder(week)}`);
    const pdfName = `${baseName}.pdf`;

    // Cópia temporária: deixa visíveis apenas as folhas do tipo escolhido.
    const temp = await drive.files.copy({
      fileId: populatedSpreadsheetId,
      requestBody: { name: `TEMP-${type}-${baseName}`, parents: [destinationFolderId] },
      fields: 'id'
    });
    const tempId = temp.data.id;
    if (!tempId) throw new Error(`Não foi possível preparar o PDF de ${type} / ${sector}.`);

    try {
      const info = await targetSheets.spreadsheets.get({ spreadsheetId: tempId });
      const prefix = type === 'HE-FIXO' ? 'HE - FIXO' : 'HE - REGISTRADO';
      const deleteRequests = (info.data.sheets || [])
        .filter((s: any) => !String(s.properties?.title || '').startsWith(prefix))
        .map((s: any) => ({ deleteSheet: { sheetId: s.properties.sheetId } }));
      if (deleteRequests.length) {
        await targetSheets.spreadsheets.batchUpdate({ spreadsheetId: tempId, requestBody: { requests: deleteRequests } });
      }

      const exported = await drive.files.export(
        { fileId: tempId, mimeType: 'application/pdf' },
        { responseType: 'arraybuffer' }
      );
      const pdfBuffer = Buffer.from(exported.data as ArrayBuffer);

      await removeExistingPdf(finalId, pdfName);
      const uploaded = await drive.files.create({
        requestBody: { name: pdfName, mimeType: 'application/pdf', parents: [finalId] },
        media: { mimeType: 'application/pdf', body: Readable.from(pdfBuffer) },
        fields: 'id,name,webViewLink'
      });
      return {
        id: uploaded.data.id,
        name: uploaded.data.name || pdfName,
        url: uploaded.data.webViewLink,
        sector,
        week,
        type,
        folderId: finalId
      };
    } finally {
      await drive.files.update({ fileId: tempId, requestBody: { trashed: true } }).catch(() => {});
    }
  };

  // Garante as duas pastas principais mesmo que um dos tipos não tenha lançamentos na semana.
  await ensureFolder(destinationFolderId, 'HE-FIXO');
  await ensureFolder(destinationFolderId, 'HE-REGISTRADO');

  const results: any[] = [];

  // Descobre as posições diretamente no MODELO RH. Assim uma mudança de uma
  // linha no modelo não desloca nomes/horários para células erradas.
  const templateLayoutCache = new Map<string, any>();
  const getSheetLayout = async (spreadsheetId: string, sheetName: string, type: 'REGISTRADO'|'FIXO') => {
    const cacheKey = `${spreadsheetId}|${sheetName}|${type}`;
    if (templateLayoutCache.has(cacheKey)) return templateLayoutCache.get(cacheKey);
    const vr = await targetSheets.spreadsheets.values.get({ spreadsheetId, range: `'${sheetName.replace(/'/g, "''")}'!A:Z`, valueRenderOption: 'FORMATTED_VALUE' });
    const matrix: any[][] = vr.data.values || [];
    const norm = (v: any) => clean(v).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    const slots: any[] = [];
    if (type === 'REGISTRADO') {
      for (let r = 0; r < matrix.length; r++) {
        for (let c = 0; c < (matrix[r] || []).length; c++) {
          if (!norm(matrix[r]?.[c]).includes('NOME COMPLETO')) continue;
          let mondayRow = -1;
          for (let rr = r + 1; rr < Math.min(matrix.length, r + 20); rr++) {
            if (norm(matrix[rr]?.[0]).includes('SEGUNDA-FEIRA')) { mondayRow = rr + 1; break; }
          }
          if (mondayRow > 0) slots.push({ nameRow: r + 1, nameCol: c + 2, firstDataRow: mondayRow });
        }
      }
    }
    const layout = { matrix, slots };
    templateLayoutCache.set(cacheKey, layout);
    return layout;
  };

  const colLetter = (n: number) => { let x=n, out=''; while(x>0){ const m=(x-1)%26; out=String.fromCharCode(65+m)+out; x=Math.floor((x-1)/26); } return out; };
  const brDate = (iso: string) => { const [y,m,d]=iso.split('-'); return `${d}/${m}/${y}`; };

  const putHeaderByLabels = (matrix: any[][], sheet: string, put: any, sector: string, week: string) => {
    const norm = (v: any) => clean(v).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    for (let r=0; r<Math.min(matrix.length, 12); r++) {
      for (let c=0; c<(matrix[r]||[]).length; c++) {
        const t=norm(matrix[r]?.[c]);
        if (t.startsWith('SETOR')) put(sheet, `${colLetter(c+2)}${r+1}`, sector);
        if (week !== 'SEM-DATA' && t === 'DATA:') put(sheet, `${colLetter(c+2)}${r+1}`, brDate(week));
        if (week !== 'SEM-DATA' && (t === 'ATE' || t === 'ATE:')) put(sheet, `${colLetter(c+2)}${r+1}`, brDate(addDays(week,6)));
      }
    }
  };

  for (const g of groups.values()) {
    const fileName = safeName(`TEMP-FECHAMENTO-${g.sector}-${dateFolder(g.week)}`);
    const copied = await drive.files.copy({
      fileId: templateSpreadsheetId,
      requestBody: { name: fileName, parents: [destinationFolderId] },
      fields: 'id,name'
    });
    const targetId = copied.data.id;
    if (!targetId) throw new Error(`Falha ao copiar o modelo para o setor ${g.sector}.`);

    try {
      const info = await targetSheets.spreadsheets.get({ spreadsheetId: targetId });
      const templateReg = info.data.sheets?.find((s: any) => s.properties?.title === 'HE - REGISTRADO');
      const templateFixo = info.data.sheets?.find((s: any) => s.properties?.title === 'HE - FIXO');
      if (!templateReg || !templateFixo) throw new Error("O MODELO RH precisa conter as abas 'HE - REGISTRADO' e 'HE - FIXO'.");

      const regs = Array.from(g.employees.values()).filter((e: any) => e.employeeType === 'REGISTRADO');
      const fixos = Array.from(g.employees.values()).filter((e: any) => e.employeeType === 'FIXO');
      const baseRegLayout = await getSheetLayout(targetId, 'HE - REGISTRADO', 'REGISTRADO');
      if (!baseRegLayout.slots.length) throw new Error("Não encontrei os blocos 'NOME COMPLETO' / 'SEGUNDA-FEIRA' na aba HE - REGISTRADO do MODELO RH.");
      const regSlotsPerSheet = baseRegLayout.slots.length;
      const regChunks = Math.max(1, Math.ceil(regs.length / regSlotsPerSheet));
      const fixoChunks = Math.max(1, Math.ceil(fixos.length / 10));

      const regSheets = ['HE - REGISTRADO'];
      const fixoSheets = ['HE - FIXO'];
      const duplicateRequests: any[] = [];
      for (let i = 1; i < regChunks; i++) duplicateRequests.push({ duplicateSheet: { sourceSheetId: templateReg.properties.sheetId, newSheetName: `HE - REGISTRADO ${i + 1}` } });
      for (let i = 1; i < fixoChunks; i++) duplicateRequests.push({ duplicateSheet: { sourceSheetId: templateFixo.properties.sheetId, newSheetName: `HE - FIXO ${i + 1}` } });
      if (duplicateRequests.length) await targetSheets.spreadsheets.batchUpdate({ spreadsheetId: targetId, requestBody: { requests: duplicateRequests } });
      for (let i = 1; i < regChunks; i++) regSheets.push(`HE - REGISTRADO ${i + 1}`);
      for (let i = 1; i < fixoChunks; i++) fixoSheets.push(`HE - FIXO ${i + 1}`);

      const updates: any[] = [];
      const registeredVerification: any[] = [];
      const q = (name: string) => `'${name.replace(/'/g, "''")}'`;
      const put = (sheet: string, cell: string, value: any) => updates.push({ range: `${q(sheet)}!${cell}`, values: [[value ?? '']] });

      // REGISTRADO: usa os slots descobertos no próprio modelo e valida cada
      // horário escrito. Nenhum colaborador sem horas entra na lista.
      for (let idx = 0; idx < regs.length; idx++) {
        const emp: any = regs[idx];
        const sheet = regSheets[Math.floor(idx / regSlotsPerSheet)];
        const layout = await getSheetLayout(targetId, sheet, 'REGISTRADO');
        const slot = layout.slots[idx % regSlotsPerSheet];
        if (!slot) throw new Error(`MODELO RH sem espaço REGISTRADO para ${emp.employeeName}.`);
        putHeaderByLabels(layout.matrix, sheet, put, g.sector, g.week);
        put(sheet, `${colLetter(slot.nameCol)}${slot.nameRow}`, emp.employeeName);
        const recMap = new Map(emp.records.map((r: any) => [normalizeDate(r.date), r]));
        for (let d = 0; d < 7; d++) {
          const rec: any = recMap.get(addDays(g.week, d));
          if (!rec) continue;
          const row = slot.firstDataRow + d;
          put(sheet, `C${row}`, rec.realEntry || '');
          put(sheet, `D${row}`, rec.punchEntry || '');
          put(sheet, `G${row}`, rec.punchExit || '');
          put(sheet, `H${row}`, rec.realExit || '');
          for (const [col, value] of [['C', rec.realEntry], ['D', rec.punchEntry], ['G', rec.punchExit], ['H', rec.realExit]] as any[]) {
            if (clean(value)) registeredVerification.push({ range: `${q(sheet)}!${col}${row}`, expected: clean(value), employee: emp.employeeName, date: normalizeDate(rec.date) });
          }
        }
      }

      // FIXO: 10 colaboradores por folha (5 blocos x 2 colunas).
      const fixedSlots = [
        {name:'B4', row:7, date:'A', entry:'B', exit:'C'}, {name:'I4', row:7, date:'H', entry:'I', exit:'J'},
        {name:'B16', row:19, date:'A', entry:'B', exit:'C'}, {name:'I16', row:19, date:'H', entry:'I', exit:'J'},
        {name:'B28', row:31, date:'A', entry:'B', exit:'C'}, {name:'I28', row:31, date:'H', entry:'I', exit:'J'},
        {name:'B40', row:43, date:'A', entry:'B', exit:'C'}, {name:'I40', row:43, date:'H', entry:'I', exit:'J'},
        {name:'B52', row:55, date:'A', entry:'B', exit:'C'}, {name:'I52', row:55, date:'H', entry:'I', exit:'J'}
      ];
      fixos.forEach((emp: any, idx: number) => {
        const sheet = fixoSheets[Math.floor(idx / 10)];
        const slot = fixedSlots[idx % 10];
        put(sheet, 'B1', g.sector);
        put(sheet, slot.name, emp.employeeName);
        const records = [...emp.records].filter((r: any) => r.realEntry || r.realExit).sort((a: any,b: any) => String(a.date).localeCompare(String(b.date))).slice(0,7);
        records.forEach((rec: any, d: number) => {
          const row = slot.row + d;
          const [y,m,day] = String(rec.date).slice(0,10).split('-');
          put(sheet, `${slot.date}${row}`, `${day}/${m}/${y}`);
          put(sheet, `${slot.entry}${row}`, rec.realEntry || '');
          put(sheet, `${slot.exit}${row}`, rec.realExit || '');
        });
      });
      for (const s of fixoSheets) {
        const fx = await targetSheets.spreadsheets.values.get({ spreadsheetId: targetId, range: `${q(s)}!A1:Z12`, valueRenderOption: 'FORMATTED_VALUE' });
        putHeaderByLabels(fx.data.values || [], s, put, g.sector, g.week);
      }

      // Trava de segurança: nenhum horário de REGISTRADO pode ser silenciosamente
      // descartado. Se existe no banco, ele precisa pertencer aos 7 dias da folha.
      for (const emp of regs) {
        for (const rec of emp.records) {
          const date = normalizeDate(rec.date);
          if (!date || !hasHours(rec)) continue;
          const offset = g.week === 'SEM-DATA' ? -1 : Math.round((Date.parse(date + 'T00:00:00Z') - Date.parse(g.week + 'T00:00:00Z')) / 86400000);
          if (offset < 0 || offset > 6) {
            throw new Error(`Falha de segurança no fechamento: ${emp.employeeName} possui horas em ${date}, mas a folha está na semana ${g.week}. O PDF não foi gerado para evitar perda de informação.`);
          }
        }
      }

      const expectedRegEmployees = regs.filter((e: any) => e.records.some((r: any) => hasHours(r))).length;
      const expectedRegCells = regs.reduce((n: number, e: any) => n + e.records.reduce((m: number, r: any) => m + [r.realEntry,r.punchEntry,r.punchExit,r.realExit].filter((v:any)=>clean(v)).length, 0), 0);
      if (expectedRegEmployees !== regs.length || expectedRegCells !== registeredVerification.length) {
        throw new Error(`Falha de auditoria REGISTRADO em ${g.sector}/${g.week}: esperados ${regs.length} colaboradores e ${expectedRegCells} horários; preparados ${expectedRegEmployees} colaboradores e ${registeredVerification.length} horários.`);
      }

      if (updates.length) {
        await targetSheets.spreadsheets.values.batchUpdate({
          spreadsheetId: targetId,
          requestBody: { valueInputOption: 'USER_ENTERED', data: updates }
        });
      }

      // Conferência pós-gravação: se qualquer horário de REGISTRADO não estiver
      // realmente na célula esperada, aborta o fechamento em vez de gerar um PDF
      // incompleto. Assim a falha nunca mais passa silenciosamente.
      if (registeredVerification.length) {
        const check = await targetSheets.spreadsheets.values.batchGet({
          spreadsheetId: targetId,
          ranges: registeredVerification.map((v: any) => v.range),
          valueRenderOption: 'FORMATTED_VALUE'
        });
        const got = check.data.valueRanges || [];
        for (let i = 0; i < registeredVerification.length; i++) {
          const exp = registeredVerification[i];
          const actual = clean(got[i]?.values?.[0]?.[0]);
          if (actual !== exp.expected) {
            throw new Error(`Falha de conferência: ${exp.employee} em ${exp.date} deveria ter ${exp.expected} em ${exp.range}, mas foi gravado '${actual || 'VAZIO'}'. O PDF foi bloqueado para evitar fechamento incompleto.`);
          }
        }
      }

      // Cada setor vira UM PDF por tipo, reunindo todas as folhas daquele tipo.
      if (fixos.length) results.push(await exportTypePdf(targetId, 'HE-FIXO', g.sector, g.week));
      if (regs.length) results.push(await exportTypePdf(targetId, 'HE-REGISTRADO', g.sector, g.week));
    } finally {
      // A planilha é apenas intermediária. O usuário final vê somente os PDFs.
      await drive.files.update({ fileId: targetId, requestBody: { trashed: true } }).catch(() => {});
    }
  }

  return results;
}
