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
