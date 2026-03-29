/**
 * ============================================================
 * SEÇÃO: API PARA O APLICATIVO WEB (REACT)
 * ============================================================
 */

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Lê os dados das abas
  const sectors = getSheetData(ss, "Setores");
  const employees = getSheetData(ss, "Funcionarios");
  const requests = getSheetData(ss, "Solicitacoes");
  
  return ContentService.createTextOutput(JSON.stringify({
    sectors: sectors,
    employees: employees,
    requests: requests
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    // O React envia como text/plain para evitar o CORS preflight
    const body = JSON.parse(e.postData.contents);
    
    if (body.action === "SYNC_DATABASE") {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
      // Salvar IDs das pastas se fornecidos
      if (body.data.folderRegId) {
        PropertiesService.getScriptProperties().setProperty('FOLDER_REG_ID', body.data.folderRegId);
      }
      if (body.data.folderFixoId) {
        PropertiesService.getScriptProperties().setProperty('FOLDER_FIXO_ID', body.data.folderFixoId);
      }
      
      // Atualiza as abas principais usadas pelo React
      if (body.data.isAdmin) {
        updateSheet(ss, "Setores", body.data.sectors);
        updateSheet(ss, "Funcionarios", body.data.employees);
      }
      mergeRequests(ss, "Solicitacoes", body.data.requests);
      
      // Atualiza a aba de registros detalhados (opcional, para relatórios na planilha)
      rebuildRegistrosDetalhados(ss);
      
      // Executa a lógica de distribuição de dados nas fichas se houver aprovados
      // Isso garante que a visualização na planilha fique atualizada
      const allRequests = getSheetData(ss, "Solicitacoes");
      processarHEsAprovadas(ss, allRequests);
      
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (body.action === "EXPORT_PDF") {
      exportarFolhasSextaFeira(body.data.folderRegId, body.data.folderFixoId);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function rebuildRegistrosDetalhados(ss) {
  const requests = getSheetData(ss, "Solicitacoes");
  const flattened = [];
  
  requests.forEach(req => {
    let recs = typeof req.records === 'string' ? JSON.parse(req.records) : req.records;
    if (Array.isArray(recs)) {
      recs.forEach(rec => {
        flattened.push({
          id_solicitacao: req.id,
          funcionario: req.employeeName,
          tipo: req.employeeType,
          setor: req.sectorName,
          data_semana: req.weekStarting,
          status: req.status,
          valor_total_pedido: req.calculatedValue,
          data_registro: rec.date,
          entrada_real: rec.realEntry,
          entrada_ponto: rec.punchEntry,
          saida_ponto: rec.punchExit,
          saida_real: rec.realExit,
          folga_vendida: rec.isFolgaVendida ? "SIM" : "NÃO",
          criado_em: req.createdAt,
          justificativa_edicao: req.editJustification || ''
        });
      });
    }
  });
  
  updateSheet(ss, "Registros_Detalhados", flattened);
}

function updateSheet(ss, sheetName, dataArray) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  if (!dataArray || dataArray.length === 0) {
    // Se não há dados, limpa a planilha mantendo o cabeçalho se existir
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    return;
  }
  
  // Extrai cabeçalhos
  const headers = Object.keys(dataArray[0]);
  
  // Prepara matriz de dados
  const rows = dataArray.map(obj => {
    return headers.map(h => {
      let val = obj[h];
      if (typeof val === 'object') {
        return JSON.stringify(val);
      }
      return val;
    });
  });
  
  // Limpa a planilha
  sheet.clearContents();
  
  // Escreve cabeçalhos e dados
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function mergeRequests(ss, sheetName, newRequests) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  if (!newRequests || newRequests.length === 0) return;
  
  const headers = Object.keys(newRequests[0]);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  const existingData = sheet.getDataRange().getValues();
  const sheetHeaders = existingData[0] || headers;
  const idIndex = sheetHeaders.indexOf("id");
  
  const existingIds = {};
  if (existingData.length > 1 && idIndex !== -1) {
    for (let i = 1; i < existingData.length; i++) {
      existingIds[existingData[i][idIndex]] = i + 1;
    }
  }
  
  let backupSheet = ss.getSheetByName("BACKUP");
  if (!backupSheet) {
    backupSheet = ss.insertSheet("BACKUP");
    backupSheet.getRange(1, 1, 1, sheetHeaders.length).setValues([sheetHeaders]);
  } else if (backupSheet.getLastRow() === 0) {
    backupSheet.getRange(1, 1, 1, sheetHeaders.length).setValues([sheetHeaders]);
  }
  
  newRequests.forEach(req => {
    const rowData = sheetHeaders.map(h => {
      let val = req[h];
      return typeof val === 'object' ? JSON.stringify(val) : val;
    });
    
    if (idIndex !== -1 && existingIds[req.id]) {
      // Update existing
      sheet.getRange(existingIds[req.id], 1, 1, sheetHeaders.length).setValues([rowData]);
    } else {
      // Append new
      sheet.appendRow(rowData);
      backupSheet.appendRow(rowData);
    }
  });
}

/**
 * Menu Personalizado
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Sistema HE')
    .addItem('Executar Sincronização Manual', 'testeManual')
    .addItem('Ativar Automação (Ao Editar Solicitações)', 'configuringGatilhoEdicao')
    .addSeparator()
    .addItem('🚀 Exportar Fichas Agora (Manual)', 'exportarFolhasSextaFeira')
    .addItem('🛑 FECHAMENTO SEMANAL (Salvar + Limpar Tudo)', 'executarFechamentoSemanal')
    .addToUi();
}

/**
 * ============================================================
 * SEÇÃO: FECHAMENTO SEMANAL E LIMPEZA
 * ============================================================
 */

function executarFechamentoSemanal() {
  const ui = SpreadsheetApp.getUi();
  const resposta = ui.alert('CONFIRMAÇÃO DE FECHAMENTO', 
    'Isso irá exportar as FOLHAS para o Drive e LIMPAR a aba de Solicitações. Deseja continuar?', 
    ui.ButtonSet.YES_NO);

  if (resposta !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // 1. Exporta antes de apagar os dados
    exportarFolhasSextaFeira();
    SpreadsheetApp.flush(); 
  } catch (e) {
    ui.alert("Erro durante a exportação: " + e.message);
    return;
  }

  // 2. Backup e Limpa o banco de dados (Solicitacoes)
  const abaSolicitacoes = ss.getSheetByName("Solicitacoes");
  if (abaSolicitacoes && abaSolicitacoes.getLastRow() > 1) {
    let abaBackup = ss.getSheetByName("BACKUP");
    if (!abaBackup) {
      abaBackup = ss.insertSheet("BACKUP");
      let headers = abaSolicitacoes.getRange(1, 1, 1, abaSolicitacoes.getLastColumn()).getValues();
      abaBackup.getRange(1, 1, 1, headers[0].length).setValues(headers);
    }
    
    let dadosParaBackup = abaSolicitacoes.getRange(2, 1, abaSolicitacoes.getLastRow() - 1, abaSolicitacoes.getLastColumn()).getValues();
    abaBackup.getRange(abaBackup.getLastRow() + 1, 1, dadosParaBackup.length, dadosParaBackup[0].length).setValues(dadosParaBackup);

    abaSolicitacoes.getRange(2, 1, abaSolicitacoes.getLastRow() - 1, abaSolicitacoes.getLastColumn()).clearContent();
  }

  // 3. Limpa as abas visuais (Fichas)
  const abasParaLimpar = ["HE - REGISTRADO", "HE - FIXO"];
  abasParaLimpar.forEach(nome => {
    const aba = ss.getSheetByName(nome);
    if (aba) {
      let matriz = aba.getDataRange().getValues();
      limparMatriz(matriz, nome.includes("REGISTRADO") ? "REGISTRADO" : "FIXO");
      aba.getRange(1, 1, matriz.length, matriz[0].length).setValues(matriz);
    }
  });

  ui.alert("Fechamento concluído com sucesso!");
}

/**
 * ============================================================
 * SEÇÃO: EXPORTAÇÃO E GESTÃO DE ARQUIVOS
 * ============================================================
 */

function exportarFolhasSextaFeira(paramFolderRegId, paramFolderFixoId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const PASTA_REGISTRADO_ID = paramFolderRegId || props.getProperty('FOLDER_REG_ID') || "1OGOxVmi2nEwI47HP9l48VdVBKQeJTVqm";
  const PASTA_FIXO_ID = paramFolderFixoId || props.getProperty('FOLDER_FIXO_ID') || "1RzzDCHznw97QxwDLh_qvf8NE8yKPNdWU";

  // Nome da pasta do dia (Ex: 24-02)
  const dataPasta = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM");

  function obterOuCriarSubpasta(idPastaPai, nomeSubpasta) {
    const pastaPai = DriveApp.getFolderById(idPastaPai);
    const subpastas = pastaPai.getFoldersByName(nomeSubpasta);
    return subpastas.hasNext() ? subpastas.next() : pastaPai.createFolder(nomeSubpasta);
  }

  try {
    const pReg = obterOuCriarSubpasta(PASTA_REGISTRADO_ID, dataPasta);
    processarExportacaoIndividual(ss, "HE - REGISTRADO", "REGISTRADO", pReg);
    
    const pFix = obterOuCriarSubpasta(PASTA_FIXO_ID, dataPasta);
    processarExportacaoIndividual(ss, "HE - FIXO", "FIXO", pFix);
  } catch(e) { console.error("Erro na exportação: " + e); }
}

function processarExportacaoIndividual(ss, nomeAba, tipo, pastaDestino) {
  const abaOrigem = ss.getSheetByName(nomeAba);
  if (!abaOrigem) return;

  const dados = abaOrigem.getDataRange().getValues();
  const dataCurta = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM");
  const saltoLinhas = (tipo === "REGISTRADO") ? 52 : 64; 
  
  let contadorNomes = {};

  for (let i = 0; i < dados.length; i += saltoLinhas) {
    if (i >= dados.length) break;
    let nomeArquivo = "";

    if (tipo === "FIXO") {
      // Pega o nome do setor nas primeiras linhas do bloco
      let nomeSetor = "";
      for (let s = 0; s < 5; s++) {
        if (dados[i+s] && dados[i+s][1]) {
          let val = dados[i+s][1].toString().toUpperCase().trim();
          if (val !== "" && !val.includes("NOME COMPLETO")) {
            nomeSetor = val.replace("SETOR:", "").trim();
            break;
          }
        }
      }
      if (!nomeSetor) nomeSetor = "GERAL";

      // Verifica se há dados no bloco antes de exportar
      let temDados = false;
      for (let r = 0; r < saltoLinhas; r++) {
        if (dados[i+r] && ((dados[i+r][1] && (dados[i+r][0]||"").toString().toUpperCase().includes("NOME")) || (dados[i+r][8] && (dados[i+r][7]||"").toString().toUpperCase().includes("NOME")))) {
          temDados = true; break;
        }
      }
      if (!temDados) continue;

      // Nome do arquivo para Fixo (Ex: GOVERNANÇA-24-02)
      if (!contadorNomes[nomeSetor]) {
        contadorNomes[nomeSetor] = 1;
        nomeArquivo = `${nomeSetor}-${dataCurta}`;
      } else {
        contadorNomes[nomeSetor]++;
        nomeArquivo = `${nomeSetor} (PARTE ${contadorNomes[nomeSetor]})-${dataCurta}`;
      }

    } else {
      // REGISTRADO: Extrai apenas o primeiro nome (Ex: MIKAELA & VALDIRENE)
      let raw1 = (dados[i+4] && dados[i+4][1]) ? dados[i+4][1].toString().trim() : "";
      let raw2 = (dados[i+28] && dados[i+28][1]) ? dados[i+28][1].toString().trim() : "";
      
      if (raw1 === "" && raw2 === "") continue;

      let pNome1 = raw1 !== "" ? raw1.split(" ")[0].toUpperCase() : "VAGO";
      let pNome2 = raw2 !== "" ? raw2.split(" ")[0].toUpperCase() : "VAGO";
      nomeArquivo = `${pNome1} & ${pNome2}`;

      if (!contadorNomes[nomeArquivo]) {
        contadorNomes[nomeArquivo] = 1;
      } else {
        contadorNomes[nomeArquivo]++;
        nomeArquivo += ` (${contadorNomes[nomeArquivo]})`;
      }
    }

    let numCols = (tipo === "FIXO") ? 13 : 11;
    let rangeFolha = abaOrigem.getRange(i + 1, 1, saltoLinhas, numCols); 
    gerarNovoArquivoSheets(nomeArquivo, rangeFolha, pastaDestino);
  }
}

/**
 * CRIAÇÃO DE ARQUIVO COM CÓPIA DE ALTURA E LARGURA (LAYOUT SULFITE)
 */
function gerarNovoArquivoSheets(nomeArquivo, rangeOrigem, pastaDestino) {
  const novoSS = SpreadsheetApp.create(nomeArquivo);
  const abaOrigem = rangeOrigem.getSheet();
  const abaCopiada = abaOrigem.copyTo(novoSS);
  abaCopiada.setName("Ficha_HE");
  
  const abas = novoSS.getSheets();
  if (abas.length > 1) novoSS.deleteSheet(abas[0]);
  
  const row = rangeOrigem.getRow();
  const col = rangeOrigem.getColumn();
  const rows = rangeOrigem.getNumRows();
  const cols = rangeOrigem.getNumColumns();
  
  const abaFinal = novoSS.insertSheet("Relatorio");

  // 1. Copia Largura das Colunas
  for (let c = 1; c <= cols; c++) {
    abaFinal.setColumnWidth(c, abaOrigem.getColumnWidth(col + c - 1));
  }

  // 2. Copia Altura das Linhas (Essencial para manter o Sulfite A4)
  for (let r = 1; r <= rows; r++) {
    abaFinal.setRowHeight(r, abaOrigem.getRowHeight(row + r - 1));
  }
  
  abaCopiada.getRange(row, col, rows, cols).copyTo(abaFinal.getRange(1, 1));
  novoSS.deleteSheet(abaCopiada);
  
  SpreadsheetApp.flush();
  
  let arquivo = DriveApp.getFileById(novoSS.getId());
  pastaDestino.addFile(arquivo);
  DriveApp.getRootFolder().removeFile(arquivo);
}

/**
 * ============================================================
 * SEÇÃO: LÓGICA DE SINCRONIZAÇÃO E APOIO
 * ============================================================
 */

function agruparSolicitacoesPorFuncionario(requests) {
  const agrupado = {};
  requests.forEach(req => {
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
    let recs = typeof req.records === 'string' ? JSON.parse(req.records) : req.records;
    agrupado[key].records = agrupado[key].records.concat(recs);
  });

  const resultado = [];

  Object.keys(agrupado).forEach(key => {
    let grupo = agrupado[key];
    let records = grupo.records;
    
    // Remove registros vazios e ordena por data
    records = records.filter(d => d.realEntry || d.realExit || d.punchEntry || d.punchExit);
    records.sort((a, b) => (a.date > b.date) ? 1 : -1);

    // Agrupar por semana (Segunda a Domingo) para TODOS
    const semanas = {};
    records.forEach(rec => {
      let partes = rec.date.split("-");
      let d = new Date(partes[0], partes[1] - 1, partes[2], 12, 0, 0);
      let diaSemana = d.getDay();
      let diffParaSegunda = diaSemana === 0 ? -6 : 1 - diaSemana;
      let segunda = new Date(d);
      segunda.setDate(d.getDate() + diffParaSegunda);
      let keySemana = segunda.getFullYear() + "-" + (segunda.getMonth() + 1) + "-" + segunda.getDate();
      
      if (!semanas[keySemana]) semanas[keySemana] = [];
      
      // Evitar duplicatas exatas de data na mesma semana (mantém o mais recente)
      let idx = semanas[keySemana].findIndex(r => r.date === rec.date);
      if (idx !== -1) {
        semanas[keySemana][idx] = rec;
      } else {
        semanas[keySemana].push(rec);
      }
    });
    
    Object.keys(semanas).forEach(keySemana => {
      let recsSemana = semanas[keySemana];
      if ((grupo.employeeType || "").toUpperCase().trim() === "REGISTRADO") {
        resultado.push({
          employeeName: grupo.employeeName,
          employeeType: grupo.employeeType,
          sectorName: grupo.sectorName,
          records: recsSemana
        });
      } else {
        // FIXO: Agrupar a cada 7 registros (limite da ficha) dentro da mesma semana
        for (let i = 0; i < recsSemana.length; i += 7) {
          resultado.push({
            employeeName: grupo.employeeName,
            employeeType: grupo.employeeType,
            sectorName: grupo.sectorName,
            records: recsSemana.slice(i, i + 7)
          });
        }
      }
    });
  });

  return resultado;
}

function processarHEsAprovadas(ss, requests) {
  const aprovados = requests.filter(r => (r.status || "").toUpperCase().trim() === "APROVADO");
  const agrupados = agruparSolicitacoesPorFuncionario(aprovados);

  const abaReg = ss.getSheetByName("HE - REGISTRADO");
  if (abaReg) {
    let range = abaReg.getDataRange();
    let matriz = range.getValues();
    let formulas = range.getFormulas(); 
    limparMatriz(matriz, "REGISTRADO");
    
    agrupados.filter(r => r.employeeType.toUpperCase().trim() === "REGISTRADO").forEach(req => {
      let rIdx = localizarFichaVaziaNaMatriz(matriz, 0, 1);
      if (rIdx !== -1) {
        matriz[rIdx][1] = req.employeeName;
        if (matriz[rIdx - 3]) matriz[rIdx - 3][1] = req.sectorName;
        preencherColunaAERegistros(matriz, formulas, rIdx + 7, req.records);
      }
    });
    restaurarFormulas(matriz, formulas); 
    abaReg.getRange(1, 1, matriz.length, matriz[0].length).setValues(matriz);
  }

  const abaFixo = ss.getSheetByName("HE - FIXO");
  if (abaFixo) {
    let range = abaFixo.getDataRange();
    let matriz = range.getValues();
    let formulas = range.getFormulas(); 
    limparMatriz(matriz, "FIXO");
    
    agrupados.filter(r => r.employeeType.toUpperCase().trim() === "FIXO").forEach(func => {
      let fIdx = -1; let colBase = -1; 
      let nomeSetorAlvo = (func.sectorName || "").toUpperCase().trim();
      for (let i = 0; i < matriz.length; i++) {
        if ((matriz[i][1] || "").toString().toUpperCase().trim() === nomeSetorAlvo) {
          let buscaEsq = localizarVagaNoBlocoSetor(matriz, i, 0);
          if (buscaEsq !== -1) { fIdx = buscaEsq; colBase = 0; break; }
          let buscaDir = localizarVagaNoBlocoSetor(matriz, i, 7);
          if (buscaDir !== -1) { fIdx = buscaDir; colBase = 7; break; }
        }
      }
      if (fIdx !== -1) {
        matriz[fIdx][colBase + 1] = func.employeeName;
        preencherHorasNaMatriz(matriz, formulas, fIdx + 3, func.records, colBase);
      }
    });
    restaurarFormulas(matriz, formulas);
    abaFixo.getRange(1, 1, matriz.length, matriz[0].length).setValues(matriz);
  }
}

function limparMatriz(matriz, tipo) {
  for (let i = 0; i < matriz.length; i++) {
    if (i === 13 || i === 14) continue; 
    let txtA = (matriz[i] && matriz[i][0]) ? matriz[i][0].toString().toUpperCase() : "";
    if (txtA.includes("NOME COMPLETO:")) {
      matriz[i][1] = ""; 
      let start = (tipo === "REGISTRADO") ? 7 : 3;
      let limit = (tipo === "REGISTRADO") ? 8 : 7;
      for (let g = 0; g < limit; g++) {
        let rIdx = i + start + g;
        if (matriz[rIdx]) {
          if (tipo === "REGISTRADO") [0, 2, 3, 6, 7].forEach(c => matriz[rIdx][c] = ""); 
          else [0, 1, 2, 4].forEach(c => matriz[rIdx][c] = ""); 
        }
      }
    }
    if (tipo === "FIXO" && matriz[i] && matriz[i][7] && matriz[i][7].toString().toUpperCase().includes("NOME COMPLETO:")) {
      matriz[i][8] = ""; 
      for (let g = 0; g < 7; g++) {
        let rIdx = i + 3 + g;
        if (matriz[rIdx]) [7, 8, 9, 11].forEach(c => matriz[rIdx][c] = "");
      }
    }
  }
}

function getSheetData(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0];
  return values.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (typeof val === 'string' && (val.startsWith('[') || val.startsWith('{'))) {
        try { val = JSON.parse(val); } catch(e) {}
      }
      obj[h] = val;
    });
    return obj;
  });
}

function parseRecords(recs) {
  if (typeof recs === 'string') { try { return JSON.parse(recs); } catch(e) { return []; } }
  return Array.isArray(recs) ? recs : [];
}

function preencherHorasNaMatriz(matriz, formulas, linhaInicio, records, col) {
  let horas = parseRecords(records);
  horas.sort((a, b) => (a.date > b.date) ? 1 : -1);
  let preenchidas = 0;
  let diasComDados = horas.filter(d => d.realEntry || d.realExit);
  for (let i = 0; i < diasComDados.length; i++) {
    if (preenchidas < 7) {
      let r = linhaInicio + preenchidas;
      if (matriz[r]) {
        let p = diasComDados[i].date.split("-");
        matriz[r][col] = new Date(p[0], p[1]-1, p[2], 12, 0, 0);
        matriz[r][col + 1] = diasComDados[i].realEntry || "";
        matriz[r][col + 2] = diasComDados[i].realExit || "";  
        let lE = (col === 0) ? "B" : "I"; let lS = (col === 0) ? "C" : "J";
        formulas[r][col + 4] = `=${lS}${r+1}-${lE}${r+1}`;
        preenchidas++;
      }
    }
  }
}

function preencherColunaAERegistros(matriz, formulas, linhaInicio, records) {
  let horas = parseRecords(records);
  if (horas.length === 0) return;
  horas.sort((a, b) => (a.date > b.date) ? 1 : -1);
  let partesData = horas[0].date.split("-"); 
  let dataRefOriginal = new Date(partesData[0], partesData[1] - 1, partesData[2], 12, 0, 0);
  let diaSemana = dataRefOriginal.getDay();
  let diffParaSegunda = diaSemana === 0 ? -6 : 1 - diaSemana;
  let dataRef = new Date(dataRefOriginal);
  dataRef.setDate(dataRefOriginal.getDate() + diffParaSegunda);
  const diasExtenso = ["DOMINGO", "SEGUNDA-FEIRA", "TERÇA-FEIRA", "QUARTA-FEIRA", "QUINTA-FEIRA", "SEXTA-FEIRA", "SÁBADO"];
  for (let i = 0; i < 7; i++) {
    let r = linhaInicio + i;
    if (!matriz[r]) continue;
    let dataLoop = new Date(dataRef);
    dataLoop.setDate(dataRef.getDate() + i);
    matriz[r][0] = diasExtenso[dataLoop.getDay()];
    let sBusca = Utilities.formatDate(dataLoop, Session.getScriptTimeZone(), "yyyy-MM-dd");
    let reg = horas.find(h => h.date === sBusca);
    if (reg) {
      matriz[r][2] = reg.realEntry || ""; matriz[r][3] = reg.punchEntry || "";
      matriz[r][6] = reg.punchExit || ""; matriz[r][7] = reg.realExit || "";
    }
  }
}

function restaurarFormulas(matriz, formulas) {
  for (let i = 0; i < formulas.length; i++) {
    for (let j = 0; j < formulas[i].length; j++) {
      if (formulas[i][j] && formulas[i][j].toString().startsWith("=")) {
        if (!(matriz[i][j] && matriz[i][j].toString().startsWith("="))) matriz[i][j] = formulas[i][j];
      }
    }
  }
}

function localizarFichaVaziaNaMatriz(matriz, colLabel, colNome) {
  for (let i = 0; i < matriz.length; i++) {
    if (matriz[i] && (matriz[i][colLabel] || "").toString().toUpperCase().includes("NOME COMPLETO:")) {
      if (!matriz[i][colNome] || (matriz[i][colNome] || "").toString().trim() === "") return i;
    }
  }
  return -1;
}

function localizarVagaNoBlocoSetor(matriz, linhaSetor, col) {
  let contador = 0;
  for (let i = linhaSetor; i < matriz.length; i++) {
    let txt = (matriz[i] && matriz[i][col]) ? matriz[i][col].toString().toUpperCase() : "";
    if (txt.includes("NOME COMPLETO:")) {
      if ((matriz[i][col + 1] || "").toString().trim() === "") return i;
      contador++;
      if (contador >= 5) break; 
    }
    if (i > linhaSetor && matriz[i] && (matriz[i][1] || "").toString().toUpperCase().includes("SETOR:")) break;
  }
  return -1;
}

function configuringGatilhoEdicao() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => { if (t.getHandlerFunction() === 'aoEditar') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('aoEditar').forSpreadsheet(ss).onEdit().create();
  SpreadsheetApp.getUi().alert("Automação Ativada!");
}

function aoEditar(e) {
  const nomeAbaAlvo = "Solicitacoes";
  if (e.source.getActiveSheet().getName() === nomeAbaAlvo) {
    processarHEsAprovadas(e.source, getSheetData(e.source, nomeAbaAlvo));
  }
}

function testeManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requests = getSheetData(ss, "Solicitacoes"); 
  if (requests.length > 0) {
    processarHEsAprovadas(ss, requests);
    SpreadsheetApp.getUi().alert("Sincronização concluída!");
  }
}
