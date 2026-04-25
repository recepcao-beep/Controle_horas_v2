import express from 'express';
import path from 'path';
import { google } from 'googleapis';
import cookieParser from 'cookie-parser';
import session from 'express-session';
import dotenv from 'dotenv';
import { distributeData, performClosing } from '../src/server/closing.js';

dotenv.config();

const app = express();
const PORT = 3000;

app.use(express.json());
app.use(cookieParser());

// Setup Google Auth with Service Account
const getGoogleAuth = () => {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || !process.env.GOOGLE_PRIVATE_KEY) {
    throw new Error('As credenciais da Conta de Serviço (GOOGLE_SERVICE_ACCOUNT_EMAIL e GOOGLE_PRIVATE_KEY) não estão configuradas nas Variáveis de Ambiente.');
  }
  return new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
    },
    scopes: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive'
    ],
  });
};

// Sheets API Proxy
app.get('/api/config/status', (req, res) => {
  res.json({
    isConfigured: !!(process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL && process.env.GOOGLE_PRIVATE_KEY)
  });
});

app.get('/api/config/service-account', (req, res) => {
  res.json({ email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || "Não configurado" });
});

app.get('/api/sheets/load', async (req, res) => {
  const spreadsheetId = req.query.spreadsheetId as string;
  if (!spreadsheetId) return res.status(400).json({ error: 'Spreadsheet ID is required' });

  if (!process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || !process.env.GOOGLE_PRIVATE_KEY) {
    return res.json({ 
      needsSetup: true, 
      error: 'As credenciais da Conta de Serviço não estão configuradas.' 
    });
  }

  try {
    const auth = getGoogleAuth();
    const sheets = google.sheets({ version: 'v4', auth });
    
    console.log(`[API] Loading data from spreadsheet: ${spreadsheetId}`);

    const response = await sheets.spreadsheets.get({
      spreadsheetId,
      includeGridData: false,
    });

    const sheetNames = response.data.sheets?.map(s => s.properties?.title || '') || [];
    console.log(`[API] Found sheets:`, sheetNames);
    
    const ranges = ['Setores!A:Z', 'Funcionarios!A:Z', 'Solicitacoes!A:Z', 'Config!A:Z'];
    
    // Filter ranges to only those that exist (case-insensitive and trimming spaces)
    const existingRanges = ranges.filter(r => {
      const targetName = r.split('!')[0].toLowerCase().trim();
      return sheetNames.some(sn => sn.toLowerCase().trim() === targetName);
    });
    console.log(`[API] Fetching ranges:`, existingRanges);

    if (existingRanges.length === 0) {
      console.log(`[API] No matching sheets found. Returning empty data.`);
      return res.json({ sectors: [], employees: [], requests: [], config: {} });
    }

    const valuesResponse = await sheets.spreadsheets.values.batchGet({
      spreadsheetId,
      ranges: existingRanges,
    });

    const valueRanges = valuesResponse.data.valueRanges || [];
    console.log(`[API] Received data for ${valueRanges.length} ranges.`);
    
    const parseSheet = (title: string) => {
      const targetTitle = title.toLowerCase().trim();
      const vr = valueRanges.find(v => {
        const rangeName = (v.range || '').split('!')[0].replace(/['"]/g, '').toLowerCase().trim();
        return rangeName.includes(targetTitle);
      });
      if (!vr || !vr.values || vr.values.length <= 1) {
        console.log(`[API] Sheet ${title} is empty or not found.`);
        return [];
      }
      const headers = vr.values[0];
      const parsed = vr.values.slice(1).map(row => {
        const obj: any = {};
        headers.forEach((h, i) => {
          let val = row[i];
          if (typeof val === 'string' && (val.startsWith('[') || val.startsWith('{'))) {
            try { val = JSON.parse(val); } catch(e) {}
          }
          obj[h] = val;
        });
        return obj;
      });
      console.log(`[API] Parsed ${parsed.length} rows from ${title}.`);
      return parsed;
    };

    const configData = parseSheet('Config');
    const config: any = {};
    configData.forEach((item: any) => {
      if (item.key) config[item.key] = item.value;
    });

    res.json({
      sectors: parseSheet('Setores'),
      employees: parseSheet('Funcionarios'),
      requests: parseSheet('Solicitacoes'),
      config
    });
  } catch (error: any) {
    console.error('[API] Error loading sheets data:', error);
    
    // Handle 404 specifically
    if (error.code === 404 || (error.message && error.message.includes('Requested entity was not found'))) {
      return res.status(404).json({ 
        error: 'Planilha não encontrada. Verifique se o ID ou URL da planilha está correto e se a Conta de Serviço tem permissão de acesso (Editor).' 
      });
    }
    
    res.status(500).json({ error: error.message });
  }
});

app.post('/api/sheets/sync', async (req, res) => {
  const { spreadsheetId, sectors, employees, requests, config } = req.body;

  if (!spreadsheetId) return res.status(400).json({ error: 'Spreadsheet ID is required' });

  if (!process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || !process.env.GOOGLE_PRIVATE_KEY) {
    return res.json({ success: false, needsSetup: true, error: 'As credenciais da Conta de Serviço não estão configuradas.' });
  }

  try {
    const auth = getGoogleAuth();
    const sheets = google.sheets({ version: 'v4', auth });
    
    // Ensure sheets exist
    const ssInfo = await sheets.spreadsheets.get({ spreadsheetId });
    const existingSheetNames = ssInfo.data.sheets?.map(s => s.properties?.title || '') || [];
    const requiredSheets = ['Setores', 'Funcionarios', 'Solicitacoes', 'Config', 'HE - REGISTRADO', 'HE - FIXO'];
    const sheetsToCreate = requiredSheets.filter(rs => !existingSheetNames.includes(rs));

    if (sheetsToCreate.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: sheetsToCreate.map(title => ({
            addSheet: { properties: { title } }
          }))
        }
      });
    }

    const prepareData = (data: any[]) => {
      if (!data || data.length === 0) return [];
      const headers = Object.keys(data[0]);
      const rows = data.map(item => headers.map(h => {
        const val = item[h];
        return (val !== null && typeof val === 'object') ? JSON.stringify(val) : val;
      }));
      return [headers, ...rows];
    };

    const sortedRequests = [...requests].sort((a, b) => {
      const sectorA = (a.sectorName || "").toUpperCase();
      const sectorB = (b.sectorName || "").toUpperCase();
      if (sectorA < sectorB) return -1;
      if (sectorA > sectorB) return 1;
      return (a.employeeName || "").localeCompare(b.employeeName || "");
    });

    const configRows = config ? Object.entries(config).map(([key, value]) => ({ key, value })) : [];

    const syncItems = [
      { sheet: 'Setores', range: 'Setores!A1', values: prepareData(sectors) },
      { sheet: 'Funcionarios', range: 'Funcionarios!A1', values: prepareData(employees) },
      { sheet: 'Solicitacoes', range: 'Solicitacoes!A1', values: prepareData(sortedRequests) },
      { sheet: 'Config', range: 'Config!A1', values: prepareData(configRows) }
    ];

    // Clear all target sheets first to ensure deletions are reflected
    for (const item of syncItems) {
      try {
        await sheets.spreadsheets.values.clear({
          spreadsheetId,
          range: `${item.sheet}!A:Z`,
        });
      } catch (e) {
        console.warn(`[Sync] Could not clear sheet ${item.sheet}:`, (e as any).message);
      }
    }

    // Filter out empty data for update
    const updateData = syncItems
      .filter(item => item.values.length > 0)
      .map(item => ({ range: item.range, values: item.values }));

    if (updateData.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId,
        requestBody: {
          valueInputOption: 'RAW',
          data: updateData,
        },
      });
    }

    // Distribute data to HE - REGISTRADO and HE - FIXO
    try {
      await distributeData(sheets, spreadsheetId, requests);
    } catch (e: any) {
      console.error("Error distributing data:", e.message);
      // We don't fail the whole sync if distribution fails, but we log it
    }

    res.json({ success: true });
  } catch (error: any) {
    console.error('Error syncing sheets data:', error);
    
    if (error.code === 404 || (error.message && error.message.includes('Requested entity was not found'))) {
      return res.status(404).json({ 
        success: false,
        error: 'Planilha não encontrada. Verifique se o ID ou URL da planilha está correto e se a Conta de Serviço tem permissão de acesso (Editor).' 
      });
    }
    
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post('/api/sheets/close', async (req, res) => {
  const { spreadsheetId } = req.body;
  if (!spreadsheetId) return res.status(400).json({ error: 'Spreadsheet ID is required' });

  if (!process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || !process.env.GOOGLE_PRIVATE_KEY) {
    return res.json({ success: false, needsSetup: true, error: 'As credenciais da Conta de Serviço não estão configuradas.' });
  }

  try {
    const auth = getGoogleAuth();
    const sheets = google.sheets({ version: 'v4', auth });
    
    await performClosing(sheets, spreadsheetId);
    
    res.json({ success: true });
  } catch (error: any) {
    console.error('Error performing closing:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post('/api/sheets/action', async (req, res) => {
  const { scriptUrl, action, data } = req.body;
  if (!scriptUrl) return res.status(400).json({ error: 'Script URL is required' });

  if (!process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || !process.env.GOOGLE_PRIVATE_KEY) {
    return res.json({ success: false, needsSetup: true, error: 'As credenciais da Conta de Serviço não estão configuradas.' });
  }

  try {
    const auth = getGoogleAuth();
    const client = await auth.getClient();
    const token = await client.getAccessToken();

    const response = await fetch(scriptUrl, {
      method: 'POST',
      headers: { 
        'Content-Type': 'text/plain',
        'Authorization': `Bearer ${token.token}`
      },
      body: JSON.stringify({ action, data }),
    });

    const result = await response.json();
    res.json(result);
  } catch (error: any) {
    console.error('Error proxying Apps Script action:', error);
    res.status(500).json({ error: error.message });
  }
});

app.get('/api/drive/files', async (req, res) => {
  const folderId = req.query.folderId as string;
  if (!folderId) return res.status(400).json({ error: 'Folder ID is required' });

  if (!process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || !process.env.GOOGLE_PRIVATE_KEY) {
    return res.json({ success: false, needsSetup: true, error: 'As credenciais da Conta de Serviço não estão configuradas.' });
  }

  try {
    const auth = getGoogleAuth();
    const drive = google.drive({ version: 'v3', auth });

    const response = await drive.files.list({
      q: `'${folderId}' in parents and trashed = false`,
      fields: 'files(id, name, mimeType, webViewLink, iconLink, thumbnailLink)',
      pageSize: 1000,
      orderBy: 'folder,name',
    });

    const allFiles = response.data.files?.map(f => ({
      id: f.id,
      name: f.name,
      type: f.mimeType === 'application/vnd.google-apps.folder' ? 'folder' : 'file',
      url: f.webViewLink,
      icon: f.iconLink,
      thumbnail: f.thumbnailLink
    })) || [];

    const folders = allFiles.filter(f => f.type === 'folder');
    const regularFiles = allFiles.filter(f => f.type === 'file');

    res.json({ success: true, data: { folders, files: regularFiles } });
  } catch (error: any) {
    console.error('Error listing drive files:', error);
    res.status(500).json({ error: error.message });
  }
});

app.get('/api/drive/download/:fileId', async (req, res) => {
  const fileId = req.params.fileId;
  if (!fileId) return res.status(400).json({ error: 'File ID is required' });

  try {
    const auth = getGoogleAuth();
    const drive = google.drive({ version: 'v3', auth });

    // First check if it's a Google Workspace document or a binary file
    const fileMeta = await drive.files.get({ fileId, fields: 'mimeType, name' });
    const mimeType = fileMeta.data.mimeType;
    const isWorkspace = mimeType?.startsWith('application/vnd.google-apps.');

    res.setHeader('Content-Disposition', `inline; filename="${encodeURIComponent(fileMeta.data.name || 'document.pdf')}"`);

    if (mimeType === 'application/vnd.google-apps.spreadsheet') {
      res.setHeader('Content-Type', 'application/pdf');
      
      const client = await auth.getClient();
      const tokenResponse = await client.getAccessToken();
      const token = tokenResponse.token;
      
      // Parâmetros para ajustar o recorte da folha:
      // fitw=true (ajustar à largura), size=A4, portrait=true (retrato), margins
      const exportUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=pdf&portrait=true&size=A4&fitw=true&gridlines=false&top_margin=0.25&bottom_margin=0.25&left_margin=0.25&right_margin=0.25`;
      
      const fetchResponse = await fetch(exportUrl, {
        headers: {
          'Authorization': `Bearer ${token}`
        }
      });
      
      if (!fetchResponse.ok) {
        throw new Error(`Failed to export PDF: ${fetchResponse.statusText}`);
      }
      
      const arrayBuffer = await fetchResponse.arrayBuffer();
      const buffer = Buffer.from(arrayBuffer);
      res.send(buffer);
      
    } else if (isWorkspace) {
      // Export as PDF
      res.setHeader('Content-Type', 'application/pdf');
      const response = await drive.files.export({
        fileId,
        mimeType: 'application/pdf'
      }, { responseType: 'stream' });
      
      response.data.pipe(res);
    } else {
      // Download directly
      res.setHeader('Content-Type', mimeType || 'application/octet-stream');
      const response = await drive.files.get({
        fileId,
        alt: 'media'
      }, { responseType: 'stream' });
      
      response.data.pipe(res);
    }
  } catch (error: any) {
    console.error('Error downloading drive file:', error);
    res.status(500).json({ error: error.message });
  }
});

// Vite middleware (Only for local development)
if (!process.env.VERCEL) {
  async function setupVite() {
    if (process.env.NODE_ENV !== 'production') {
      const { createServer: createViteServer } = await import('vite');
      const vite = await createViteServer({
        server: { middlewareMode: true },
        appType: 'spa',
      });
      app.use(vite.middlewares);
    } else {
      const distPath = path.join(process.cwd(), 'dist');
      app.use(express.static(distPath));
      app.get('*', (req, res) => {
        res.sendFile(path.join(distPath, 'index.html'));
      });
    }

    app.listen(PORT, '0.0.0.0', () => {
      console.log(`Server running on http://localhost:${PORT}`);
    });
  }
  setupVite();
}

export default app;
