
import React, { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { 
  Settings, 
  Play, 
  ArrowLeft, 
  LayoutDashboard, 
  Users, 
  MapPin, 
  ClipboardList, 
  CheckCircle, 
  XCircle, 
  LogOut, 
  ChevronRight, 
  ChevronDown,
  ChevronUp,
  Edit2, 
  Database, 
  RefreshCw, 
  Share2, 
  Lock, 
  Copy,
  Clock,
  AlertCircle,
  TrendingUp,
  DollarSign,
  Calendar,
  Menu,
  Moon,
  Sun,
  Folder,
  FileText,
  Grid,
  List,
  Printer
} from 'lucide-react';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  Tooltip, 
  ResponsiveContainer,
  PieChart, 
  Pie, 
  Cell,
  Legend,
  CartesianGrid
} from 'recharts';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { 
  EmployeeType, 
  RequestStatus, 
  Sector, 
  Employee, 
  TimeRequest, 
  AppState,
  TimeRecord 
} from './types';
import { 
  formatCurrency, 
  parseCurrency,
  timeToDecimal,
  formatDecimalHours,
  getWeekDays 
} from './utils';

// Constantes
const STORAGE_KEY = 'controle_horas_db_v3';
const DEFAULT_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1Ksam1nwTxzveH0BaWKftQnyDBuOBJYCX5A3FwmP9EnY/edit#gid=67462249';
const COLORS = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899'];

const App: React.FC = () => {
  // --- Estados do Banco de Dados ---
  const [sectors, setSectors] = useState<Sector[]>([]);
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [requests, setRequests] = useState<TimeRequest[]>([]);
  const [dbUrl, setDbUrl] = useState(DEFAULT_SHEET_URL);
  const [scriptUrl, setScriptUrl] = useState('https://script.google.com/macros/s/AKfycby3TJGwqYjXMWyysi7CgHeNI48Kpmr8fCeFv3Ozr7252RvxdcNwlaNoRkgJYZHb2Il8/exec');
  const [folderRegId, setFolderRegId] = useState('1OGOxVmi2nEwI47HP9l48VdVBKQeJTVqm');
  const [folderFixoId, setFolderFixoId] = useState('1RzzDCHznw97QxwDLh_qvf8NE8yKPNdWU');
  
  const extractFolderId = (input: string) => {
    if (!input) return '';
    const match = input.match(/[-\w]{15,}/);
    return match ? match[0] : input;
  };
  const [isInitialLoad, setIsInitialLoad] = useState(true);
  const [isSyncing, setIsSyncing] = useState(false);
  const [driveFiles, setDriveFiles] = useState<{folders: any[], files: any[]}>({folders: [], files: []});
  const [currentFolderId, setCurrentFolderId] = useState<string | null>(null);
  const [folderHistory, setFolderHistory] = useState<{id: string, name: string}[]>([]);
  const [isLoadingFiles, setIsLoadingFiles] = useState(false);
  const [alertMessage, setAlertMessage] = useState<string | null>(null);
  const [confirmDialog, setConfirmDialog] = useState<{ message: string, onConfirm: () => void } | null>(null);
  const [fileViewMode, setFileViewMode] = useState<'list' | 'grid'>('list');
  const [folderCache, setFolderCache] = useState<Record<string, {folders: any[], files: any[]}>>({});
  const hasLoadedRef = useRef(false);

  // --- Estado Global da Navegação ---
  const [state, setState] = useState<AppState>({
    view: 'HOME',
    flowType: null,
    adminSubView: 'DASHBOARD'
  });

  // --- Estados de Formulários ---
  const [adminPassword, setAdminPassword] = useState('');
  const [isAuth, setIsAuth] = useState(false);
  const [selectedSector, setSelectedSector] = useState<string>('');
  const [selectedEmployee, setSelectedEmployee] = useState<string>('');
  const [showFormModal, setShowFormModal] = useState(false);
  const [currentWeek, setCurrentWeek] = useState(new Date().toISOString().split('T')[0]);

  // Estados Admin
  const [newSec, setNewSec] = useState({ name: '', fixedRate: 0 });
  const [newEmpData, setNewEmpData] = useState({ name: '', sectorId: '', salary: 0, monthlyHours: 220, type: EmployeeType.REGISTRADO });
  const [employeeSectorFilter, setEmployeeSectorFilter] = useState<string>('ALL');
  // Estado para controlar qual funcionário está sendo editado
  const [editingEmployeeId, setEditingEmployeeId] = useState<string | null>(null);
  
  const [modalRecords, setModalRecords] = useState<TimeRecord[]>([]);

  // Edição e Controle de Acesso
  const [editingRequestId, setEditingRequestId] = useState<string | null>(null);
  const [editJustification, setEditJustification] = useState('');
  
  const extractSpreadsheetId = (input: string) => {
    if (!input) return '';
    const match = input.match(/\/d\/([-\w]{25,})/) || input.match(/^([-\w]{25,})$/);
    return match ? (match[1] || match[0]) : input;
  };

  const [isDarkMode, setIsDarkMode] = useState(() => {
    const saved = localStorage.getItem('theme');
    if (saved) return saved === 'dark';
    return window.matchMedia('(prefers-color-scheme: dark)').matches;
  });

  useEffect(() => {
    if (isDarkMode) {
      document.documentElement.classList.add('dark');
      localStorage.setItem('theme', 'dark');
    } else {
      document.documentElement.classList.remove('dark');
      localStorage.setItem('theme', 'light');
    }
  }, [isDarkMode]);

  const [isServiceAccountSetup, setIsServiceAccountSetup] = useState(true);
  const lastSyncedDataRef = useRef<string>('{"sectors":[],"employees":[],"requests":[]}');
  const [generatedLink, setGeneratedLink] = useState('');

  const exportToExcel = useCallback(() => {
    try {
      const dataToExport = requests.map(req => {
        const employee = employees.find(e => e.id === req.employeeId);
        const sector = sectors.find(s => s.id === req.sectorId);
        
        return {
          'Data': req.date,
          'Funcionário': employee?.name || 'N/A',
          'Setor': sector?.name || 'N/A',
          'Tipo': req.employeeType,
          'Horas': req.hours,
          'Valor Calculado': req.calculatedValue,
          'Status': req.status,
          'Justificativa': req.justification || '',
          'ID': req.id
        };
      });

      const worksheet = XLSX.utils.json_to_sheet(dataToExport);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Solicitacoes");

      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const data = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8' });
      saveAs(data, `Relatorio_Horas_${new Date().toISOString().split('T')[0]}.xlsx`);
      setAlertMessage("Relatório Excel gerado com sucesso!");
    } catch (error) {
      console.error("Erro ao exportar Excel:", error);
      setAlertMessage("Erro ao gerar relatório Excel.");
    }
  }, [requests, employees, sectors]);

  // --- Lógica de Dashboard (useMemo) ---
  const filteredAndSortedEmployees = useMemo(() => {
    let filtered = employees;
    if (employeeSectorFilter !== 'ALL') {
      filtered = filtered.filter(e => e.sectorId === employeeSectorFilter);
    }
    return filtered.sort((a, b) => a.name.localeCompare(b.name));
  }, [employees, employeeSectorFilter]);

  const dashboardData = useMemo(() => {
    const approved = requests.filter(r => r.status === RequestStatus.APROVADO);
    
    // Total Geral Gasto
    const totalSpent = approved.reduce((acc, curr) => acc + parseCurrency(curr.calculatedValue), 0);
    
    // Dados por Setor
    const expensesBySector = sectors.map(sector => {
      const value = approved
        .filter(r => r.sectorId === sector.id)
        .reduce((acc, curr) => acc + parseCurrency(curr.calculatedValue), 0);
      return { name: sector.name, value };
    }).filter(item => item.value > 0).sort((a, b) => b.value - a.value);

    // Dados por Tipo (Registrado vs Fixo)
    const expensesByType = [
      { 
        name: 'Registrado', 
        value: approved.filter(r => r.employeeType === EmployeeType.REGISTRADO).reduce((acc, curr) => acc + parseCurrency(curr.calculatedValue), 0),
        color: '#2563eb' // Blue-600
      },
      { 
        name: 'Fixo', 
        value: approved.filter(r => r.employeeType === EmployeeType.FIXO).reduce((acc, curr) => acc + parseCurrency(curr.calculatedValue), 0),
        color: '#16a34a' // Green-600
      }
    ].filter(i => i.value > 0);

    return { totalSpent, expensesBySector, expensesByType, approvedCount: approved.length };
  }, [requests, sectors]);

  // --- Lógica de Sincronização de Dados ---

  const loadDatabase = useCallback(async (urlToUse?: string) => {
    const targetUrl = urlToUse || dbUrl;
    if (!targetUrl) {
      setIsInitialLoad(false);
      return;
    }

    setIsSyncing(true);
    try {
      const spreadsheetId = extractSpreadsheetId(targetUrl);
      const response = await fetch(`/api/sheets/load?spreadsheetId=${spreadsheetId}`);
      const data = await response.json();
      
      if (data && data.needsSetup) {
        setIsServiceAccountSetup(false);
        throw new Error(data.error);
      } else {
        setIsServiceAccountSetup(true);
      }
      
      if (data && !data.error) {
        const uniqueSectors = Array.from(new Map((data.sectors || []).map((s: any) => [s.id, s])).values()) as Sector[];
        const uniqueEmployees = Array.from(new Map((data.employees || []).map((e: any) => [e.id, e])).values()) as Employee[];
        const activeRequests = (data.requests || []).filter((r: TimeRequest) => r.status !== RequestStatus.DELETADO);
        const uniqueRequests = Array.from(new Map(activeRequests.map((r: any) => [r.id, r])).values()) as TimeRequest[];

        setSectors(uniqueSectors);
        setEmployees(uniqueEmployees);
        setRequests(uniqueRequests);
        if (urlToUse) setDbUrl(urlToUse);
        
        lastSyncedDataRef.current = JSON.stringify({
          sectors: uniqueSectors,
          employees: uniqueEmployees,
          requests: uniqueRequests
        });

        localStorage.setItem(STORAGE_KEY, JSON.stringify({
          sectors: uniqueSectors,
          employees: uniqueEmployees,
          requests: uniqueRequests,
          dbUrl: targetUrl,
          scriptUrl,
          folderRegId,
          folderFixoId
        }));
      } else if (data && data.error) {
        throw new Error(data.error);
      }
    } catch (error: any) {
      if (error.message && error.message.includes('Conta de Serviço')) {
        setAlertMessage(error.message);
      } else if (error.message && error.message.includes('Planilha não encontrada')) {
        setAlertMessage(error.message);
      } else {
        console.error("Erro ao carregar dados:", error);
        setAlertMessage("Erro ao comunicar com o servidor. Verifique os logs para mais detalhes.");
      }
      
      // Fallback para localStorage
      const localData = localStorage.getItem(STORAGE_KEY);
      if (localData) {
        try {
          const parsed = JSON.parse(localData);
          const uniqueSectors = Array.from(new Map((parsed.sectors || []).map((s: any) => [s.id, s])).values()) as Sector[];
          const uniqueEmployees = Array.from(new Map((parsed.employees || []).map((e: any) => [e.id, e])).values()) as Employee[];
          const activeLocalRequests = (parsed.requests || []).filter((r: TimeRequest) => r.status !== RequestStatus.DELETADO);
          const uniqueRequests = Array.from(new Map(activeLocalRequests.map((r: any) => [r.id, r])).values()) as TimeRequest[];

          setSectors(uniqueSectors);
          setEmployees(uniqueEmployees);
          setRequests(uniqueRequests);
          if (parsed.folderRegId) setFolderRegId(parsed.folderRegId);
          if (parsed.folderFixoId) setFolderFixoId(parsed.folderFixoId);
          
          lastSyncedDataRef.current = JSON.stringify({
            sectors: uniqueSectors,
            employees: uniqueEmployees,
            requests: uniqueRequests
          });
        } catch (e) {
          // Ignora erro de parse local
          lastSyncedDataRef.current = JSON.stringify({
            sectors: [],
            employees: [],
            requests: []
          });
        }
      } else {
        lastSyncedDataRef.current = JSON.stringify({
          sectors: [],
          employees: [],
          requests: []
        });
      }
    } finally {
      setIsSyncing(false);
      setIsInitialLoad(false);
    }
  }, [dbUrl, scriptUrl, folderRegId, folderFixoId]);

  const executarFechamentoSemanal = async () => {
    setConfirmDialog({
      message: 'CONFIRMAÇÃO DE FECHAMENTO\n\nIsso irá duplicar as FOLHAS (REGISTRADO e FIXO) como backup, salvar as solicitações na aba BACKUP e LIMPAR a aba de Solicitações e as Fichas. Deseja continuar?',
      onConfirm: async () => {
        setIsSyncing(true);
        try {
          const response = await fetch('/api/sheets/close', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              spreadsheetId: extractSpreadsheetId(dbUrl)
            }),
          });
          
          const result = await response.json();
          if (result.success) {
            setAlertMessage("Fechamento concluído com sucesso!");
            loadDatabase(); // Reload data to clear requests locally
          } else {
            setAlertMessage("Erro no fechamento: " + (result.error || "Verifique se a Conta de Serviço tem permissão de acesso."));
          }
        } catch (error) {
          console.error("Erro ao executar fechamento:", error);
          setAlertMessage("Falha na comunicação com o servidor.");
        } finally {
          setIsSyncing(false);
        }
      }
    });
  };

  const fetchDriveFiles = async (folderId: string, folderName: string) => {
    // Check cache first for faster navigation
    if (folderCache[folderId]) {
      setDriveFiles(folderCache[folderId]);
      setCurrentFolderId(folderId);
      setFolderHistory(prev => {
        const index = prev.findIndex(f => f.id === folderId);
        if (index !== -1) {
          return prev.slice(0, index + 1);
        } else {
          return [...prev, { id: folderId, name: folderName }];
        }
      });
      return;
    }

    setIsLoadingFiles(true);
    try {
      const response = await fetch(`/api/drive/files?folderId=${folderId}`);
      const result = await response.json();
      
      if (result.success) {
        setDriveFiles(result.data);
        setCurrentFolderId(folderId);
        setFolderCache(prev => ({ ...prev, [folderId]: result.data }));
        setFolderHistory(prev => {
          const index = prev.findIndex(f => f.id === folderId);
          if (index !== -1) return prev.slice(0, index + 1);
          return [...prev, { id: folderId, name: folderName }];
        });
      } else {
        // Obter o email da conta de serviço para instruir o usuário
        const configRes = await fetch('/api/config/service-account');
        const config = await configRes.json();
        const saEmail = config.email;
        
        setAlertMessage(`Erro ao carregar arquivos: ${result.error || "Acesso negado"}.\n\nCertifique-se de que a pasta no Google Drive foi compartilhada com o email da Conta de Serviço:\n\n${saEmail}`);
      }
    } catch (error) {
      console.error("Erro ao listar arquivos:", error);
      setAlertMessage("Falha na comunicação com o servidor ao carregar arquivos.");
    } finally {
      setIsLoadingFiles(false);
    }
  };

  const printFile = (fileId: string) => {
    // Abre o arquivo em uma nova aba para visualização/impressão.
    // Evita o uso de iframe oculto que causa erros de cross-origin (CORS) 
    // com o visualizador de PDF nativo do navegador e bloqueio de pop-ups.
    const printUrl = `/api/drive/download/${fileId}`;
    window.open(printUrl, '_blank');
  };

  const exportToPDF = async () => {
    if (!scriptUrl) {
      setAlertMessage("Configure a URL do Apps Script primeiro.");
      return;
    }
    
    setIsSyncing(true);
    try {
      const response = await fetch('/api/sheets/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          scriptUrl: scriptUrl,
          action: "EXPORT_PDF",
          data: { folderRegId, folderFixoId }
        }),
      });
      
      const result = await response.json();
      if (result.success) {
        setAlertMessage("Fichas exportadas com sucesso!");
      } else {
        setAlertMessage("Erro ao exportar: " + (result.error || "Verifique se as pastas do Google Drive estão configuradas corretamente e se você tem permissão de acesso."));
      }
    } catch (error) {
      console.error("Erro na exportação:", error);
      setAlertMessage("Erro ao comunicar com o servidor.");
    } finally {
      setIsSyncing(false);
    }
  };

  const syncDatabase = useCallback(async (currentData: { sectors: Sector[], employees: Employee[], requests: TimeRequest[] }) => {
    if (!dbUrl) return;

    setIsSyncing(true);

    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({ ...currentData, dbUrl, scriptUrl, folderRegId, folderFixoId }));
      
      const spreadsheetId = extractSpreadsheetId(dbUrl);
      const response = await fetch('/api/sheets/sync', {
        method: 'POST',
        headers: { 
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'Cache-Control': 'no-cache'
        },
        body: JSON.stringify({
          spreadsheetId,
          sectors: currentData.sectors,
          employees: currentData.employees,
          requests: currentData.requests
        })
      });
      
      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error || `Erro do servidor: ${response.status}`);
      }

      const result = await response.json();
      if (!result.success) throw new Error(result.error || "Erro na sincronização via API");
    } catch (error: any) {
      console.error("Erro detalhado de sincronização:", error);
      
      if (error.message && error.message.includes('Conta de Serviço')) {
        setAlertMessage(error.message);
      } else if (error.message && error.message.includes('Planilha não encontrada')) {
        setAlertMessage(error.message);
      } else {
        // No iPhone, queremos ver o erro real se falhar
        setAlertMessage(`Erro de Sincronização: ${error.message}. Verifique sua conexão ou se o app está em modo privado.`);
      }
      // Falha silenciosa na sincronização para outros erros, dados já estão no localStorage
    } finally {
      setIsSyncing(false);
    }
  }, [dbUrl, scriptUrl, folderRegId, folderFixoId, isAuth]);

  useEffect(() => {
    if (hasLoadedRef.current) return;
    hasLoadedRef.current = true;

    const params = new URLSearchParams(window.location.search);
    const token = params.get('t');

    if (token) {
      const timestamp = parseInt(token, 10);
      const now = Date.now();
      if (isNaN(timestamp) || (now - timestamp > 24 * 60 * 60 * 1000)) {
        setState(prev => ({ ...prev, view: 'EXPIRED' }));
        return;
      }
    }
    
    const localData = localStorage.getItem(STORAGE_KEY);
    let initialUrl = DEFAULT_SHEET_URL;
    if (localData) {
      try {
        const parsed = JSON.parse(localData);
        if (parsed.scriptUrl) setScriptUrl(parsed.scriptUrl);
        if (parsed.folderRegId) setFolderRegId(parsed.folderRegId);
        if (parsed.folderFixoId) setFolderFixoId(parsed.folderFixoId);
        if (parsed.dbUrl) {
          // Se a URL salva for a antiga, forçamos a nova
          if (parsed.dbUrl.includes('1HRQ3L-iU-nYMKvKc3zcGoOw70uz8Sf0vfOiseMwlFWY')) {
            initialUrl = DEFAULT_SHEET_URL;
            setDbUrl(DEFAULT_SHEET_URL);
          } else {
            initialUrl = parsed.dbUrl;
            setDbUrl(parsed.dbUrl);
          }
        }
      } catch (e) {}
    }
    
    loadDatabase(initialUrl);
  }, [loadDatabase]);

  useEffect(() => {
    if (!isInitialLoad && state.view !== 'EXPIRED') {
      const currentDataString = JSON.stringify({ sectors, employees, requests });
      if (currentDataString !== lastSyncedDataRef.current) {
        const timer = setTimeout(() => {
          syncDatabase({ sectors, employees, requests });
          lastSyncedDataRef.current = currentDataString;
        }, 1500);
        return () => clearTimeout(timer);
      }
    }
  }, [sectors, employees, requests, isInitialLoad, syncDatabase, state.view]);

  useEffect(() => {
    if (showFormModal && !editingRequestId) {
      const weekDays = getWeekDays(new Date(currentWeek));
      setModalRecords(weekDays.map(date => ({
        date, realEntry: '', punchEntry: '', punchExit: '', realExit: '', isFolgaVendida: false
      })));
    }
  }, [showFormModal, currentWeek, editingRequestId]);

  useEffect(() => {
    if (!isInitialLoad) {
      const currentData = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
      localStorage.setItem(STORAGE_KEY, JSON.stringify({
        ...currentData,
        dbUrl,
        scriptUrl,
        folderRegId,
        folderFixoId
      }));
    }
  }, [dbUrl, scriptUrl, folderRegId, folderFixoId, isInitialLoad]);

  // --- Handlers ---
  const handleAdminLogin = () => {
    // any password allows entry - as requested to "remove protection"
    setIsAuth(true);
    setState(prev => ({ ...prev, view: 'ADMIN' }));
    setAdminPassword('');
  };

  const generateAccessLink = () => {
    const timestamp = Date.now();
    const link = `${window.location.origin}${window.location.pathname}?t=${timestamp}`;
    setGeneratedLink(link);
    navigator.clipboard.writeText(link);
    setAlertMessage('Link de acesso válido por 24h copiado!');
  };

  const submitRequest = () => {
    let targetEmployeeId = selectedEmployee;
    let targetSectorId = selectedSector;
    let targetFlowType = state.flowType;

    if (editingRequestId) {
      const originalReq = requests.find(r => r.id === editingRequestId);
      if (originalReq) {
        targetEmployeeId = originalReq.employeeId;
        targetSectorId = originalReq.sectorId;
        targetFlowType = originalReq.employeeType;
      }
    }

    const employee = employees.find(e => String(e.id) === String(targetEmployeeId));
    const sector = sectors.find(s => String(s.id) === String(employee?.sectorId || targetSectorId));
    
    if (targetFlowType === EmployeeType.REGISTRADO && !employee) {
      setAlertMessage("Erro: Dados do funcionário não encontrados para recálculo.");
      return;
    }

    let totalDiffHours = 0;
    let totalPayment = 0;

    if (targetFlowType === EmployeeType.REGISTRADO && employee) {
      const salary = parseCurrency(employee.salary);
      const monthlyHours = parseFloat(String(employee.monthlyHours)) || 220;
      const hourlyBase = salary / monthlyHours;
      const overtimeRate = hourlyBase * 1.25;
      
      modalRecords.forEach(r => {
        let dailyHours = 0;

        if (r.isFolgaVendida) {
          if (r.realEntry && r.realExit) {
             const start = timeToDecimal(r.realEntry);
             const end = timeToDecimal(r.realExit);
             let diff = end - start;
             if (diff < 0) diff += 24;
             dailyHours += diff;
          }
        } else {
          if (r.realEntry && r.punchEntry) {
            const real = timeToDecimal(r.realEntry);
            const punch = timeToDecimal(r.punchEntry);
            if (real < punch) {
              dailyHours += (punch - real);
            }
          }
          if (r.realExit && r.punchExit) {
            const real = timeToDecimal(r.realExit);
            const punch = timeToDecimal(r.punchExit);
            if (real > punch) {
              dailyHours += (real - punch);
            }
          }
        }
        
        if (dailyHours > 0) {
          totalDiffHours += dailyHours;
          // HE-REGISTRADO não tem vale transporte (+R$12,00 por dia) só o HE-FIXO
          totalPayment += (dailyHours * overtimeRate) + (targetFlowType === EmployeeType.FIXO ? 12 : 0);
        }
      });
    } else {
      const hourlyRate = parseCurrency(sector?.fixedRate);
      modalRecords.forEach(r => { 
        if (r.realEntry && r.realExit) {
          const start = timeToDecimal(r.realEntry);
          const end = timeToDecimal(r.realExit);
          let dailyHours = end - start;
          if (dailyHours < 0) dailyHours += 24;
          if (dailyHours > 0) {
            // HE-FIXO tem vale transporte (+R$12,00 por dia)
            totalPayment += (dailyHours * hourlyRate) + 12;
            totalDiffHours += dailyHours;
          }
        }
      });
    }

    if (editingRequestId) {
      setRequests(requests.map(r => r.id === editingRequestId ? {
        ...r, 
        records: modalRecords, 
        calculatedValue: totalPayment, 
        totalTimeDecimal: totalDiffHours,
        editJustification 
      } : r));
      setEditingRequestId(null);
      setEditJustification('');
    } else {
      const newReq: TimeRequest = {
        id: Math.random().toString(36).substr(2, 9),
        employeeId: employee?.id || 'fixo-' + Date.now(),
        employeeName: employee?.name || selectedEmployee || 'Colaborador Fixo',
        employeeType: targetFlowType!,
        sectorId: sector?.id || '',
        sectorName: sector?.name || '',
        weekStarting: currentWeek,
        records: modalRecords,
        status: RequestStatus.PENDENTE,
        calculatedValue: totalPayment,
        totalTimeDecimal: totalDiffHours,
        createdAt: new Date().toISOString()
      };
      setRequests([newReq, ...requests]);
      setState(prev => ({ ...prev, view: 'SUCCESS' }));
    }
    setShowFormModal(false);
  };

  // --- UI Components ---

  const RequestCard: React.FC<{ req: TimeRequest }> = ({ req }) => {
    const [isExpanded, setIsExpanded] = useState(false);

    const getDailyHours = (r: TimeRecord, type: EmployeeType) => {
        let total = 0;
        if (type === EmployeeType.FIXO) {
            if (r.realEntry && r.realExit) {
                let diff = timeToDecimal(r.realExit) - timeToDecimal(r.realEntry);
                if (diff < 0) diff += 24;
                total = diff;
            }
        } else {
            if (r.isFolgaVendida) {
                 if (r.realEntry && r.realExit) {
                    let diff = timeToDecimal(r.realExit) - timeToDecimal(r.realEntry);
                    if (diff < 0) diff += 24;
                    total = diff;
                 }
            } else {
                 if (r.realEntry && r.punchEntry) {
                    const diff = timeToDecimal(r.punchEntry) - timeToDecimal(r.realEntry);
                    if(diff > 0) total += diff;
                 }
                 if (r.realExit && r.punchExit) {
                    const diff = timeToDecimal(r.realExit) - timeToDecimal(r.punchExit);
                    if(diff > 0) total += diff;
                 }
            }
        }
        return total > 0 ? formatDecimalHours(total) : '-';
    };

    return (
      <div className="bg-white dark:bg-gray-800 p-4 rounded-2xl border border-gray-100 dark:border-gray-700 shadow-sm hover:shadow-md transition mb-3 group">
        <div className="flex justify-between items-start mb-2">
          <span className={`px-2 py-0.5 rounded text-[9px] font-black uppercase ${req.employeeType === EmployeeType.REGISTRADO ? 'bg-blue-100 dark:bg-blue-900/30 text-blue-700 dark:text-blue-400' : 'bg-green-100 dark:bg-green-900/30 text-green-700 dark:text-green-400'}`}>
            {req.employeeType}
          </span>
          <div className="text-right">
            <p className="text-[10px] text-gray-500 dark:text-gray-400 font-bold mb-0.5">{formatDecimalHours(req.totalTimeDecimal)}</p>
            <p className="text-sm font-black text-gray-900 dark:text-white">{formatCurrency(req.calculatedValue)}</p>
          </div>
        </div>
        <h4 className="text-sm font-bold text-gray-800 dark:text-gray-200 line-clamp-1">{req.employeeName}</h4>
        <p className="text-[10px] text-gray-400 dark:text-gray-500 mb-3">{req.sectorName} • Sem: {new Date(req.weekStarting).toLocaleDateString('pt-BR')}</p>
        
        {isExpanded && (
            <div className="mt-2 mb-4 bg-gray-50 dark:bg-gray-900 rounded-xl p-2 overflow-x-auto">
                <table className="w-full text-[10px] text-left">
                    <thead>
                        <tr className="text-gray-400 dark:text-gray-500 border-b border-gray-200 dark:border-gray-700">
                            <th className="pb-1 font-semibold">Dia</th>
                            <th className="pb-1 font-semibold">Ent.</th>
                            <th className="pb-1 font-semibold text-gray-300 dark:text-gray-600">P.Ent</th>
                            <th className="pb-1 font-semibold text-gray-300 dark:text-gray-600">P.Sai</th>
                            <th className="pb-1 font-semibold">Sai.</th>
                            <th className="pb-1 font-semibold text-right">H.</th>
                        </tr>
                    </thead>
                    <tbody>
                        {req.records.map((r, idx) => (
                            <tr key={idx} className={`border-b border-gray-100 dark:border-gray-800 last:border-0 ${r.isFolgaVendida ? 'bg-blue-50/50 dark:bg-blue-900/10' : ''}`}>
                                <td className="py-1.5 font-bold text-gray-600 dark:text-gray-400">
                                    {(() => {
                                        // Safari friendly date parsing
                                        const [year, month, day] = r.date.split('-');
                                        const d = new Date(Number(year), Number(month) - 1, Number(day));
                                        return d.toLocaleDateString('pt-BR', { weekday: 'short' }).slice(0,3);
                                    })()}
                                    {r.isFolgaVendida && <span className="block text-[8px] text-blue-600 dark:text-blue-400 font-black">FOLGA</span>}
                                </td>
                                <td className="py-1.5 text-gray-700 dark:text-gray-300">{r.realEntry || '-'}</td>
                                <td className="py-1.5 text-gray-400 dark:text-gray-500">{r.punchEntry || '-'}</td>
                                <td className="py-1.5 text-gray-400 dark:text-gray-500">{r.punchExit || '-'}</td>
                                <td className="py-1.5 text-gray-700 dark:text-gray-300">{r.realExit || '-'}</td>
                                <td className="py-1.5 text-right font-bold text-gray-800 dark:text-gray-200">{getDailyHours(r, req.employeeType)}</td>
                            </tr>
                        ))}
                    </tbody>
                </table>

                {/* Detalhamento do Cálculo */}
                <div className="mt-4 pt-4 border-t border-gray-200 dark:border-gray-700">
                    <div className="bg-blue-50/50 dark:bg-blue-900/20 rounded-xl p-3 space-y-1 font-mono text-[10px]">
                        {(() => {
                            const employee = employees.find(e => String(e.id) === String(req.employeeId));
                            const sector = sectors.find(s => String(s.id) === String(req.sectorId));
                            const daysWorked = req.records.filter(r => {
                                const h = getDailyHours(r, req.employeeType);
                                return h !== '-' && h !== '0h 00m';
                            }).length;

                            if (req.employeeType === EmployeeType.REGISTRADO) {
                                const salary = parseCurrency(employee?.salary);
                                const monthlyHours = parseFloat(String(employee?.monthlyHours)) || 220;
                                const hourlyBase = salary / monthlyHours;
                                const overtimeRate = hourlyBase * 1.25;
                                const hoursTotal = req.totalTimeDecimal * overtimeRate;
                                const bonusTotal = daysWorked * 12;

                                return (
                                    <>
                                        <div className="flex justify-between"><span>Base:</span><span>{formatCurrency(salary)} / {monthlyHours}h = {formatCurrency(hourlyBase)}/h</span></div>
                                        <div className="flex justify-between text-blue-600 dark:text-blue-400 font-bold"><span>HE (+25%):</span><span>{formatCurrency(overtimeRate)}/h</span></div>
                                        <div className="flex justify-between pt-1 border-t border-blue-100 dark:border-blue-800"><span>Horas:</span><span>{formatDecimalHours(req.totalTimeDecimal)} × {formatCurrency(overtimeRate)} = {formatCurrency(hoursTotal)}</span></div>
                                    </>
                                );
                            } else {
                                const hourlyRate = parseCurrency(sector?.fixedRate);
                                const hoursTotal = req.totalTimeDecimal * hourlyRate;
                                const bonusTotal = daysWorked * 12;

                                return (
                                    <>
                                        <div className="flex justify-between"><span>Valor Fixo:</span><span>{formatCurrency(hourlyRate)}/h</span></div>
                                        <div className="flex justify-between pt-1 border-t border-blue-100 dark:border-blue-800"><span>Horas:</span><span>{formatDecimalHours(req.totalTimeDecimal)} × {formatCurrency(hourlyRate)} = {formatCurrency(hoursTotal)}</span></div>
                                        <div className="flex justify-between"><span>Vale Transporte:</span><span>{daysWorked} dias × R$ 12,00 = {formatCurrency(bonusTotal)}</span></div>
                                    </>
                                );
                            }
                        })()}
                        <div className="pt-2 border-t border-blue-200 dark:border-blue-700 flex justify-between text-xs font-black text-gray-900 dark:text-white">
                            <span>TOTAL:</span>
                            <span>{formatCurrency(req.calculatedValue)}</span>
                        </div>
                    </div>
                </div>
            </div>
        )}

        <div className="flex gap-1 items-center">
            <button 
                onClick={() => setIsExpanded(!isExpanded)} 
                className="bg-gray-50 dark:bg-gray-700 text-gray-400 dark:text-gray-300 p-2 rounded-lg hover:bg-gray-100 dark:hover:bg-gray-600 transition mr-1"
                title="Ver Detalhes"
            >
                {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
            </button>

          {req.status === RequestStatus.PENDENTE && (
            <>
              <button onClick={() => setRequests(requests.map(r => r.id === req.id ? {...r, status: RequestStatus.APROVADO} : r))} className="flex-1 bg-green-50 text-green-600 p-2 rounded-lg hover:bg-green-600 hover:text-white transition flex justify-center"><CheckCircle className="w-4 h-4" /></button>
              <button onClick={() => setRequests(requests.map(r => r.id === req.id ? {...r, status: RequestStatus.REJEITADO} : r))} className="flex-1 bg-red-50 text-red-600 p-2 rounded-lg hover:bg-red-600 hover:text-white transition flex justify-center"><XCircle className="w-4 h-4" /></button>
            </>
          )}
          <button onClick={() => {
            setEditingRequestId(req.id);
            setModalRecords(JSON.parse(JSON.stringify(req.records)));
            setCurrentWeek(req.weekStarting);
            setEditJustification(req.editJustification || '');
            setShowFormModal(true);
          }} className="flex-1 bg-gray-50 dark:bg-gray-700 text-gray-400 dark:text-gray-300 p-2 rounded-lg hover:bg-gray-200 dark:hover:bg-gray-600 transition flex justify-center"><Edit2 className="w-4 h-4" /></button>
          <button onClick={() => setRequests(requests.map(r => r.id === req.id ? {...r, status: RequestStatus.DELETADO} : r))} className="bg-gray-50 dark:bg-gray-700 text-gray-300 dark:text-gray-400 p-2 rounded-lg hover:bg-red-50 dark:hover:bg-red-900/30 hover:text-red-400 dark:hover:text-red-400 transition flex justify-center"><XCircle className="w-4 h-4" /></button>
        </div>
      </div>
    );
  };

  const EmployeeRow: React.FC<{ e: Employee }> = ({ e }) => {
    const [isExpanded, setIsExpanded] = useState(false);
    const sector = sectors.find(s => s.id === e.sectorId);
    const salary = parseCurrency(e.salary);
    const monthlyHours = parseFloat(String(e.monthlyHours)) || 220;
    const hourlyBase = salary / monthlyHours;
    const overtimeRate = hourlyBase * 1.25;

    return (
      <>
        <tr className="border-b dark:border-gray-700 hover:bg-gray-50 dark:hover:bg-gray-900/50 transition-colors">
          <td className="py-4 dark:text-gray-200">
            <button onClick={() => setIsExpanded(!isExpanded)} className="flex items-center gap-2 hover:text-blue-600 transition-colors font-bold">
              {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
              {e.name}
            </button>
          </td>
          <td className="py-4 dark:text-gray-300">{sector?.name}</td>
          <td className="py-4 dark:text-gray-300 font-bold text-blue-600 dark:text-blue-400">{formatCurrency(overtimeRate)}</td>
          <td className="py-4 text-right flex justify-end gap-2">
            <button onClick={() => { setEditingEmployeeId(e.id); setNewEmpData({ name: e.name, sectorId: e.sectorId, salary: e.salary, monthlyHours: e.monthlyHours, type: e.type }); }} className="text-blue-500 hover:text-blue-400"><Edit2 className="w-4 h-4" /></button>
            <button onClick={() => setEmployees(employees.filter(emp => emp.id !== e.id))} className="text-red-400 hover:text-red-300"><XCircle className="w-4 h-4" /></button>
          </td>
        </tr>
        {isExpanded && (
          <tr className="bg-blue-50/30 dark:bg-blue-900/10">
            <td colSpan={4} className="p-4">
              <div className="bg-white dark:bg-gray-800 rounded-2xl p-4 border border-blue-100 dark:border-blue-800 shadow-sm space-y-3 max-w-md">
                <h5 className="text-[10px] font-black uppercase text-blue-600 dark:text-blue-400 mb-2">Memória de Cálculo (Hora Extra)</h5>
                <div className="space-y-1 text-xs font-mono">
                  {e.type === EmployeeType.REGISTRADO ? (
                    <>
                      <div className="flex justify-between text-gray-600 dark:text-gray-400">
                        <span>Base:</span>
                        <span>{formatCurrency(salary)} / {monthlyHours}h</span>
                      </div>
                      <div className="flex justify-between border-t border-gray-100 dark:border-gray-700 pt-1">
                        <span>Hora Base:</span>
                        <span>{formatCurrency(hourlyBase)}</span>
                      </div>
                      <div className="flex justify-between font-bold text-blue-600 dark:text-blue-400">
                        <span>HE (+25%):</span>
                        <span>{formatCurrency(overtimeRate)}</span>
                      </div>
                    </>
                  ) : (
                    <>
                      <div className="flex justify-between text-gray-600 dark:text-gray-400">
                        <span>Valor Hora Fixo:</span>
                        <span>{formatCurrency(parseCurrency(sector?.fixedRate))}</span>
                      </div>
                      <div className="pt-2 mt-2 border-t border-blue-100 dark:border-blue-800 text-[9px] text-gray-500 dark:text-gray-400 italic">
                        * + R$ 12,00/dia (Vale Transporte) no fechamento.
                      </div>
                    </>
                  )}
                </div>
              </div>
            </td>
          </tr>
        )}
      </>
    );
  };

  const renderAdminRequestsSubView = () => (
    <div className="h-full flex flex-col gap-6">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center bg-white dark:bg-gray-800 p-6 rounded-3xl shadow-sm border border-gray-100 dark:border-gray-700 gap-4">
        <h2 className="text-xl md:text-2xl font-black text-gray-800 dark:text-gray-200">Fluxo de Solicitações</h2>
        <button onClick={() => syncDatabase({ sectors, employees, requests })} className="w-full md:w-auto flex items-center justify-center gap-2 bg-blue-600 text-white px-4 py-3 rounded-xl text-sm font-bold shadow-lg shadow-blue-100 dark:shadow-none hover:bg-blue-700 active:scale-95 transition-transform"><RefreshCw className="w-4 h-4" /> Forçar Sincronização</button>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 flex-1 min-h-0 pb-20 md:pb-0">
        {[
            { title: 'Pendentes', status: RequestStatus.PENDENTE, icon: Clock, color: 'blue' },
            { title: 'Aprovados', status: RequestStatus.APROVADO, icon: CheckCircle, color: 'green' },
            { title: 'Rejeitados', status: RequestStatus.REJEITADO, icon: XCircle, color: 'red' }
        ].map((col) => (
            <div key={col.status} className={`flex flex-col rounded-3xl p-4 border ${col.color === 'blue' ? 'bg-gray-100/50 dark:bg-gray-800/50 border-gray-200/50 dark:border-gray-700/50' : col.color === 'green' ? 'bg-green-50/30 dark:bg-green-900/10 border-green-100/50 dark:border-green-800/30' : 'bg-red-50/30 dark:bg-red-900/10 border-red-100/50 dark:border-red-800/30'}`}>
                <div className="flex items-center justify-between mb-4 px-2">
                    <h3 className={`text-sm font-black uppercase flex items-center gap-2 text-${col.color}-600 dark:text-${col.color}-400`}>
                        <col.icon className="w-4 h-4" /> {col.title}
                    </h3>
                    <span className={`bg-${col.color}-100 dark:bg-${col.color}-900/30 text-${col.color}-700 dark:text-${col.color}-400 text-[10px] px-2 py-0.5 rounded-full font-bold`}>
                        {requests.filter(r => r.status === col.status).length}
                    </span>
                </div>
                <div className="flex-1 overflow-y-auto pr-1">
                    {requests.filter(r => r.status === col.status).map(req => <RequestCard key={req.id} req={req} />)}
                </div>
            </div>
        ))}
      </div>
    </div>
  );

  // --- Renderização Principal ---

  if (state.view === 'EXPIRED') {
    return (
      <div className="flex flex-col items-center justify-center min-h-screen text-center px-4 bg-gray-100 dark:bg-gray-900">
        <div className="bg-white dark:bg-gray-800 p-12 rounded-3xl shadow-2xl max-w-lg w-full border border-transparent dark:border-gray-700">
          <div className="bg-red-100 dark:bg-red-900/30 w-20 h-20 rounded-2xl flex items-center justify-center mx-auto mb-8 text-red-600 dark:text-red-400"><Lock className="w-10 h-10" /></div>
          <h1 className="text-3xl font-bold mb-4 dark:text-white">Acesso Expirado</h1>
          <p className="text-gray-500 dark:text-gray-400 mb-8">Este link expirou. Peça um novo acesso.</p>
          <button onClick={() => window.location.href = window.location.origin + window.location.pathname} className="bg-gray-800 dark:bg-gray-700 text-white px-8 py-3 rounded-xl font-bold hover:bg-gray-700 dark:hover:bg-gray-600 transition">Início</button>
        </div>
      </div>
    );
  }

  // Admin Navigation Items
  const navItems = [
    { id: 'DASHBOARD', label: 'Dash', icon: LayoutDashboard },
    { id: 'SECTORS', label: 'Setores', icon: MapPin },
    { id: 'EMPLOYEES', label: 'Func.', icon: Users },
    { id: 'REQUESTS', label: 'Solicit.', icon: ClipboardList },
    { id: 'INTEGRATIONS', label: 'Sync', icon: Database },
    { id: 'FILES', label: 'Arquivos', icon: Folder },
  ];

  const appsScriptCode = `
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
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Aguarda até 10 segundos por um lock
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
      mergeRequests(ss, "Solicitacoes", body.data.requests, body.data.isAdmin);
      
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
    
    if (body.action === "FECHAMENTO_SEMANAL") {
      executarFechamentoSemanalAPI(body.data.folderRegId, body.data.folderFixoId);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (body.action === "LIST_FILES") {
      const result = listDriveFiles(body.data.folderId);
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: result }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: String(error) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
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

function mergeRequests(ss, sheetName, newRequests, isAdmin) {
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
  const statusIndex = sheetHeaders.indexOf("status");
  
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
      if (isAdmin) {
        const existingRow = existingIds[req.id];
        let existingStatus = "PENDENTE";
        if (statusIndex !== -1 && existingRow <= existingData.length) {
          existingStatus = existingData[existingRow - 1][statusIndex];
        }
        
        // Se a solicitação já foi aprovada/rejeitada no servidor,
        // não deixe um cliente com estado "PENDENTE" sobrescrever isso.
        if (existingStatus !== "PENDENTE" && req.status === "PENDENTE") {
          // Ignora a atualização para não apagar o status aprovado
        } else {
          sheet.getRange(existingRow, 1, 1, sheetHeaders.length).setValues([rowData]);
        }
      }
    } else {
      // Append new
      sheet.appendRow(rowData);
      backupSheet.appendRow(rowData);
      if (idIndex !== -1 && req.id) {
        existingIds[req.id] = sheet.getLastRow();
      }
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
    ui.alert("Erro durante a exportação: " + String(e));
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
      let range = aba.getDataRange();
      let matriz = range.getValues();
      let formulas = range.getFormulas(); // Coleta as fórmulas para não apagá-las
      
      limparMatriz(matriz, nome.includes("REGISTRADO") ? "REGISTRADO" : "FIXO");
      
      restaurarFormulas(matriz, formulas); // Restaura as fórmulas antes de salvar
      aba.getRange(1, 1, matriz.length, matriz[0].length).setValues(matriz);
    }
  });

  ui.alert("Fechamento concluído com sucesso!");
}

function executarFechamentoSemanalAPI(paramFolderRegId, paramFolderFixoId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // 1. Exporta antes de apagar os dados
    exportarFolhasSextaFeira(paramFolderRegId, paramFolderFixoId);
    SpreadsheetApp.flush(); 
  } catch (e) {
    throw new Error("Erro durante a exportação: " + String(e));
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
      let range = aba.getDataRange();
      let matriz = range.getValues();
      let formulas = range.getFormulas(); // Coleta as fórmulas para não apagá-las
      
      limparMatriz(matriz, nome.includes("REGISTRADO") ? "REGISTRADO" : "FIXO");
      
      restaurarFormulas(matriz, formulas); // Restaura as fórmulas antes de salvar
      aba.getRange(1, 1, matriz.length, matriz[0].length).setValues(matriz);
    }
  });
}

function extractFolderId(input) {
  if (!input) return '';
  const match = String(input).match(/[-\w]{15,}/);
  return match ? match[0] : input;
}

function listDriveFiles(folderId) {
  if (!folderId) return { folders: [], files: [] };
  const cleanId = extractFolderId(folderId);
  try {
    const folder = DriveApp.getFolderById(cleanId);
    const folders = [];
    const files = [];
    
    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
      const sub = subFolders.next();
      folders.push({ id: sub.getId(), name: sub.getName() });
    }
    
    const folderFiles = folder.getFiles();
    while (folderFiles.hasNext()) {
      const file = folderFiles.next();
      files.push({ id: file.getId(), name: file.getName(), url: file.getUrl() });
    }
    
    return { folders, files };
  } catch (e) {
    let errorMsg = "Erro ao listar arquivos: " + String(e);
    if (String(e).includes("Access denied") || String(e).includes("No item with the given ID")) {
      errorMsg += "\\n\\n(DICA IMPORTANTE: O Google Apps Script não tem permissão para acessar esta pasta. Se você tem certeza que o ID está correto, o problema é a forma como o script foi implantado.\\n\\nSolução: No Apps Script, vá em 'Implantar' > 'Gerenciar implantações' > Editar (lápis) > Em 'Executar como', mude para 'Usuário que acessa o aplicativo web'. Salve e tente novamente.)";
    }
    throw new Error(errorMsg);
  }
}

/**
 * ============================================================
 * SEÇÃO: EXPORTAÇÃO E GESTÃO DE ARQUIVOS
 * ============================================================
 */

function obterOuCriarSubpasta(pastaPai, nomeSubpasta) {
  const subpastas = pastaPai.getFoldersByName(nomeSubpasta);
  return subpastas.hasNext() ? subpastas.next() : pastaPai.createFolder(nomeSubpasta);
}

function exportarFolhasSextaFeira(paramFolderRegId, paramFolderFixoId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const PASTA_REGISTRADO_ID = extractFolderId(paramFolderRegId || props.getProperty('FOLDER_REG_ID') || "1OGOxVmi2nEwI47HP9l48VdVBKQeJTVqm");
  const PASTA_FIXO_ID = extractFolderId(paramFolderFixoId || props.getProperty('FOLDER_FIXO_ID') || "1RzzDCHznw97QxwDLh_qvf8NE8yKPNdWU");

  const hoje = new Date();
  const ano = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "yyyy");
  
  // Array com nomes dos meses em português
  const meses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
  const mesNome = meses[hoje.getMonth()];
  
  const dataPasta = Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd-MM");

  try {
    // Para HE - REGISTRADO
    const pastaRaizReg = DriveApp.getFolderById(PASTA_REGISTRADO_ID);
    const pastaAnoReg = obterOuCriarSubpasta(pastaRaizReg, ano);
    const pastaMesReg = obterOuCriarSubpasta(pastaAnoReg, mesNome);
    const pReg = obterOuCriarSubpasta(pastaMesReg, dataPasta);
    processarExportacaoIndividual(ss, "HE - REGISTRADO", "REGISTRADO", pReg);
    
    // Para HE - FIXO
    const pastaRaizFixo = DriveApp.getFolderById(PASTA_FIXO_ID);
    const pastaAnoFixo = obterOuCriarSubpasta(pastaRaizFixo, ano);
    const pastaMesFixo = obterOuCriarSubpasta(pastaAnoFixo, mesNome);
    const pFix = obterOuCriarSubpasta(pastaMesFixo, dataPasta);
    processarExportacaoIndividual(ss, "HE - FIXO", "FIXO", pFix);
  } catch(e) { 
    console.error("Erro na exportação: " + e);
    let errorMsg = "Erro ao acessar pastas do Drive ou exportar arquivos: " + String(e);
    if (String(e).includes("Access denied") || String(e).includes("No item with the given ID")) {
      errorMsg += "\\n\\n(DICA IMPORTANTE: O Google Apps Script não tem permissão para acessar esta pasta. Se você tem certeza que o ID está correto, o problema é a forma como o script foi implantado.\\n\\nSolução: No Apps Script, vá em 'Implantar' > 'Gerenciar implantações' > Editar (lápis) > Em 'Executar como', mude para 'Usuário que acessa o aplicativo web'. Salve e tente novamente.)";
    }
    throw new Error(errorMsg);
  }
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
        nomeArquivo = \`\${nomeSetor}-\${dataCurta}\`;
      } else {
        contadorNomes[nomeSetor]++;
        nomeArquivo = \`\${nomeSetor} (PARTE \${contadorNomes[nomeSetor]})-\${dataCurta}\`;
      }

    } else {
      // REGISTRADO: Extrai apenas o primeiro nome (Ex: MIKAELA & VALDIRENE)
      let raw1 = (dados[i+4] && dados[i+4][1]) ? dados[i+4][1].toString().trim() : "";
      let raw2 = (dados[i+28] && dados[i+28][1]) ? dados[i+28][1].toString().trim() : "";
      
      if (raw1 === "" && raw2 === "") continue;

      // Extrai o nome do setor para criar a subpasta (Linha 2 do bloco)
      let nomeSetor = (dados[i+1] && dados[i+1][1]) ? dados[i+1][1].toString().toUpperCase().trim() : "GERAL";
      let pastaSetor = obterOuCriarSubpasta(pastaDestino, nomeSetor);

      let pNome1 = raw1 !== "" ? raw1.split(" ")[0].toUpperCase() : "VAGO";
      let pNome2 = raw2 !== "" ? raw2.split(" ")[0].toUpperCase() : "VAGO";
      nomeArquivo = \`\${pNome1} & \${pNome2}\`;

      if (!contadorNomes[nomeArquivo]) {
        contadorNomes[nomeArquivo] = 1;
      } else {
        contadorNomes[nomeArquivo]++;
        nomeArquivo += \` (\${contadorNomes[nomeArquivo]})\`;
      }
      let numCols = 11;
      let rangeFolha = abaOrigem.getRange(i + 1, 1, saltoLinhas, numCols); 
      gerarNovoArquivoSheets(nomeArquivo, rangeFolha, pastaSetor);
      continue;
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
  arquivo.moveTo(pastaDestino);
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
    
    const regEmployees = agrupados.filter(r => r.employeeType.toUpperCase().trim() === "REGISTRADO");
    
    // Agrupar por setor para garantir que 2 do mesmo setor fiquem na mesma folha
    const sectorsMap = {};
    regEmployees.forEach(emp => {
      const sName = (emp.sectorName || "GERAL").toUpperCase().trim();
      if (!sectorsMap[sName]) sectorsMap[sName] = [];
      sectorsMap[sName].push(emp);
    });

    let currentSheetIdx = 0; 
    Object.keys(sectorsMap).sort().forEach(sName => {
      const emps = sectorsMap[sName];
      for (let i = 0; i < emps.length; i += 2) {
        // Encontra a próxima folha sulfite vazia (bloco de 52 linhas)
        while (currentSheetIdx < matriz.length && (matriz[currentSheetIdx + 4] && (matriz[currentSheetIdx + 4][1] || "").toString().trim() !== "")) {
           currentSheetIdx += 52;
        }
        if (currentSheetIdx >= matriz.length) break;

        // Preenche o primeiro funcionário (Slot 1)
        const emp1 = emps[i];
        matriz[currentSheetIdx + 4][1] = emp1.employeeName;
        matriz[currentSheetIdx + 1][1] = emp1.sectorName; // Nome do setor na linha 2 do bloco
        preencherColunaAERegistros(matriz, formulas, currentSheetIdx + 4 + 7, emp1.records);

        // Preenche o segundo funcionário (Slot 2) se houver outro do mesmo setor
        if (i + 1 < emps.length) {
          const emp2 = emps[i + 1];
          matriz[currentSheetIdx + 28][1] = emp2.employeeName;
          preencherColunaAERegistros(matriz, formulas, currentSheetIdx + 28 + 7, emp2.records);
        }
        
        currentSheetIdx += 52;
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
        // Formata a data de YYYY-MM-DD para DD/MM/YYYY
        let partes = diasComDados[i].date.split("-");
        let dataFormatada = partes[2] + "/" + partes[1] + "/" + partes[0];
        
        matriz[r][col] = dataFormatada;
        matriz[r][col + 1] = diasComDados[i].realEntry || "";
        matriz[r][col + 2] = diasComDados[i].realExit || "";  
        let lE = (col === 0) ? "B" : "I"; let lS = (col === 0) ? "C" : "J";
        formulas[r][col + 4] = \`=\${lS}\${r+1}-\${lE}\${r+1}\`;
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
  
  const diasDaSemana = ["SEGUNDA-FEIRA", "TERÇA-FEIRA", "QUARTA-FEIRA", "QUINTA-FEIRA", "SEXTA-FEIRA", "SÁBADO", "DOMINGO"];
  
  for (let i = 0; i < 7; i++) {
    let r = linhaInicio + i;
    if (!matriz[r]) continue;
    let dataLoop = new Date(dataRef);
    dataLoop.setDate(dataRef.getDate() + i);
    
    // Preenche a coluna 0 com o dia da semana
    matriz[r][0] = diasDaSemana[i];
    
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

`

  return (
    <div className="min-h-screen bg-gray-50 dark:bg-gray-900 text-gray-900 dark:text-gray-100 transition-colors duration-200">
      <button 
        onClick={() => setIsDarkMode(!isDarkMode)} 
        className="fixed top-4 right-4 p-3 rounded-full bg-white dark:bg-gray-800 text-gray-800 dark:text-gray-200 shadow-lg hover:bg-gray-100 dark:hover:bg-gray-700 transition-colors z-50 border border-gray-200 dark:border-gray-700"
        title="Alternar Tema"
      >
        {isDarkMode ? <Sun className="w-5 h-5" /> : <Moon className="w-5 h-5" />}
      </button>

      {state.view === 'HOME' && (
        <div className="flex flex-col items-center justify-center min-h-[80vh] text-center px-4">
          <div className="bg-white dark:bg-gray-800 p-8 md:p-12 rounded-3xl shadow-2xl max-w-lg w-full transform transition hover:scale-105 duration-300 border border-transparent dark:border-gray-700">
            <div className="bg-blue-600 w-16 h-16 md:w-20 md:h-20 rounded-2xl flex items-center justify-center mx-auto mb-8 shadow-lg shadow-blue-200 dark:shadow-none"><ClipboardList className="text-white w-8 h-8 md:w-10 md:h-10" /></div>
            <h1 className="text-3xl md:text-4xl font-extrabold text-gray-800 dark:text-white mb-4">Controle de Horas</h1>
            <p className="text-gray-500 dark:text-gray-400 mb-10 text-base md:text-lg">Gerenciamento eficiente de jornadas semanais.</p>
            <button onClick={() => setState(prev => ({ ...prev, view: 'SELECTION' }))} className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-4 px-8 rounded-xl transition shadow-xl shadow-blue-100 dark:shadow-none flex items-center justify-center gap-3 group text-lg">
              Começar <Play className="w-5 h-5 group-hover:translate-x-1 transition" />
            </button>
            <div className="mt-8 pt-8 border-t border-gray-100 dark:border-gray-700 relative">
              <input type="password" placeholder="Acesso Admin" className="w-full pl-12 pr-4 py-4 bg-gray-50 dark:bg-gray-900 border border-gray-200 dark:border-gray-700 rounded-xl outline-none text-black dark:text-white focus:bg-white dark:focus:bg-gray-800 focus:border-blue-500 transition text-base" value={adminPassword} onChange={(e) => setAdminPassword(e.target.value)} onKeyPress={(e) => e.key === 'Enter' && handleAdminLogin()} />
              <Settings className="absolute left-4 top-[70%] -translate-y-1/2 text-gray-400 dark:text-gray-500 w-5 h-5" />
            </div>
          </div>
        </div>
      )}

      {state.view === 'SELECTION' && (
        <div className="flex flex-col items-center justify-center min-h-[80vh] px-4">
          <button onClick={() => setState(prev => ({ ...prev, view: 'HOME' }))} className="mb-8 flex items-center gap-2 text-gray-500 dark:text-gray-400 hover:text-blue-600 dark:hover:text-blue-400 transition font-medium p-2"><ArrowLeft className="w-5 h-5" /> Voltar</button>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6 md:gap-8 max-w-4xl w-full">
            <button onClick={() => setState(prev => ({ ...prev, view: 'FLOW', flowType: EmployeeType.REGISTRADO }))} className="bg-white dark:bg-gray-800 p-8 md:p-10 rounded-3xl shadow-xl hover:border-blue-500 dark:hover:border-blue-400 border-2 border-transparent dark:border-gray-700 transition-all group text-left flex flex-row md:flex-col items-center md:items-start gap-6 md:gap-0">
              <div className="bg-blue-50 dark:bg-blue-900/30 w-14 h-14 md:w-16 md:h-16 rounded-2xl flex items-center justify-center md:mb-6 group-hover:bg-blue-600 transition shrink-0"><Users className="text-blue-600 dark:text-blue-400 w-7 h-7 md:w-8 md:h-8 group-hover:text-white transition" /></div>
              <h2 className="text-xl md:text-2xl font-bold dark:text-white">Registrado</h2>
            </button>
            <button onClick={() => setState(prev => ({ ...prev, view: 'FLOW', flowType: EmployeeType.FIXO }))} className="bg-white dark:bg-gray-800 p-8 md:p-10 rounded-3xl shadow-xl hover:border-green-500 dark:hover:border-green-400 border-2 border-transparent dark:border-gray-700 transition-all group text-left flex flex-row md:flex-col items-center md:items-start gap-6 md:gap-0">
              <div className="bg-green-50 dark:bg-green-900/30 w-14 h-14 md:w-16 md:h-16 rounded-2xl flex items-center justify-center md:mb-6 group-hover:bg-green-600 transition shrink-0"><MapPin className="text-green-600 dark:text-green-400 w-7 h-7 md:w-8 md:h-8 group-hover:text-white transition" /></div>
              <h2 className="text-xl md:text-2xl font-bold dark:text-white">Fixo</h2>
            </button>
          </div>
        </div>
      )}

      {state.view === 'FLOW' && (
        <div className="max-w-xl mx-auto py-8 px-4">
          <button onClick={() => setState(prev => ({ ...prev, view: 'SELECTION' }))} className="mb-6 flex items-center gap-2 text-gray-500 dark:text-gray-400 hover:text-blue-600 dark:hover:text-blue-400 transition font-medium p-2"><ArrowLeft className="w-5 h-5" /> Voltar</button>
          <div className="bg-white dark:bg-gray-800 rounded-3xl shadow-xl p-6 md:p-8 space-y-6 md:space-y-8 border border-transparent dark:border-gray-700">
            <h2 className="text-2xl font-bold text-gray-800 dark:text-white border-b dark:border-gray-700 pb-4">{state.flowType}</h2>
            <div>
              <label className="block text-sm font-semibold mb-2 dark:text-gray-300">Setor</label>
              <select className="w-full p-4 border dark:border-gray-600 rounded-xl bg-gray-50 dark:bg-gray-900 text-black dark:text-white text-base focus:bg-white dark:focus:bg-gray-800 focus:border-blue-500 transition" value={selectedSector} onChange={(e) => setSelectedSector(e.target.value)}>
                <option value="">Selecione...</option>
                {sectors.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </div>
            {selectedSector && (
              <div>
                <label className="block text-sm font-semibold mb-2 dark:text-gray-300">Funcionário</label>
                {state.flowType === EmployeeType.REGISTRADO ? (
                  <select className="w-full p-4 border dark:border-gray-600 rounded-xl bg-gray-50 dark:bg-gray-900 text-black dark:text-white text-base focus:bg-white dark:focus:bg-gray-800 focus:border-blue-500 transition" value={selectedEmployee} onChange={(e) => setSelectedEmployee(e.target.value)}>
                    <option value="">Selecione...</option>
                    {employees.filter(e => String(e.sectorId) === String(selectedSector) && e.type === state.flowType).sort((a, b) => a.name.localeCompare(b.name)).map(e => <option key={e.id} value={e.id}>{e.name}</option>)}
                  </select>
                ) : (
                  <input type="text" placeholder="Nome" className="w-full p-4 border dark:border-gray-600 rounded-xl bg-gray-50 dark:bg-gray-900 text-black dark:text-white text-base focus:bg-white dark:focus:bg-gray-800 focus:border-blue-500 transition" value={selectedEmployee} onChange={(e) => setSelectedEmployee(e.target.value)} />
                )}
              </div>
            )}
            <button disabled={!selectedSector || !selectedEmployee} onClick={() => setShowFormModal(true)} className="w-full bg-blue-600 disabled:bg-gray-300 dark:disabled:bg-gray-700 dark:disabled:text-gray-500 text-white font-bold py-4 rounded-xl shadow-lg active:scale-95 transition text-lg">Lançar Horários</button>
          </div>
        </div>
      )}

      {state.view === 'SUCCESS' && (
        <div className="flex flex-col items-center justify-center min-h-[80vh] text-center px-4">
          <div className="bg-white dark:bg-gray-800 p-8 md:p-12 rounded-3xl shadow-2xl max-w-lg w-full transform transition hover:scale-105 duration-300 border border-transparent dark:border-gray-700">
            <div className="bg-green-100 dark:bg-green-900/30 w-20 h-20 rounded-full flex items-center justify-center mx-auto mb-8 shadow-lg shadow-green-100 dark:shadow-none">
              <CheckCircle className="text-green-600 dark:text-green-400 w-10 h-10" />
            </div>
            <h1 className="text-3xl md:text-4xl font-extrabold text-gray-800 dark:text-white mb-4">Ok, registrado!</h1>
            <p className="text-gray-500 dark:text-gray-400 mb-10 text-base md:text-lg">Muito obrigado pelo preenchimento.</p>
            <button 
              onClick={() => setState(prev => ({ ...prev, view: 'HOME' }))} 
              className="w-full bg-gray-800 dark:bg-gray-700 hover:bg-gray-900 dark:hover:bg-gray-600 text-white font-bold py-4 px-8 rounded-xl transition shadow-xl flex items-center justify-center gap-3 text-lg"
            >
              Sair
            </button>
          </div>
        </div>
      )}

      {state.view === 'ADMIN' && (
        <div className="flex flex-col md:flex-row h-screen bg-gray-50 dark:bg-gray-900 overflow-hidden">
          {/* Desktop Sidebar */}
          <div className="hidden md:flex w-72 bg-white dark:bg-gray-800 border-r border-gray-100 dark:border-gray-700 flex-col p-6 shadow-sm z-20">
            <div className="flex items-center gap-3 mb-12"><div className="bg-blue-600 p-2 rounded-lg text-white"><Settings className="w-5 h-5" /></div><h2 className="text-xl font-black dark:text-white">Admin</h2></div>
            <nav className="flex-1 space-y-2">
              {navItems.map(item => (
                <button key={item.id} onClick={() => setState(prev => ({ ...prev, adminSubView: item.id as any }))} className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl font-bold transition-all ${state.adminSubView === item.id ? 'bg-blue-600 text-white shadow-lg' : 'text-gray-500 dark:text-gray-400 hover:bg-gray-50 dark:hover:bg-gray-700'}`}>
                  <item.icon className="w-5 h-5" />{item.label}
                </button>
              ))}
            </nav>
            <button onClick={() => { setIsAuth(false); setState(prev => ({ ...prev, view: 'HOME' })) }} className="flex items-center gap-3 px-4 py-3 text-red-500 font-bold hover:bg-red-50 dark:hover:bg-red-900/20 rounded-xl transition mt-auto"><LogOut className="w-5 h-5" /> Sair</button>
          </div>

          {/* Mobile Bottom Nav */}
          <div className="md:hidden fixed bottom-0 left-0 right-0 bg-white dark:bg-gray-800 border-t border-gray-200 dark:border-gray-700 flex justify-around p-2 z-50 safe-area-bottom">
            {navItems.map(item => (
                <button key={item.id} onClick={() => setState(prev => ({ ...prev, adminSubView: item.id as any }))} className={`flex flex-col items-center justify-center p-2 rounded-xl transition-all ${state.adminSubView === item.id ? 'text-blue-600 dark:text-blue-400' : 'text-gray-400 dark:text-gray-500'}`}>
                    <item.icon className={`w-6 h-6 mb-1 ${state.adminSubView === item.id ? 'fill-current' : ''}`} />
                    <span className="text-[10px] font-bold">{item.label}</span>
                </button>
            ))}
            <button onClick={() => { setIsAuth(false); setState(prev => ({ ...prev, view: 'HOME' })) }} className="flex flex-col items-center justify-center p-2 text-red-400 dark:text-red-500">
                <LogOut className="w-6 h-6 mb-1" />
                <span className="text-[10px] font-bold">Sair</span>
            </button>
          </div>

          {/* Main Content */}
          <div className="flex-1 overflow-y-auto p-4 md:p-10 relative pb-24 md:pb-10 bg-gray-50 dark:bg-gray-900">
            {isSyncing && <div className="absolute top-4 right-4 md:top-10 md:right-10 flex items-center gap-2 text-blue-600 dark:text-blue-400 font-bold text-xs md:text-sm bg-blue-50 dark:bg-blue-900/30 px-3 py-1 md:px-4 md:py-2 rounded-full border border-blue-100 dark:border-blue-800 z-10"><RefreshCw className="w-3 h-3 md:w-4 md:h-4 animate-spin" /> Atualizando...</div>}
            
            {state.adminSubView === 'DASHBOARD' && (
              <div className="space-y-6 md:space-y-8">
                {/* Header do Dashboard com Ações */}
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <div>
                    <h2 className="text-2xl font-black text-gray-800 dark:text-white">Dashboard</h2>
                    <p className="text-sm text-gray-500 dark:text-gray-400">Visão geral do sistema e exportação de dados.</p>
                  </div>
                  <div className="flex items-center gap-3">
                    <button 
                      onClick={exportToExcel}
                      className="flex items-center gap-2 bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded-xl font-bold transition-all shadow-sm active:scale-95"
                    >
                      <FileText className="w-4 h-4" />
                      Exportar Excel
                    </button>
                  </div>
                </div>

                {/* Cards de Resumo */}
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 md:gap-6">
                  {[
                      { label: 'Total Gasto', val: formatCurrency(dashboardData.totalSpent), icon: DollarSign, color: 'text-gray-800 dark:text-gray-100' },
                      { label: 'Aprovadas', val: dashboardData.approvedCount, icon: CheckCircle, color: 'text-green-600 dark:text-green-400' },
                      { label: 'Pendentes', val: requests.filter(r => r.status === RequestStatus.PENDENTE).length, icon: Clock, color: 'text-blue-600 dark:text-blue-400' },
                      { label: 'Ticket Médio', val: dashboardData.approvedCount > 0 ? formatCurrency(dashboardData.totalSpent / dashboardData.approvedCount) : 'R$ 0,00', icon: TrendingUp, color: 'text-purple-600 dark:text-purple-400' }
                  ].map((stat, i) => (
                      <div key={i} className="bg-white dark:bg-gray-800 p-5 rounded-2xl shadow-sm border border-gray-100 dark:border-gray-700">
                        <div className="flex items-center gap-2 text-gray-400 dark:text-gray-500 mb-2">
                            <stat.icon className="w-4 h-4" />
                            <span className="text-xs font-bold uppercase">{stat.label}</span>
                        </div>
                        <h3 className={`text-2xl font-black ${stat.color}`}>{stat.val}</h3>
                      </div>
                  ))}
                </div>

                {/* Gráficos */}
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                    <div className="bg-white dark:bg-gray-800 p-6 md:p-8 rounded-3xl border border-gray-100 dark:border-gray-700 shadow-sm flex flex-col min-h-[400px]">
                        <h3 className="text-lg font-bold text-gray-800 dark:text-white mb-6 flex items-center gap-2">
                            <MapPin className="w-5 h-5 text-blue-600 dark:text-blue-400" />
                            Gastos por Setor
                        </h3>
                        <div className="h-[300px] w-full">
                            {dashboardData.expensesBySector.length > 0 ? (
                                <ResponsiveContainer width="99%" height={300} minWidth={0} minHeight={0}>
                                    <BarChart data={dashboardData.expensesBySector} layout="vertical" margin={{ top: 5, right: 30, left: 10, bottom: 5 }}>
                                        <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke={isDarkMode ? '#374151' : '#e5e7eb'} />
                                        <XAxis type="number" hide />
                                        <YAxis dataKey="name" type="category" width={80} tick={{fontSize: 10, fill: isDarkMode ? '#9ca3af' : '#6b7280'}} />
                                        <Tooltip cursor={{fill: isDarkMode ? '#374151' : '#f3f4f6'}} contentStyle={{ backgroundColor: isDarkMode ? '#1f2937' : '#fff', borderColor: isDarkMode ? '#374151' : '#e5e7eb', color: isDarkMode ? '#f3f4f6' : '#111827' }} />
                                        <Bar dataKey="value" fill="#3b82f6" radius={[0, 4, 4, 0]} barSize={24}>
                                            {dashboardData.expensesBySector.map((entry, index) => (
                                                <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                                            ))}
                                        </Bar>
                                    </BarChart>
                                </ResponsiveContainer>
                            ) : (
                                <div className="h-full flex flex-col items-center justify-center text-gray-300 dark:text-gray-600"><AlertCircle className="w-10 h-10 mb-2" /><p className="text-sm">Sem dados</p></div>
                            )}
                        </div>
                    </div>

                    <div className="bg-white dark:bg-gray-800 p-6 md:p-8 rounded-3xl border border-gray-100 dark:border-gray-700 shadow-sm flex flex-col min-h-[400px]">
                        <h3 className="text-lg font-bold text-gray-800 dark:text-white mb-6 flex items-center gap-2">
                            <Users className="w-5 h-5 text-green-600 dark:text-green-400" />
                            Registrado vs Fixo
                        </h3>
                        <div className="h-[300px] w-full">
                            {dashboardData.expensesByType.length > 0 ? (
                                <ResponsiveContainer width="99%" height={300} minWidth={0} minHeight={0}>
                                    <PieChart>
                                        <Pie data={dashboardData.expensesByType} cx="50%" cy="50%" innerRadius={60} outerRadius={100} paddingAngle={5} dataKey="value">
                                            {dashboardData.expensesByType.map((entry, index) => <Cell key={`cell-${index}`} fill={entry.color} />)}
                                        </Pie>
                                        <Tooltip contentStyle={{ backgroundColor: isDarkMode ? '#1f2937' : '#fff', borderColor: isDarkMode ? '#374151' : '#e5e7eb', color: isDarkMode ? '#f3f4f6' : '#111827' }} />
                                        <Legend verticalAlign="bottom" height={36} iconType="circle" wrapperStyle={{ color: isDarkMode ? '#f3f4f6' : '#111827' }} />
                                    </PieChart>
                                </ResponsiveContainer>
                            ) : (
                                <div className="h-full flex flex-col items-center justify-center text-gray-300 dark:text-gray-600"><AlertCircle className="w-10 h-10 mb-2" /><p className="text-sm">Sem dados</p></div>
                            )}
                        </div>
                    </div>
                </div>
              </div>
            )}

            {state.adminSubView === 'REQUESTS' && renderAdminRequestsSubView()}

            {state.adminSubView === 'SECTORS' && (
              <div className="bg-white dark:bg-gray-800 p-6 md:p-8 rounded-3xl border border-gray-100 dark:border-gray-700 space-y-8">
                <h2 className="text-2xl font-bold dark:text-white">Setores</h2>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4 bg-gray-50 dark:bg-gray-900 p-6 rounded-2xl">
                  <input type="text" placeholder="Nome" className="p-4 border dark:border-gray-700 rounded-xl bg-white dark:bg-gray-800 text-black dark:text-white text-base focus:border-blue-500 outline-none" value={newSec.name} onChange={(e) => setNewSec({ ...newSec, name: e.target.value })} />
                  <input type="number" placeholder="Valor Hora" className="p-4 border dark:border-gray-700 rounded-xl bg-white dark:bg-gray-800 text-black dark:text-white text-base focus:border-blue-500 outline-none" value={newSec.fixedRate || ''} onChange={(e) => setNewSec({ ...newSec, fixedRate: parseFloat(e.target.value) })} />
                  <button onClick={() => { if(newSec.name) { setSectors([...sectors, {...newSec, id: Math.random().toString(36).substr(2, 9)}]); setNewSec({name: '', fixedRate: 0}); } }} className="bg-blue-600 text-white font-bold rounded-xl py-3 active:scale-95 transition">Adicionar</button>
                </div>
                {/* Responsive List: Card on Mobile, Table on Desktop */}
                <div className="hidden md:block">
                    <table className="w-full text-left"><thead><tr className="text-gray-400 dark:text-gray-500 text-xs border-b dark:border-gray-700"><th className="py-4">Setor</th><th className="py-4">Valor Hora</th><th className="py-4 text-right">Ação</th></tr></thead><tbody>{sectors.map(s => (<tr key={s.id} className="border-b dark:border-gray-700"><td className="py-4 font-semibold dark:text-gray-200">{s.name}</td><td className="py-4 dark:text-gray-300">{formatCurrency(s.fixedRate)}</td><td className="py-4 text-right"><button onClick={() => setSectors(sectors.filter(sec => sec.id !== s.id))} className="text-red-500 hover:text-red-400"><XCircle className="w-5 h-4" /></button></td></tr>))}</tbody></table>
                </div>
                <div className="md:hidden space-y-3">
                    {sectors.map(s => (
                        <div key={s.id} className="bg-gray-50 dark:bg-gray-900 p-4 rounded-xl flex justify-between items-center border border-gray-100 dark:border-gray-700">
                            <div>
                                <h4 className="font-bold text-gray-800 dark:text-gray-200">{s.name}</h4>
                                <p className="text-sm text-gray-500 dark:text-gray-400">{formatCurrency(s.fixedRate)} / hora</p>
                            </div>
                            <button onClick={() => setSectors(sectors.filter(sec => sec.id !== s.id))} className="text-red-500 hover:text-red-400 p-2"><XCircle className="w-6 h-6" /></button>
                        </div>
                    ))}
                </div>
              </div>
            )}

            {state.adminSubView === 'EMPLOYEES' && (
              <div className="bg-white dark:bg-gray-800 p-6 md:p-8 rounded-3xl border border-gray-100 dark:border-gray-700 space-y-8">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <h2 className="text-2xl font-bold dark:text-white">Funcionários</h2>
                  <select 
                    className="p-3 border dark:border-gray-700 rounded-xl bg-gray-50 dark:bg-gray-900 text-black dark:text-white text-sm focus:border-blue-500 outline-none w-full md:w-auto"
                    value={employeeSectorFilter}
                    onChange={(e) => setEmployeeSectorFilter(e.target.value)}
                  >
                    <option value="ALL">Todos os Setores</option>
                    {sectors.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-5 gap-4 bg-gray-50 dark:bg-gray-900 p-6 rounded-2xl">
                  <input type="text" placeholder="Nome" className="p-4 border dark:border-gray-700 rounded-xl bg-white dark:bg-gray-800 text-black dark:text-white text-base focus:border-blue-500 outline-none" value={newEmpData.name} onChange={(e) => setNewEmpData({ ...newEmpData, name: e.target.value })} />
                  <select className="p-4 border dark:border-gray-700 rounded-xl bg-white dark:bg-gray-800 text-black dark:text-white text-base focus:border-blue-500 outline-none" value={newEmpData.sectorId} onChange={(e) => setNewEmpData({ ...newEmpData, sectorId: e.target.value })}><option value="">Setor...</option>{sectors.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}</select>
                  <input type="number" placeholder="Salário" className="p-4 border dark:border-gray-700 rounded-xl bg-white dark:bg-gray-800 text-black dark:text-white text-base focus:border-blue-500 outline-none" value={newEmpData.salary || ''} onChange={(e) => setNewEmpData({ ...newEmpData, salary: parseFloat(e.target.value) })} />
                  <input type="number" placeholder="Horas" className="p-4 border dark:border-gray-700 rounded-xl bg-white dark:bg-gray-800 text-black dark:text-white text-base focus:border-blue-500 outline-none" value={newEmpData.monthlyHours || ''} onChange={(e) => setNewEmpData({ ...newEmpData, monthlyHours: parseFloat(e.target.value) })} />
                  <div className="flex gap-2">
                    <button 
                      onClick={() => { 
                        if(newEmpData.name) { 
                          if (editingEmployeeId) {
                            setEmployees(employees.map(e => e.id === editingEmployeeId ? { ...e, ...newEmpData } : e));
                            setEditingEmployeeId(null);
                          } else {
                            setEmployees([...employees, {...newEmpData, id: Math.random().toString(36).substr(2, 9)}]); 
                          }
                          setNewEmpData({name: '', sectorId: '', salary: 0, monthlyHours: 220, type: EmployeeType.REGISTRADO}); 
                        } 
                      }} 
                      className={`${editingEmployeeId ? 'bg-green-600' : 'bg-blue-600'} text-white font-bold rounded-xl flex-1 py-3 active:scale-95 transition`}
                    >
                      {editingEmployeeId ? 'Salvar' : 'Add'}
                    </button>
                    {editingEmployeeId && (
                      <button onClick={() => { setEditingEmployeeId(null); setNewEmpData({name: '', sectorId: '', salary: 0, monthlyHours: 220, type: EmployeeType.REGISTRADO}); }} className="bg-gray-200 dark:bg-gray-700 text-gray-600 dark:text-gray-300 font-bold rounded-xl px-3 hover:bg-gray-300 dark:hover:bg-gray-600 transition"><XCircle className="w-5 h-5" /></button>
                    )}
                  </div>
                </div>
                
                {/* Desktop View */}
                <div className="hidden md:block">
                    <table className="w-full text-left">
                      <thead>
                        <tr className="border-b dark:border-gray-700 text-xs text-gray-400 dark:text-gray-500">
                          <th className="py-4">Nome</th>
                          <th className="py-4">Setor</th>
                          <th className="py-4">Valor Hora HE</th>
                          <th className="py-4 text-right">Ação</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredAndSortedEmployees.map(e => <EmployeeRow key={e.id} e={e} />)}
                      </tbody>
                    </table>
                </div>

                {/* Mobile View */}
                <div className="md:hidden space-y-4">
                    {filteredAndSortedEmployees.map(e => {
                        const sector = sectors.find(s => s.id === e.sectorId);
                        const salary = parseCurrency(e.salary);
                        const monthlyHours = parseFloat(String(e.monthlyHours)) || 220;
                        const hourlyBase = salary / monthlyHours;
                        const overtimeRate = hourlyBase * 1.25;

                        return (
                            <div key={e.id} className="bg-gray-50 dark:bg-gray-900 p-4 rounded-xl border border-gray-100 dark:border-gray-700 flex flex-col gap-2">
                                <div className="flex justify-between items-start">
                                    <div>
                                        <h4 className="font-bold text-gray-800 dark:text-gray-200 text-lg">{e.name}</h4>
                                        <p className="text-xs text-gray-500 dark:text-gray-400 uppercase font-bold">{sector?.name}</p>
                                    </div>
                                    <div className="flex gap-2">
                                        <button onClick={() => { setEditingEmployeeId(e.id); setNewEmpData({ name: e.name, sectorId: e.sectorId, salary: e.salary, monthlyHours: e.monthlyHours, type: e.type }); }} className="bg-white dark:bg-gray-800 p-2 rounded-lg text-blue-600 dark:text-blue-400 border border-gray-200 dark:border-gray-700 hover:bg-gray-50 dark:hover:bg-gray-700"><Edit2 className="w-5 h-5" /></button>
                                        <button onClick={() => setEmployees(employees.filter(emp => emp.id !== e.id))} className="bg-white dark:bg-gray-800 p-2 rounded-lg text-red-500 dark:text-red-400 border border-gray-200 dark:border-gray-700 hover:bg-gray-50 dark:hover:bg-gray-700"><XCircle className="w-5 h-5" /></button>
                                    </div>
                                </div>
                                <div className="flex items-center gap-2 mt-1">
                                    <span className="text-sm text-gray-600 dark:text-gray-300 bg-white dark:bg-gray-800 px-2 py-1 rounded border dark:border-gray-700">{formatCurrency(salary)}</span>
                                    <span className="text-sm text-gray-600 dark:text-gray-300 bg-white dark:bg-gray-800 px-2 py-1 rounded border dark:border-gray-700">{monthlyHours}h</span>
                                    <span className="text-sm font-bold text-green-600 dark:text-green-400 ml-auto">{formatCurrency(overtimeRate)}/h HE</span>
                                </div>
                                <div className="mt-2 pt-2 border-t border-gray-200 dark:border-gray-700 text-[10px] text-gray-500 dark:text-gray-400 font-mono">
                                    Cálculo: ({formatCurrency(salary)} / {monthlyHours}h) * 1.25 = {formatCurrency(overtimeRate)}
                                </div>
                            </div>
                        );
                    })}
                </div>
              </div>
            )}

            {state.adminSubView === 'INTEGRATIONS' && (
              <div className="bg-white dark:bg-gray-800 p-6 md:p-8 rounded-3xl border border-gray-100 dark:border-gray-700 space-y-8">
                <h2 className="text-2xl font-bold dark:text-white flex items-center gap-2"><Database className="text-blue-600 dark:text-blue-400" /> Sincronização</h2>
                
                {/* Google Service Account Section */}
                {!isServiceAccountSetup ? (
                  <div className="p-6 md:p-8 bg-red-50 dark:bg-red-900/20 rounded-3xl border border-red-200 dark:border-red-800 space-y-4">
                    <div className="flex items-center gap-3">
                      <div className="bg-white dark:bg-gray-800 p-2 rounded-lg shadow-sm">
                        <AlertCircle className="w-6 h-6 text-red-600" />
                      </div>
                      <div>
                        <h3 className="text-xl font-bold text-red-800 dark:text-red-400">Configuração Necessária</h3>
                        <p className="text-sm text-red-600 dark:text-red-300">O robô de sincronização ainda não foi configurado.</p>
                      </div>
                    </div>
                    
                    <div className="bg-white dark:bg-gray-800 p-4 rounded-xl text-sm text-gray-700 dark:text-gray-300 space-y-3">
                      <p>Para que o aplicativo possa ler e escrever na sua planilha automaticamente, você precisa configurar as variáveis de ambiente no AI Studio:</p>
                      <ol className="list-decimal list-inside space-y-2 ml-2">
                        <li>Abra o menu de <strong>Settings</strong> (ícone de engrenagem) no AI Studio.</li>
                        <li>Vá até a seção <strong>Environment Variables</strong>.</li>
                        <li>Adicione a variável <code>GOOGLE_SERVICE_ACCOUNT_EMAIL</code> com o e-mail do seu robô.</li>
                        <li>Adicione a variável <code>GOOGLE_PRIVATE_KEY</code> com a chave privada completa (incluindo as tags BEGIN e END).</li>
                      </ol>
                      <p className="text-xs text-gray-500 mt-4">Após salvar as variáveis, o servidor será reiniciado e esta mensagem desaparecerá.</p>
                    </div>
                  </div>
                ) : (
                  <div className="p-6 md:p-8 bg-blue-50 dark:bg-blue-900/20 rounded-3xl border border-blue-100 dark:border-blue-800 space-y-4">
                    <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
                      <div className="flex items-center gap-3">
                        <div className="bg-white dark:bg-gray-800 p-2 rounded-lg shadow-sm">
                          <Share2 className="w-6 h-6 text-blue-600" />
                        </div>
                        <div>
                          <h3 className="text-xl font-bold text-gray-800 dark:text-white">Conexão via Conta de Serviço (API)</h3>
                          <p className="text-sm text-gray-500 dark:text-gray-400">A comunicação com a planilha é feita automaticamente pelo servidor.</p>
                        </div>
                      </div>
                      <span className="flex items-center gap-1 text-sm font-bold text-green-600 bg-green-100 dark:bg-green-900/40 px-3 py-1 rounded-full">
                        <CheckCircle className="w-4 h-4" /> Ativo
                      </span>
                    </div>
                    
                    <div className="flex flex-col md:flex-row gap-3">
                      <button 
                        onClick={() => syncDatabase({ sectors, employees, requests })}
                        className="flex-1 bg-green-600 hover:bg-green-700 text-white py-4 rounded-xl font-bold transition-all shadow-md active:scale-95 flex items-center justify-center gap-2"
                      >
                        <RefreshCw className={`w-5 h-5 ${isSyncing ? 'animate-spin' : ''}`} />
                        Sincronizar Agora (API)
                      </button>
                    </div>
                  </div>
                )}

                <div className="p-6 md:p-8 border-2 border-dashed border-blue-200 dark:border-blue-800 rounded-3xl bg-blue-50/10 dark:bg-blue-900/10 space-y-4">
                  <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
                    <div><h3 className="text-xl font-bold text-gray-800 dark:text-gray-200">Link de Acesso (24h)</h3><p className="text-sm text-gray-500 dark:text-gray-400">Cria um link temporário para preenchimento externo.</p></div>
                    <button onClick={generateAccessLink} className="w-full md:w-auto bg-blue-600 text-white px-6 py-3 rounded-xl font-bold flex items-center justify-center gap-2 shadow-lg hover:bg-blue-700 transition active:scale-95"><Share2 className="w-5 h-5" /> Gerar Link</button>
                  </div>
                  {generatedLink && <div className="flex items-center gap-2 bg-white dark:bg-gray-900 p-4 rounded-xl border border-blue-100 dark:border-gray-700"><input readOnly value={generatedLink} className="flex-1 text-xs text-gray-400 dark:text-gray-500 bg-transparent outline-none font-mono" /><button onClick={() => { navigator.clipboard.writeText(generatedLink); setAlertMessage('Copiado!'); }} className="text-blue-600 dark:text-blue-400 p-2"><Copy className="w-4 h-4" /></button></div>}
                </div>
                <div className="p-6 md:p-8 border border-gray-100 dark:border-gray-700 rounded-3xl bg-gray-50/50 dark:bg-gray-900/50 space-y-4">
                  <h3 className="text-xl font-bold text-gray-800 dark:text-gray-200">Identificação da Planilha</h3>
                  <p className="text-xs text-gray-400 dark:text-gray-500 flex items-center gap-2"><AlertCircle className="w-3 h-3" /> Insira o ID da Planilha ou a URL completa do Google Sheets.</p>
                  <input type="text" placeholder="Ex: https://docs.google.com/spreadsheets/d/..." className="w-full p-4 border dark:border-gray-700 rounded-xl bg-white dark:bg-gray-800 text-black dark:text-white outline-none text-base focus:border-blue-500" value={dbUrl} onChange={(e) => setDbUrl(e.target.value)} />
                  <div className="flex flex-col md:flex-row gap-3">
                    <button onClick={() => loadDatabase()} className="flex-1 bg-white dark:bg-gray-800 border border-blue-600 dark:border-blue-500 text-blue-600 dark:text-blue-400 px-6 py-4 rounded-xl font-bold active:bg-blue-50 dark:active:bg-gray-700 transition">Importar Dados</button>
                    <button onClick={exportToPDF} className="flex-1 bg-blue-600 text-white px-8 py-4 rounded-xl font-bold shadow-lg active:scale-95 transition">Exportar Drive</button>
                  </div>
                  <div className="mt-4">
                    <button onClick={executarFechamentoSemanal} className="w-full bg-red-600 hover:bg-red-700 text-white px-8 py-4 rounded-xl font-bold shadow-lg active:scale-95 transition flex items-center justify-center gap-2">
                      <AlertCircle className="w-5 h-5" />
                      FECHAMENTO SEMANAL (Salvar + Limpar Tudo)
                    </button>
                  </div>
                </div>

                <div className="p-6 md:p-8 border border-gray-100 dark:border-gray-700 rounded-3xl bg-gray-50/50 dark:bg-gray-900/50 space-y-4">
                  <h3 className="text-xl font-bold text-gray-800 dark:text-gray-200">Configuração do Apps Script</h3>
                  <p className="text-xs text-gray-400 dark:text-gray-500">Insira a URL do Apps Script e os IDs das pastas do Google Drive onde as fichas serão salvas.</p>
                  
                  <div className="space-y-4">
                    <div>
                      <label className="block text-sm font-bold text-gray-700 dark:text-gray-300 mb-1">URL do Apps Script (Para Exportação/Fechamento)</label>
                      <input type="text" placeholder="Ex: https://script.google.com/macros/s/..." className="w-full p-3 border dark:border-gray-700 rounded-xl bg-white dark:bg-gray-800 text-black dark:text-white outline-none text-sm font-mono focus:border-blue-500" value={scriptUrl} onChange={(e) => setScriptUrl(e.target.value)} />
                    </div>
                    <div>
                      <label className="block text-sm font-bold text-gray-700 dark:text-gray-300 mb-1">ID da Pasta (HE Registrado)</label>
                      <input type="text" className="w-full p-3 border dark:border-gray-700 rounded-xl bg-white dark:bg-gray-800 text-black dark:text-white outline-none text-sm font-mono focus:border-blue-500" value={folderRegId} onChange={(e) => { setFolderRegId(extractFolderId(e.target.value)); }} />
                    </div>
                    <div>
                      <label className="block text-sm font-bold text-gray-700 dark:text-gray-300 mb-1">ID da Pasta (HE Fixo)</label>
                      <input type="text" className="w-full p-3 border dark:border-gray-700 rounded-xl bg-white dark:bg-gray-800 text-black dark:text-white outline-none text-sm font-mono focus:border-blue-500" value={folderFixoId} onChange={(e) => { setFolderFixoId(extractFolderId(e.target.value)); }} />
                    </div>
                  </div>
                </div>
              </div>
            )}

            {state.adminSubView === 'FILES' && (
              <div className="bg-white dark:bg-gray-800 p-6 md:p-8 rounded-3xl border border-gray-100 dark:border-gray-700 space-y-8">
                <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
                  <h2 className="text-2xl font-bold dark:text-white flex items-center gap-2"><Folder className="text-blue-600 dark:text-blue-400" /> Arquivos Salvos</h2>
                  <div className="flex flex-wrap gap-2 w-full md:w-auto">
                    <button onClick={() => fetchDriveFiles(folderRegId, 'HE Registrado')} className="flex-1 md:flex-none bg-blue-50 dark:bg-blue-900/30 text-blue-600 dark:text-blue-400 px-4 py-2 rounded-lg font-bold hover:bg-blue-100 dark:hover:bg-blue-900/50 transition">Ver HE Registrado</button>
                    <button onClick={() => fetchDriveFiles(folderFixoId, 'HE Fixo')} className="flex-1 md:flex-none bg-blue-50 dark:bg-blue-900/30 text-blue-600 dark:text-blue-400 px-4 py-2 rounded-lg font-bold hover:bg-blue-100 dark:hover:bg-blue-900/50 transition">Ver HE Fixo</button>
                    <div className="flex bg-gray-100 dark:bg-gray-800 rounded-lg p-1 ml-auto md:ml-2">
                      <button onClick={() => setFileViewMode('list')} className={`p-1.5 rounded-md transition ${fileViewMode === 'list' ? 'bg-white dark:bg-gray-700 text-blue-600 dark:text-blue-400 shadow-sm' : 'text-gray-500 dark:text-gray-400 hover:text-gray-700 dark:hover:text-gray-300'}`} title="Lista"><List className="w-5 h-5" /></button>
                      <button onClick={() => setFileViewMode('grid')} className={`p-1.5 rounded-md transition ${fileViewMode === 'grid' ? 'bg-white dark:bg-gray-700 text-blue-600 dark:text-blue-400 shadow-sm' : 'text-gray-500 dark:text-gray-400 hover:text-gray-700 dark:hover:text-gray-300'}`} title="Grade"><Grid className="w-5 h-5" /></button>
                    </div>
                  </div>
                </div>

                {isLoadingFiles ? (
                  <div className="flex justify-center items-center p-12">
                    <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600"></div>
                  </div>
                ) : currentFolderId ? (
                  <div className="space-y-4">
                    {/* Breadcrumbs */}
                    <div className="flex items-center gap-2 text-sm text-gray-600 dark:text-gray-400 bg-gray-50 dark:bg-gray-900 p-3 rounded-xl overflow-x-auto">
                      {folderHistory.map((folder, index) => (
                        <div key={folder.id} className="flex items-center gap-2 whitespace-nowrap">
                          {index > 0 && <span>/</span>}
                          <button 
                            onClick={() => fetchDriveFiles(folder.id, folder.name)}
                            className={`hover:text-blue-600 dark:hover:text-blue-400 transition ${index === folderHistory.length - 1 ? 'font-bold text-gray-900 dark:text-gray-100' : ''}`}
                          >
                            {folder.name}
                          </button>
                        </div>
                      ))}
                    </div>

                    {/* Folders and Files List */}
                    <div className="bg-gray-50 dark:bg-gray-900 rounded-2xl border border-gray-100 dark:border-gray-700 overflow-hidden">
                      {driveFiles.folders.length === 0 && driveFiles.files.length === 0 ? (
                        <div className="p-8 text-center text-gray-500 dark:text-gray-400">
                          Pasta vazia
                        </div>
                      ) : fileViewMode === 'list' ? (
                        <ul className="divide-y divide-gray-100 dark:divide-gray-800">
                          {driveFiles.folders.map(folder => (
                            <li key={folder.id}>
                              <button 
                                onClick={() => fetchDriveFiles(folder.id, folder.name)}
                                className="w-full flex items-center gap-3 p-4 hover:bg-gray-100 dark:hover:bg-gray-800 transition text-left"
                              >
                                <Folder className="w-5 h-5 text-blue-500 flex-shrink-0" />
                                <span className="font-medium text-gray-800 dark:text-gray-200 truncate">{folder.name}</span>
                              </button>
                            </li>
                          ))}
                          {driveFiles.files.map(file => (
                            <li key={file.id} className="w-full flex items-center justify-between p-4 hover:bg-gray-100 dark:hover:bg-gray-800 transition border-b border-gray-100 dark:border-gray-800 last:border-0">
                              <a 
                                href={file.url} 
                                target="_blank" 
                                rel="noopener noreferrer"
                                className="flex items-center gap-3 flex-1"
                              >
                                <FileText className="w-5 h-5 text-red-500 flex-shrink-0" />
                                <span className="font-medium text-gray-800 dark:text-gray-200 truncate">{file.name}</span>
                              </a>
                              <button
                                onClick={(e) => { e.preventDefault(); printFile(file.id); }}
                                className="p-2 text-gray-500 hover:text-blue-600 dark:text-gray-400 dark:hover:text-blue-400 hover:bg-blue-50 dark:hover:bg-blue-900/30 rounded-lg transition ml-2"
                                title="Imprimir arquivo"
                              >
                                <Printer className="w-5 h-5" />
                              </button>
                            </li>
                          ))}
                        </ul>
                      ) : (
                        <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 gap-4 p-4">
                          {driveFiles.folders.map(folder => (
                            <button 
                              key={folder.id}
                              onClick={() => fetchDriveFiles(folder.id, folder.name)}
                              className="flex flex-col items-center justify-center p-4 bg-white dark:bg-gray-800 rounded-xl border border-gray-100 dark:border-gray-700 hover:border-blue-300 dark:hover:border-blue-700 hover:shadow-md transition gap-3 text-center group"
                            >
                              <Folder className="w-10 h-10 text-blue-500 group-hover:scale-110 transition-transform" />
                              <span className="font-medium text-sm text-gray-800 dark:text-gray-200 line-clamp-2">{folder.name}</span>
                            </button>
                          ))}
                          {driveFiles.files.map(file => (
                            <div key={file.id} className="relative group">
                              <a 
                                href={file.url} 
                                target="_blank" 
                                rel="noopener noreferrer"
                                className="flex flex-col items-center justify-center p-4 bg-white dark:bg-gray-800 rounded-xl border border-gray-100 dark:border-gray-700 hover:border-red-300 dark:hover:border-red-700 hover:shadow-md transition gap-3 text-center h-full"
                              >
                                <FileText className="w-10 h-10 text-red-500 group-hover:scale-110 transition-transform" />
                                <span className="font-medium text-sm text-gray-800 dark:text-gray-200 line-clamp-2">{file.name}</span>
                              </a>
                              <button
                                onClick={(e) => { e.preventDefault(); printFile(file.id); }}
                                className="absolute top-2 right-2 p-2 bg-white dark:bg-gray-700 text-gray-500 hover:text-blue-600 dark:text-gray-300 dark:hover:text-blue-400 rounded-full shadow-sm opacity-0 group-hover:opacity-100 transition-opacity border border-gray-200 dark:border-gray-600"
                                title="Imprimir arquivo"
                              >
                                <Printer className="w-4 h-4" />
                              </button>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  </div>
                ) : (
                  <div className="text-center p-12 border-2 border-dashed border-gray-200 dark:border-gray-700 rounded-3xl">
                    <Folder className="w-12 h-12 text-gray-400 mx-auto mb-4" />
                    <h3 className="text-lg font-bold text-gray-800 dark:text-gray-200 mb-2">Nenhuma pasta selecionada</h3>
                    <p className="text-gray-500 dark:text-gray-400">Selecione uma das pastas acima para visualizar os arquivos.</p>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      )}

      {showFormModal && (
        <div className="fixed inset-0 bg-white dark:bg-gray-900 md:bg-black/60 md:dark:bg-black/80 md:backdrop-blur-sm z-[60] flex items-center justify-center md:p-4 overflow-hidden">
          <div className="bg-white dark:bg-gray-800 md:rounded-3xl shadow-2xl w-full max-w-4xl h-full md:max-h-[90vh] overflow-y-auto p-4 md:p-8 relative flex flex-col">
            <button onClick={() => { setShowFormModal(false); setEditingRequestId(null); }} className="absolute top-4 right-4 text-gray-400 dark:text-gray-500 hover:text-gray-600 dark:hover:text-gray-300 bg-gray-100 dark:bg-gray-700 rounded-full p-1"><XCircle className="w-8 h-8" /></button>
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-6 mt-2 md:mt-0">
              <h2 className="text-2xl font-bold dark:text-white">Fechamento Semanal</h2>
              <div className="flex items-center gap-4 bg-gray-50 dark:bg-gray-900 p-2 rounded-xl border border-gray-100 dark:border-gray-700"><span className="text-sm font-medium text-gray-500 dark:text-gray-400">Semana:</span><input type="date" className="bg-transparent text-black dark:text-white outline-none text-sm font-bold" value={currentWeek} onChange={(e) => !editingRequestId && setCurrentWeek(e.target.value)} disabled={!!editingRequestId} /></div>
            </div>
            
            <div className="space-y-4 flex-1 overflow-y-auto pb-20">
              {modalRecords.map((r, idx) => {
                const activeRequestType = editingRequestId 
                  ? requests.find(r => r.id === editingRequestId)?.employeeType 
                  : state.flowType;
                  
                const isRegistradoFlow = activeRequestType === EmployeeType.REGISTRADO;

                return (
                  <div key={idx} className="bg-gray-50 dark:bg-gray-900 p-4 md:p-6 rounded-2xl border border-gray-100 dark:border-gray-700 shadow-sm space-y-3">
                    <div className="flex justify-between items-center border-b dark:border-gray-700 pb-2 border-gray-200">
                      <span className="font-bold text-gray-800 dark:text-gray-200 capitalize text-lg">{new Date(r.date + 'T00:00:00').toLocaleDateString('pt-BR', { weekday: 'short', day: '2-digit' })}</span>
                      {isRegistradoFlow && (
                        <button onClick={() => { const n = [...modalRecords]; n[idx].isFolgaVendida = !n[idx].isFolgaVendida; setModalRecords(n); }} className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all border ${r.isFolgaVendida ? 'bg-blue-600 border-blue-600 text-white' : 'bg-white dark:bg-gray-800 border-gray-300 dark:border-gray-600 text-gray-500 dark:text-gray-400'}`}>Folga Vendida</button>
                      )}
                    </div>
                    <div className={`grid gap-3 ${isRegistradoFlow && !r.isFolgaVendida ? 'grid-cols-2 md:grid-cols-4' : 'grid-cols-2'}`}>
                      <div className="flex flex-col"><label className="text-[10px] uppercase font-bold text-gray-400 dark:text-gray-500 mb-1">Entrada</label><input type="time" className="w-full p-3 border dark:border-gray-700 rounded-xl bg-white dark:bg-gray-800 text-black dark:text-white text-lg text-center font-bold outline-none focus:border-blue-500" value={r.realEntry} onChange={(e) => { const n = [...modalRecords]; n[idx].realEntry = e.target.value; setModalRecords(n); }} /></div>
                      {isRegistradoFlow && !r.isFolgaVendida && (
                        <>
                          <div className="flex flex-col"><label className="text-[10px] uppercase font-bold text-gray-400 dark:text-gray-500 mb-1 text-center">P. Ent</label><input type="time" className="w-full p-3 border dark:border-gray-700 rounded-xl bg-white dark:bg-gray-800 text-gray-500 dark:text-gray-400 text-lg text-center outline-none focus:border-blue-500" value={r.punchEntry} onChange={(e) => { const n = [...modalRecords]; n[idx].punchEntry = e.target.value; setModalRecords(n); }} /></div>
                          <div className="flex flex-col"><label className="text-[10px] uppercase font-bold text-gray-400 dark:text-gray-500 mb-1 text-center">P. Sai</label><input type="time" className="w-full p-3 border dark:border-gray-700 rounded-xl bg-white dark:bg-gray-800 text-gray-500 dark:text-gray-400 text-lg text-center outline-none focus:border-blue-500" value={r.punchExit} onChange={(e) => { const n = [...modalRecords]; n[idx].punchExit = e.target.value; setModalRecords(n); }} /></div>
                        </>
                      )}
                      <div className="flex flex-col"><label className="text-[10px] uppercase font-bold text-gray-400 dark:text-gray-500 mb-1 text-right">Saída</label><input type="time" className="w-full p-3 border dark:border-gray-700 rounded-xl bg-white dark:bg-gray-800 text-black dark:text-white text-lg text-center font-bold outline-none focus:border-blue-500" value={r.realExit} onChange={(e) => { const n = [...modalRecords]; n[idx].realExit = e.target.value; setModalRecords(n); }} /></div>
                    </div>
                  </div>
                );
              })}
            
                {editingRequestId && <div className="mt-4"><label className="block text-sm font-semibold mb-2 dark:text-gray-200">Justificativa da Edição</label><textarea className="w-full p-4 border dark:border-gray-700 rounded-xl bg-white dark:bg-gray-800 text-black dark:text-white text-base focus:border-blue-500 outline-none" rows={3} value={editJustification} onChange={(e) => setEditJustification(e.target.value)} placeholder="Por que você está alterando este registro?" /></div>}
            </div>

            <div className="mt-4 pt-4 border-t border-gray-100 dark:border-gray-700 flex gap-3 bg-white dark:bg-gray-800 sticky bottom-0 z-10 pb-6 md:pb-0">
                <button onClick={() => { setShowFormModal(false); setEditingRequestId(null); }} className="flex-1 py-4 bg-gray-100 dark:bg-gray-700 text-gray-700 dark:text-gray-300 font-bold rounded-xl active:scale-95 transition">Cancelar</button>
                <button onClick={submitRequest} className="flex-[2] py-4 bg-blue-600 text-white font-bold rounded-xl shadow-lg active:scale-95 transition">Salvar</button>
            </div>
          </div>
        </div>
      )}

      {/* Alert Modal */}
      {alertMessage && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white dark:bg-gray-800 p-6 rounded-2xl shadow-xl max-w-sm w-full text-center space-y-4">
            <p className="text-gray-800 dark:text-gray-200 font-medium">{alertMessage}</p>
            <button onClick={() => setAlertMessage(null)} className="w-full bg-blue-600 text-white py-3 rounded-xl font-bold hover:bg-blue-700 transition">OK</button>
          </div>
        </div>
      )}

      {/* Confirm Modal */}
      {confirmDialog && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white dark:bg-gray-800 p-6 rounded-2xl shadow-xl max-w-sm w-full text-center space-y-6">
            <p className="text-gray-800 dark:text-gray-200 font-medium whitespace-pre-line">{confirmDialog.message}</p>
            <div className="flex gap-3">
              <button onClick={() => setConfirmDialog(null)} className="flex-1 bg-gray-100 dark:bg-gray-700 text-gray-700 dark:text-gray-300 py-3 rounded-xl font-bold hover:bg-gray-200 dark:hover:bg-gray-600 transition">Cancelar</button>
              <button onClick={() => { confirmDialog.onConfirm(); setConfirmDialog(null); }} className="flex-1 bg-red-600 text-white py-3 rounded-xl font-bold hover:bg-red-700 transition">Confirmar</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default App;
