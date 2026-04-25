
export enum EmployeeType {
  REGISTRADO = 'REGISTRADO',
  FIXO = 'FIXO'
}

export enum RequestStatus {
  PENDENTE = 'PENDENTE',
  APROVADO = 'APROVADO',
  REJEITADO = 'REJEITADO',
  DELETADO = 'DELETADO'
}

export interface Sector {
  id: string;
  name: string;
  fixedRate: number; // Value for "Fixo" employees in this sector
}

export interface Employee {
  id: string;
  name: string;
  sectorId: string;
  type: EmployeeType;
  salary: number;
  monthlyHours: number;
  fixedDayOff?: number;
}

export interface TimeRecord {
  date: string;
  realEntry: string;
  punchEntry: string;
  punchExit: string;
  realExit: string;
  isFolgaVendida: boolean;
}

export interface TimeRequest {
  id: string;
  employeeId: string;
  employeeName: string;
  employeeType: EmployeeType;
  sectorId: string;
  sectorName: string;
  weekStarting: string;
  records: TimeRecord[];
  status: RequestStatus;
  calculatedValue: number;
  totalTimeDecimal: number;
  createdAt: string;
  editJustification?: string;
}

export interface AppState {
  view: 'HOME' | 'SELECTION' | 'FLOW' | 'ADMIN' | 'EXPIRED' | 'SUCCESS';
  flowType: EmployeeType | null;
  adminSubView: 'DASHBOARD' | 'SECTORS' | 'EMPLOYEES' | 'REQUESTS' | 'INTEGRATIONS';
}
