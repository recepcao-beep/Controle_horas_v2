
export const parseCurrency = (value: any): number => {
  if (value === undefined || value === null) return 0;
  if (typeof value === 'number') return (isNaN(value) || !isFinite(value)) ? 0 : value;
  if (typeof value === 'string') {
    let cleanStr = value.replace(/[^\d.,-]/g, '');
    
    const lastCommaIndex = cleanStr.lastIndexOf(',');
    const lastDotIndex = cleanStr.lastIndexOf('.');
    
    if (lastCommaIndex > lastDotIndex) {
      // e.g. "1.234,56" -> remove dots, replace comma with dot
      cleanStr = cleanStr.replace(/\./g, '').replace(',', '.');
    } else if (lastDotIndex > lastCommaIndex && lastCommaIndex !== -1) {
      // e.g. "1,234.56" -> remove commas
      cleanStr = cleanStr.replace(/,/g, '');
    } else if (lastCommaIndex !== -1) {
      // e.g. "1234,56" -> replace comma with dot
      cleanStr = cleanStr.replace(',', '.');
    }
    
    const parsed = parseFloat(cleanStr);
    return isNaN(parsed) ? 0 : parsed;
  }
  return 0;
};

export const formatCurrency = (value: any) => {
  if (value === undefined || value === null || value === '' || value === 'N/A') return 'R$ 0,00';
  const num = parseCurrency(value);
  
  if (!isFinite(num)) return 'R$ 0,00';
  
  return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(num);
};

export const timeToDecimal = (time: any): number => {
  if (!time || typeof time !== 'string' || !time.includes(':')) return 0;
  const parts = time.split(':');
  const hours = parseInt(parts[0], 10);
  const minutes = parseInt(parts[1], 10);
  if (isNaN(hours) || isNaN(minutes)) return 0;
  return hours + minutes / 60;
};

export const formatDecimalHours = (decimal: number): string => {
  if (isNaN(decimal) || !isFinite(decimal)) return '0h 00m';
  const totalMinutes = Math.round(decimal * 60);
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  return `${hours}h ${minutes.toString().padStart(2, '0')}m`;
};

export const calculateTimeDiff = (real: string, punch: string): number => {
  return Math.abs(timeToDecimal(real) - timeToDecimal(punch));
};

export const getWeekDays = (baseDate: Date) => {
  const days = [];
  const d = new Date(baseDate);
  d.setHours(0, 0, 0, 0);
  
  const day = d.getDay(); // 0 (Dom) a 6 (Sab)
  // Calcula a diferença para chegar na Segunda-feira (1)
  // Se for Domingo (0), volta 6 dias. Se for Segunda (1), volta 0. Se for Terça (2), volta 1...
  const diffToMonday = day === 0 ? -6 : 1 - day;
  d.setDate(d.getDate() + diffToMonday);

  for (let i = 0; i < 7; i++) {
    const dayInstance = new Date(d);
    dayInstance.setDate(d.getDate() + i);
    days.push(dayInstance.toISOString().split('T')[0]);
  }
  return days;
};
