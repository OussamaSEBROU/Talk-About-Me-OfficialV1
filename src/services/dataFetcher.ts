import * as XLSX from 'xlsx';
import { MemorialPerson } from '../types';

export const fetchAndParseData = async (): Promise<MemorialPerson[]> => {
  try {
    const response = await fetch('/victims_data.xlsx');
    if (!response.ok) throw new Error('Failed to fetch Excel file');
    
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    let data = XLSX.utils.sheet_to_json<MemorialPerson>(workbook.Sheets[sheetName], { range: 1 });
    
    // Generate offset lat/lng for points if missing
    data = data.map((p, index) => {
       const t = Math.random();
       const w = (Math.random() - 0.5) * 0.08;
       return {
         ...p,
         lat: p.lat ?? (31.23 + (t * 0.34) + w * -0.66), 
         lng: p.lng ?? (34.22 + (t * 0.30) + w * 0.75),
         // Optional fallback id/index if missing in excel
         Index: p.Index || String(index + 1),
         ID: p.ID || "00000"
       };
    });
    
    return data;
  } catch (error) {
    console.error('Failed to parse Excel file, returning empty array:', error);
    return [];
  }
};
