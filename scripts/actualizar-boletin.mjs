#!/usr/bin/env node
// Procesa los boletines .xlsm/.xlsx que estén en boletines/ (subidos a mano cada mes) y
// actualiza data/creditos-sector.json. No hace ningún pedido de red: el sitio del BCP
// bloquea descargas automatizadas (403 aun desde los runners de GitHub Actions), así que
// el paso de "bajar el boletín" lo hace una persona en su navegador — esta Action solo
// extrae y publica los datos una vez que el archivo ya está en el repo.
//
// Flujo: bajás el boletín del BCP como siempre, lo subís a la carpeta boletines/ del repo
// (Add file → Upload files en GitHub), y el push dispara esta Action sola.

import { readFile, writeFile, mkdir, readdir, unlink } from 'node:fs/promises';
import path from 'node:path';
import * as XLSX from 'xlsx';

const BOLETINES_DIR = path.join(process.cwd(), 'boletines');
const DATA_PATH = path.join(process.cwd(), 'data', 'creditos-sector.json');

function normalizeAscii(s){
  return String(s).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
}

/* ============================== parser de "5. Cred. por sector" ============================== */
/* Misma lógica que usa index.html en el navegador (mantener ambas en sync si se cambia una). */

const BANK_ALIASES = { 'Amambay':'Basa', 'FINEXPAR':'Zeta' };

function parseLocaleNumber(v){
  if (v === null || v === undefined) return NaN;
  if (typeof v === 'number') return v;
  let s = String(v).trim();
  if (s === '' || s === '-') return NaN;
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1,-1); }
  s = s.replace(/Gs\.?/gi,'').replace(/%/g,'').trim();
  const hasComma = s.includes(','), hasDot = s.includes('.');
  if (hasComma && hasDot){ s = s.replace(/\./g,'').replace(',', '.'); }
  else if (hasComma && !hasDot){ s = s.replace(',', '.'); }
  else if (hasDot && !hasComma){
    const parts = s.split('.');
    if (parts.length > 2){ s = s.replace(/\./g,''); }
    else if (parts[1] && parts[1].length === 3){ s = s.replace('.', ''); }
  }
  const num = parseFloat(s);
  if (isNaN(num)) return NaN;
  return neg ? -num : num;
}

function normalizePeriodValue(raw){
  if (raw === null || raw === undefined || raw === '') return null;
  if (typeof raw === 'number'){
    const ms = Math.round((raw - 25569) * 86400 * 1000);
    const d = new Date(ms);
    if (!isNaN(d.getTime())) return d.getUTCFullYear() + '-' + String(d.getUTCMonth()+1).padStart(2,'0');
    return String(raw);
  }
  const s = String(raw).trim();
  let m = s.match(/^(\d{4})[-/](\d{1,2})/);
  if (m) return m[1] + '-' + m[2].padStart(2,'0');
  return s;
}

function parseBcpBulletinSheet(rows){
  let headerRow = -1, fechaCol = -1;
  outer:
  for (let r=0; r<rows.length; r++){
    const row = rows[r]||[];
    for (let c=0; c<row.length-1; c++){
      const fecha = row[c], sector = row[c+1];
      if (fecha && String(fecha).trim().toLowerCase()==='fecha' && sector && String(sector).trim().toLowerCase().includes('sector')){
        headerRow = r; fechaCol = c; break outer;
      }
    }
  }
  if (headerRow===-1) return { error:'No encontramos la tabla de "Cred. por sector" en esta hoja.' };
  const sectorCol = fechaCol+1;

  const header = rows[headerRow]||[];
  const headerAbove = rows[headerRow-1]||[];
  const bankCols = [];
  const lastCol = Math.max(header.length, headerAbove.length);
  for (let c=sectorCol+1; c<lastCol; c++){
    let name = header[c];
    if (name===null || name===undefined || String(name).trim()===''){ name = headerAbove[c]; }
    if (name!==null && name!==undefined && String(name).trim()!==''){ bankCols.push({ col:c, name:String(name).trim() }); }
  }
  if (bankCols.length===0) return { error:'No encontramos columnas de bancos.' };

  const records = [];
  let period = null;
  for (let r=headerRow+1; r<rows.length; r++){
    const row = rows[r]||[];
    const fechaVal = row[fechaCol];
    const sectorVal = row[sectorCol];
    let isTotalRow = false;
    if (fechaVal!==null && fechaVal!==undefined && String(fechaVal).trim()!==''){
      const s = String(fechaVal).trim();
      if (/^total/i.test(s)){
        isTotalRow = true;
        const m = s.match(/(\d{4}[-/]\d{1,2})/);
        if (m) period = normalizePeriodValue(m[1]);
      } else {
        period = normalizePeriodValue(fechaVal);
      }
    }
    const sectorName = isTotalRow ? 'Total cartera' : (sectorVal!==null && sectorVal!==undefined ? String(sectorVal).trim() : '');
    if (!sectorName || !period) continue;
    bankCols.forEach(bc=>{
      const num = parseLocaleNumber(row[bc.col]);
      if (isNaN(num)) return;
      const banco = BANK_ALIASES[bc.name] || bc.name;
      records.push({ periodo: period, sector: sectorName, banco, monto: num });
    });
  }
  if (records.length===0) return { error:'No encontramos filas de sectores.' };
  return { records };
}

function findBulletinSheetName(sheetNames){
  return sheetNames.find(n=>{
    const norm = normalizeAscii(n);
    return norm.includes('cred') && norm.includes('sector');
  });
}

/* ============================== main ============================== */

async function main(){
  let files;
  try{
    files = await readdir(BOLETINES_DIR);
  }catch(err){
    console.log('No existe la carpeta boletines/ todavía (o está vacía). Nada que procesar.');
    return;
  }
  const bulletinFiles = files.filter(f => /\.(xlsm|xlsx)$/i.test(f));
  if (bulletinFiles.length === 0){
    console.log('No hay archivos .xlsm/.xlsx en boletines/. Nada que procesar.');
    return;
  }

  let existing = { records: [] };
  try{
    existing = JSON.parse(await readFile(DATA_PATH, 'utf-8'));
    if (!Array.isArray(existing.records)) existing.records = [];
  }catch(err){ /* primera corrida: todavía no existe data/creditos-sector.json */ }

  const merged = new Map();
  existing.records.forEach(r=> merged.set(r.periodo+'|'+r.sector+'|'+r.banco, r));

  let procesados = 0;
  for (const file of bulletinFiles){
    const filePath = path.join(BOLETINES_DIR, file);
    console.log('Procesando', file, '...');
    try{
      const buf = await readFile(filePath);
      const wb = XLSX.read(buf, { type: 'buffer', cellDates: false });
      const sheetName = findBulletinSheetName(wb.SheetNames);
      if (!sheetName){ console.error('  Se salteó (no se encontró la hoja "5. Cred. por sector"):', file); continue; }
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header:1, raw:true, defval:null });
      const result = parseBcpBulletinSheet(rows);
      if (result.error){ console.error('  Se salteó', file, '-', result.error); continue; }
      result.records.forEach(r=> merged.set(r.periodo+'|'+r.sector+'|'+r.banco, r));
      console.log('  OK:', result.records.length, 'filas de', [...new Set(result.records.map(r=>r.periodo))].join(', '));
      procesados++;
      // ya está extraído en data/creditos-sector.json; borramos el .xlsm original del repo
      // para no ir acumulando archivos de ~15 MB cada mes.
      await unlink(filePath);
    }catch(err){
      console.error('  Se salteó', file, '- error inesperado:', err.message);
    }
  }

  if (procesados === 0){
    console.log('Ningún archivo se pudo procesar. No se modifica data/creditos-sector.json.');
    return;
  }

  const salida = {
    actualizado: new Date().toISOString(),
    records: Array.from(merged.values()),
  };

  await mkdir(path.dirname(DATA_PATH), { recursive: true });
  await writeFile(DATA_PATH, JSON.stringify(salida, null, 2) + '\n', 'utf-8');
  console.log('Guardado', DATA_PATH, 'con', salida.records.length, 'filas totales (', procesados, 'boletines procesados en esta corrida).');
}

main().catch(err=>{
  console.error('FALLÓ el procesamiento de boletines:', err.message);
  process.exit(1);
});
