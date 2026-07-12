#!/usr/bin/env node
// Descarga el boletín de bancos más reciente publicado por el BCP, extrae la hoja
// "5. Cred. por sector" y actualiza data/creditos-sector.json. Pensado para correr
// mensualmente desde un GitHub Action (.github/workflows/actualizar-boletin.yml).
//
// Estrategia: la página de boletines del BCP (Liferay) enlaza a documentos con URLs
// del tipo /documents/{grupo}/{carpeta}/{nombre}.xlsm/{uuid}?t=... donde carpeta y uuid
// cambian cada vez que suben un archivo nuevo, así que no se puede predecir el link del
// próximo mes: hay que leer la página cada vez y quedarse con el enlace de "Bancos" más
// reciente según el mes/año que trae el propio nombre del archivo (ej. "May26").
//
// Ejemplo real de link (mayo 2026):
// https://www.bcp.gov.py/documents/20117/2502605/1.+Bolet%C3%ADn+Bancos+May26+4.xlsm/0c1adf04-6a34-cc04-d188-8c14229c8898?t=1782317455148

import { readFile, writeFile, mkdir } from 'node:fs/promises';
import path from 'node:path';
import * as XLSX from 'xlsx';

const LISTADO_URL = 'https://www.bcp.gov.py/web/institucional/boletines-formato-macros';
const DATA_PATH = path.join(process.cwd(), 'data', 'creditos-sector.json');
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36';

const MESES = {
  ene:1, jan:1, feb:2, mar:3, abr:4, apr:4, may:5, jun:6, jul:7,
  ago:8, aug:8, sep:9, set:9, oct:10, nov:11, dic:12, dec:12,
};

/* ============================== localizar el boletín vigente ============================== */

async function fetchText(url){
  const res = await fetch(url, { headers: { 'User-Agent': UA, 'Accept-Language': 'es-PY,es;q=0.9' } });
  if (!res.ok) throw new Error('HTTP ' + res.status + ' al pedir ' + url);
  return res.text();
}

function decodeHtmlEntities(s){
  return s.replace(/&amp;/g,'&').replace(/&quot;/g,'"').replace(/&#39;/g,"'").replace(/&lt;/g,'<').replace(/&gt;/g,'>');
}

function normalizeAscii(s){
  return String(s).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
}

// Extrae "May" + "26" de nombres tipo "...Bancos+May26+4.xlsm" o "...Bancos_May_2026...xlsm"
function parseMonthYearFromName(name){
  const decoded = decodeURIComponent(name);
  let m = decoded.match(/([A-Za-z]{3,4})[_\s.+-]?(\d{2,4})(?!\d)/);
  if (!m) return null;
  const mes = MESES[normalizeAscii(m[1]).slice(0,3)];
  if (!mes) return null;
  let anio = parseInt(m[2], 10);
  if (anio < 100) anio += 2000;
  return { anio, mes };
}

async function findLatestBulletinUrl(){
  const html = await fetchText(LISTADO_URL);
  // hrefs a documentos Liferay que terminan en .xlsm/<uuid>, con o sin querystring de cache-busting
  const re = /href="([^"]*\/documents\/\d+\/\d+\/[^"]*?\.xlsm\/[0-9a-fA-F-]{20,36}(?:\?[^"]*)?)"/gi;
  const candidatos = [];
  let match;
  while ((match = re.exec(html)) !== null){
    const hrefRaw = decodeHtmlEntities(match[1]);
    const href = hrefRaw.startsWith('http') ? hrefRaw : new URL(hrefRaw, LISTADO_URL).toString();
    const nombreSegmento = decodeURIComponent(href.split('/documents/')[1].split('/').find(seg => seg.toLowerCase().endsWith('.xlsm')) || '');
    const nombreNorm = normalizeAscii(nombreSegmento);
    // nos interesa el boletín de "Bancos" (no Financieras, Cooperativas, Casas de Cambio, etc.)
    if (!nombreNorm.includes('banco')) continue;
    if (!nombreNorm.includes('bolet')) continue;
    const my = parseMonthYearFromName(nombreSegmento);
    candidatos.push({ href, nombre: nombreSegmento, anio: my?.anio ?? 0, mes: my?.mes ?? 0 });
  }
  if (candidatos.length === 0){
    console.error('No se encontró ningún enlace de boletín de Bancos en la página. Volcado parcial del HTML para depurar:');
    console.error(html.slice(0, 4000));
    throw new Error('No se pudo ubicar el boletín de Bancos en ' + LISTADO_URL);
  }
  candidatos.sort((a,b)=> (b.anio*12+b.mes) - (a.anio*12+a.mes));
  return candidatos[0];
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
  console.log('Buscando el boletín de Bancos más reciente en', LISTADO_URL);
  const latest = await findLatestBulletinUrl();
  console.log('Encontrado:', latest.nombre, '(', latest.anio + '-' + String(latest.mes).padStart(2,'0'), ') ->', latest.href);

  const res = await fetch(latest.href, { headers: { 'User-Agent': UA } });
  if (!res.ok) throw new Error('No se pudo descargar el boletín: HTTP ' + res.status);
  const buf = Buffer.from(await res.arrayBuffer());
  console.log('Descargado', buf.length, 'bytes');

  const wb = XLSX.read(buf, { type: 'buffer', cellDates: false });
  const sheetName = findBulletinSheetName(wb.SheetNames);
  if (!sheetName) throw new Error('No se encontró la hoja "5. Cred. por sector" en el archivo descargado. Hojas disponibles: ' + wb.SheetNames.join(', '));
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header:1, raw:true, defval:null });
  const result = parseBcpBulletinSheet(rows);
  if (result.error) throw new Error('Error al parsear la hoja "' + sheetName + '": ' + result.error);
  console.log('Extraídas', result.records.length, 'filas de "' + sheetName + '"');

  let existing = { records: [] };
  try{
    existing = JSON.parse(await readFile(DATA_PATH, 'utf-8'));
    if (!Array.isArray(existing.records)) existing.records = [];
  }catch(err){ /* primera corrida: todavía no existe data/creditos-sector.json */ }

  const merged = new Map();
  existing.records.forEach(r=> merged.set(r.periodo+'|'+r.sector+'|'+r.banco, r));
  result.records.forEach(r=> merged.set(r.periodo+'|'+r.sector+'|'+r.banco, r));

  const salida = {
    actualizado: new Date().toISOString(),
    fuente: latest.href,
    records: Array.from(merged.values()),
  };

  await mkdir(path.dirname(DATA_PATH), { recursive: true });
  await writeFile(DATA_PATH, JSON.stringify(salida, null, 2) + '\n', 'utf-8');
  console.log('Guardado', DATA_PATH, 'con', salida.records.length, 'filas totales.');
}

main().catch(err=>{
  console.error('FALLÓ la actualización del boletín:', err.message);
  process.exit(1);
});
