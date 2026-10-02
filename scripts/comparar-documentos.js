#!/usr/bin/env node
// =============================================
// Compara dos documentos generados (.xlsx o .docx) para probar equivalencia
// entre dos despliegues del constructor (p. ej. produccion vs 207).
//
//  1. Por partes del ZIP: cada parte descomprimida debe ser identica byte a
//     byte, salvo los metadatos (docProps/core.xml: fechas de creacion).
//  2. Solo .xlsx: celda a celda con ExcelJS en todas las hojas: valor,
//     formula, formato numerico, estilo, combinadas y dimensiones.
//
// Uso: node scripts/comparar-documentos.js a.xlsx b.xlsx [c.xlsx d.xlsx ...]
// Sale con 1 si algun par difiere.
// =============================================

const fs = require('fs');
const AdmZip = require('adm-zip');
const ExcelJS = require('exceljs');

const PARTES_METADATOS = new Set(['docProps/core.xml']);

function compararZip(a, b) {
  const za = new AdmZip(a);
  const zb = new AdmZip(b);
  const ea = new Map(za.getEntries().map((e) => [e.entryName, e]));
  const eb = new Map(zb.getEntries().map((e) => [e.entryName, e]));
  const difs = [];
  for (const nombre of new Set([...ea.keys(), ...eb.keys()])) {
    if (PARTES_METADATOS.has(nombre)) continue;
    if (!ea.has(nombre) || !eb.has(nombre)) {
      difs.push(`parte ${nombre} solo en ${ea.has(nombre) ? 'A' : 'B'}`);
      continue;
    }
    if (!ea.get(nombre).getData().equals(eb.get(nombre).getData())) difs.push(`parte ${nombre} distinta`);
  }
  return { partes: ea.size, difs };
}

const json = (v) => JSON.stringify(v === undefined ? null : v);

async function compararCeldas(a, b) {
  const wa = new ExcelJS.Workbook();
  const wb = new ExcelJS.Workbook();
  await wa.xlsx.load(fs.readFileSync(a));
  await wb.xlsx.load(fs.readFileSync(b));
  const difs = [];
  let celdas = 0;

  const hojasA = wa.worksheets.map((w) => w.name);
  const hojasB = wb.worksheets.map((w) => w.name);
  if (json(hojasA) !== json(hojasB)) difs.push(`hojas distintas: ${json(hojasA)} vs ${json(hojasB)}`);

  for (const sa of wa.worksheets) {
    const sb = wb.getWorksheet(sa.name);
    if (!sb) continue;
    if (sa.dimensions?.range !== sb.dimensions?.range) difs.push(`[${sa.name}] dimensiones ${sa.dimensions?.range} vs ${sb.dimensions?.range}`);
    const ma = json([...(sa.model.merges || [])].sort());
    const mb = json([...(sb.model.merges || [])].sort());
    if (ma !== mb) difs.push(`[${sa.name}] celdas combinadas distintas`);

    const direcciones = new Set();
    for (const s of [sa, sb]) s.eachRow({ includeEmpty: true }, (row) => row.eachCell({ includeEmpty: true }, (c) => direcciones.add(c.address)));
    for (const addr of direcciones) {
      celdas++;
      const ca = sa.getCell(addr);
      const cb = sb.getCell(addr);
      for (const [campo, va, vb] of [
        ['valor', json(ca.value), json(cb.value)],
        ['formula', json(ca.formula), json(cb.formula)],
        ['numFmt', json(ca.numFmt), json(cb.numFmt)],
        ['estilo', json(ca.style), json(cb.style)]
      ]) {
        if (va !== vb) difs.push(`[${sa.name}!${addr}] ${campo}: ${va.slice(0, 80)} vs ${vb.slice(0, 80)}`);
      }
    }
  }
  return { celdas, difs };
}

async function compararPar(a, b) {
  const zip = compararZip(a, b);
  const res = { a, b, partes: zip.partes, difs: [...zip.difs], celdas: null };
  if (a.toLowerCase().endsWith('.xlsx')) {
    const c = await compararCeldas(a, b);
    res.celdas = c.celdas;
    res.difs.push(...c.difs);
  }
  return res;
}

async function main() {
  const archivos = process.argv.slice(2);
  if (archivos.length < 2 || archivos.length % 2) {
    console.error('Uso: comparar-documentos.js a1 b1 [a2 b2 ...]');
    process.exit(2);
  }
  let distintos = 0;
  for (let i = 0; i < archivos.length; i += 2) {
    const r = await compararPar(archivos[i], archivos[i + 1]);
    const detalle = r.celdas === null ? `${r.partes} partes` : `${r.partes} partes, ${r.celdas} celdas`;
    if (r.difs.length) {
      distintos++;
      console.log(`DISTINTO  ${r.a} vs ${r.b} (${detalle}, ${r.difs.length} diferencias)`);
      for (const d of r.difs.slice(0, 20)) console.log(`   - ${d}`);
    } else {
      console.log(`IGUAL     ${r.a} vs ${r.b} (${detalle})`);
    }
  }
  process.exit(distintos ? 1 : 0);
}

if (require.main === module) {
  main().catch((e) => {
    console.error(`ERROR: ${e.message}`);
    process.exit(2);
  });
}

module.exports = { compararPar };
