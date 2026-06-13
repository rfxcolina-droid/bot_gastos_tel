const ExcelJS = require("exceljs");
const fs      = require("fs");
const path    = require("path");
const { subirExcel } = require("./drive");

const DIR   = path.resolve("./datos");
const EXCEL = path.join(DIR, "Registro_Gastos.xlsx");
const CTR   = path.join(DIR, "correlativo.json");

fs.mkdirSync(DIR, { recursive: true });

function siguienteCorr() {
  try { return JSON.parse(fs.readFileSync(CTR, "utf8")).next || 1; } catch { return 1; }
}
function avanzarCorr(n) {
  fs.writeFileSync(CTR, JSON.stringify({ next: n + 1 }), "utf8");
}

async function guardarGasto(g) {
  const wb     = new ExcelJS.Workbook();
  const HCOLOR = "FF128C7E";

  if (fs.existsSync(EXCEL)) {
    await wb.xlsx.readFile(EXCEL);
  } else {
    const ws = wb.addWorksheet("Gastos");
    ws.columns = [
      { header: "Correlativo", key: "correlativo", width: 14 },
      { header: "Fecha",       key: "fecha",        width: 13 },
      { header: "Monto",       key: "monto",        width: 14 },
      { header: "Moneda",      key: "moneda",        width: 8  },
      { header: "Comercio",    key: "comercio",      width: 22 },
      { header: "Motivo",      key: "motivo",        width: 20 },
      { header: "Destino",     key: "destino",       width: 20 },
      { header: "Detalle",     key: "detalle",       width: 30 },
      { header: "Foto",        key: "foto",          width: 20 },
      { header: "URL Foto",    key: "fotoUrl",       width: 45 },
    ];
    ws.getRow(1).eachCell(c => {
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HCOLOR } };
      c.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
      c.alignment = { horizontal: "center", vertical: "middle" };
    });
    ws.getRow(1).height = 22;
    wb.addWorksheet("Por Motivo");
    wb.addWorksheet("Por Destino");
  }

  const ws  = wb.getWorksheet("Gastos");
  const tot = Math.max((ws.lastRow?.number || 1) - 1, 0);

  const fila = ws.addRow({
    correlativo: `GASTO_${String(g.corr).padStart(4, "0")}`,
    fecha:    g.fecha    || "",
    monto:    g.monto    || 0,
    moneda:   g.moneda   || "CLP",
    comercio: g.comercio || "",
    motivo:   g.motivo   || "",
    destino:  g.destino  || "",
    detalle:  g.detalle  || "",
    foto:     g.fotoNombre || "",
    fotoUrl:  g.fotoUrl  || "",
  });
  fila.getCell("monto").numFmt = "#,##0";
  if (fila.number % 2 === 0)
    fila.eachCell({ includeEmpty: true }, c => {
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F8F6" } };
    });

  // Resúmenes
  for (const [hoja, campo] of [["Por Motivo","motivo"],["Por Destino","destino"]]) {
    let wr = wb.getWorksheet(hoja) || wb.addWorksheet(hoja);
    wr.spliceRows(1, wr.lastRow?.number || 0);
    wr.columns = [
      { header: campo === "motivo" ? "Motivo" : "Destino", key: "cat",   width: 26 },
      { header: "Total CLP",                                key: "total", width: 16 },
    ];
    wr.getRow(1).eachCell(c => {
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HCOLOR } };
      c.font = { bold: true, color: { argb: "FFFFFFFF" } };
    });
    const mp = {};
    ws.eachRow((r, i) => {
      if (i === 1) return;
      const k = String(r.getCell(campo).value || "Sin clasificar");
      mp[k] = (mp[k] || 0) + (Number(r.getCell("monto").value) || 0);
    });
    Object.entries(mp).sort((a, b) => b[1] - a[1]).forEach(([k, v]) => {
      const r = wr.addRow({ cat: k, total: v });
      r.getCell("total").numFmt = "#,##0";
    });
  }

  await wb.xlsx.writeFile(EXCEL);

  // Subir Excel actualizado a Google Drive
  try {
    await subirExcel(EXCEL, "Registro_Gastos.xlsx");
  } catch (e) {
    console.error("Error subiendo Excel a Drive:", e.message);
  }

  return tot + 1;
}

module.exports = { guardarGasto, siguienteCorr, avanzarCorr, EXCEL };
