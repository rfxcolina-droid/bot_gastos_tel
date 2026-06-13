require("dotenv").config();
const TelegramBot = require("node-telegram-bot-api");
const ExcelJS     = require("exceljs");
const axios       = require("axios");
const fs          = require("fs");
const path        = require("path");

const bot   = new TelegramBot(process.env.TELEGRAM_TOKEN, { polling: true });
const DATOS = path.resolve("./datos");
const EXCEL = path.join(DATOS, "Registro_Gastos.xlsx");
const CTR   = path.join(DATOS, "correlativo.json");

fs.mkdirSync(DATOS, { recursive: true });

// ── Correlativo ──────────────────────────────────────────
function siguienteCorr() {
  try { return JSON.parse(fs.readFileSync(CTR, "utf8")).next || 1; } catch { return 1; }
}
function avanzarCorr(n) {
  fs.writeFileSync(CTR, JSON.stringify({ next: n + 1 }), "utf8");
}

// ── Sesiones ─────────────────────────────────────────────
const S = {};
function ses(id) { if (!S[id]) S[id] = { paso: "inicio", d: {} }; return S[id]; }
function setPaso(id, paso, d) { S[id] = { paso, d: d !== undefined ? d : (S[id]?.d || {}) }; }
function reset(id) { S[id] = { paso: "inicio", d: {} }; }

// ── Guardar en Excel ─────────────────────────────────────
async function guardarExcel(g) {
  try {
    const wb = new ExcelJS.Workbook();
    const HCOLOR = "FF128C7E";

    if (fs.existsSync(EXCEL)) {
      await wb.xlsx.readFile(EXCEL);
    } else {
      // Crear planilla nueva
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
      ];
      ws.getRow(1).eachCell(cell => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HCOLOR } };
        cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
        cell.alignment = { horizontal: "center", vertical: "middle" };
      });
      ws.getRow(1).height = 22;
      wb.addWorksheet("Por Motivo");
      wb.addWorksheet("Por Destino");
    }

    const ws  = wb.getWorksheet("Gastos");
    const tot = Math.max((ws.lastRow?.number || 1) - 1, 0);

    const fila = ws.addRow({
      correlativo: "GASTO_" + String(g.corr).padStart(4, "0"),
      fecha:    g.fecha   || "",
      monto:    g.monto   || 0,
      moneda:   g.moneda  || "CLP",
      comercio: g.comercio|| "",
      motivo:   g.motivo  || "",
      destino:  g.destino || "",
      detalle:  g.detalle || "",
    });
    fila.getCell("monto").numFmt = "#,##0";
    if (fila.number % 2 === 0) {
      fila.eachCell({ includeEmpty: true }, c => {
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F8F6" } };
      });
    }

    // Actualizar resúmenes
    for (const [hoja, campo] of [["Por Motivo","motivo"], ["Por Destino","destino"]]) {
      let wr = wb.getWorksheet(hoja);
      if (!wr) wr = wb.addWorksheet(hoja);
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
    console.log("Excel guardado OK. Total registros:", tot + 1);
    return tot + 1;

  } catch (err) {
    console.error("ERROR guardando Excel:", err.message);
    throw err;
  }
}

// ── Analizar imagen con Claude ────────────────────────────
async function analizarImagen(fileId) {
  // 1. Obtener info del archivo desde Telegram
  const fileInfo = await bot.getFile(fileId);
  const fileUrl  = `https://api.telegram.org/file/bot${process.env.TELEGRAM_TOKEN}/${fileInfo.file_path}`;

  // 2. Descargar imagen
  const imgRes = await axios.get(fileUrl, { responseType: "arraybuffer", timeout: 20000 });
  const b64    = Buffer.from(imgRes.data).toString("base64");

  // Detectar tipo de imagen
  const ext      = fileInfo.file_path.split(".").pop().toLowerCase();
  const mimeMap  = { jpg: "image/jpeg", jpeg: "image/jpeg", png: "image/png", webp: "image/webp" };
  const mimeType = mimeMap[ext] || "image/jpeg";

  console.log("Imagen descargada:", fileInfo.file_path, "tipo:", mimeType, "tamaño:", imgRes.data.byteLength);

  // 3. Llamar a Claude
  const resp = await axios.post(
    "https://api.anthropic.com/v1/messages",
    {
      model: "claude-haiku-4-5-20251001",
      max_tokens: 300,
      messages: [{
        role: "user",
        content: [
          {
            type: "image",
            source: { type: "base64", media_type: mimeType, data: b64 }
          },
          {
            type: "text",
            text: `Mira esta imagen de boleta o recibo. Extrae los datos y responde UNICAMENTE con este JSON (sin markdown, sin texto extra):
{"fecha":"DD/MM/YYYY","monto":numero,"moneda":"CLP","comercio":"nombre del local"}
Si algun dato no se ve, usa null. Si la imagen no es una boleta responde: {"error":"no es boleta"}`
          }
        ]
      }]
    },
    {
      headers: {
        "x-api-key":          process.env.ANTHROPIC_API_KEY,
        "anthropic-version":  "2023-06-01",
        "content-type":       "application/json",
      },
      timeout: 30000
    }
  );

  const texto = resp.data?.content?.[0]?.text || "";
  console.log("Respuesta Claude:", texto);
  const clean = texto.replace(/```json|```/g, "").trim();
  try { return JSON.parse(clean); } catch { return { error: "parse", raw: texto }; }
}

// ── Teclados ─────────────────────────────────────────────
const KB = {
  cancelar:  { reply_markup: { keyboard: [["Cancelar"]], resize_keyboard: true } },
  confirmar: { reply_markup: { keyboard: [["Confirmar", "Cancelar"]], resize_keyboard: true, one_time_keyboard: true } },
  omitir:    { reply_markup: { keyboard: [["Omitir", "Cancelar"]], resize_keyboard: true, one_time_keyboard: true } },
  inicio:    { reply_markup: { keyboard: [["Ver planilla"]], resize_keyboard: true } },
};

// ── Función para pedir motivo (paso siguiente tras foto) ─
async function pedirMotivo(id, draft) {
  const mFmt = draft.monto != null
    ? `${draft.moneda || "CLP"} $${Number(draft.monto).toLocaleString("es-CL")}`
    : "no detectado";
  setPaso(id, "motivo", draft);
  await bot.sendMessage(id,
    `Datos detectados:\n\nFecha: ${draft.fecha || "no detectada"}\nMonto: ${mFmt}${draft.comercio ? `\nComercio: ${draft.comercio}` : ""}\n\nCual es el MOTIVO del gasto?\n(ej: Alimentacion, Transporte, Utiles, Representacion...)`,
    KB.cancelar
  );
}

// ── Manejador principal ───────────────────────────────────
bot.on("message", async (msg) => {
  const id  = msg.chat.id;
  const txt = (msg.text || "").trim();
  const cur = ses(id);

  // Comandos globales
  if (txt === "Cancelar" || txt === "/cancelar") {
    reset(id);
    return bot.sendMessage(id, "Cancelado. Envia una foto de boleta para empezar.", KB.inicio);
  }
  if (txt === "/planilla" || txt === "Ver planilla") {
    if (!fs.existsSync(EXCEL))
      return bot.sendMessage(id, "Aun no tienes gastos registrados.\nEnvia una foto de boleta para empezar.");
    await bot.sendMessage(id, "Enviando tu planilla Excel...");
    return bot.sendDocument(id, EXCEL, { caption: "Tu planilla de gastos actualizada" });
  }
  if (txt === "/start") {
    reset(id);
    return bot.sendMessage(id,
      "Hola! Soy tu bot de gastos.\n\n" +
      "Envia una FOTO de boleta o recibo y registrare el gasto automaticamente.\n\n" +
      "Comandos:\n/planilla - recibir el Excel\n/cancelar - cancelar",
      KB.inicio
    );
  }

  // ── Foto recibida ─────────────────────────────────────
  if (msg.photo || msg.document?.mime_type?.startsWith("image/")) {
    if (cur.paso !== "inicio") {
      return bot.sendMessage(id, "Escribe Cancelar primero si quieres empezar de nuevo.", KB.cancelar);
    }

    const msj = await bot.sendMessage(id, "Leyendo boleta con IA...");

    try {
      const fileId = msg.photo
        ? msg.photo[msg.photo.length - 1].file_id
        : msg.document.file_id;

      const datos = await analizarImagen(fileId);
      await bot.deleteMessage(id, msj.message_id).catch(() => {});

      if (datos.error) {
        reset(id);
        return bot.sendMessage(id,
          "No pude leer la boleta automaticamente.\n\nIngresemos los datos manualmente.\n\nCual es la FECHA?\n(ej: 13/06/2026 o escribe 'hoy')",
          KB.cancelar
        );
      }

      const corr  = siguienteCorr();
      const fecha = datos.fecha || new Date().toLocaleDateString("es-CL");
      await pedirMotivo(id, {
        corr, fecha,
        monto:    datos.monto,
        moneda:   datos.moneda || "CLP",
        comercio: datos.comercio || "",
      });

    } catch (err) {
      await bot.deleteMessage(id, msj.message_id).catch(() => {});
      console.error("Error analizando imagen:", err.response?.data || err.message);

      // Fallback: modo manual
      const corr = siguienteCorr();
      setPaso(id, "fecha", { corr });
      await bot.sendMessage(id,
        "No pude conectar con la IA ahora.\n\nIngresemos los datos manualmente.\n\nCual es la FECHA?\n(ej: 13/06/2026 o escribe 'hoy')",
        KB.cancelar
      );
    }
    return;
  }

  // ── Flujo manual (fecha → monto → comercio) ──────────
  if (cur.paso === "fecha") {
    const fecha = txt.toLowerCase() === "hoy"
      ? new Date().toLocaleDateString("es-CL") : txt;
    setPaso(id, "monto", { ...cur.d, fecha });
    return bot.sendMessage(id, "Cual es el MONTO?\n(solo numeros, ej: 9350)", KB.cancelar);
  }

  if (cur.paso === "monto") {
    const monto = Number(txt.replace(/\./g, "").replace(/,/g, "."));
    if (isNaN(monto) || monto <= 0)
      return bot.sendMessage(id, "Escribe solo el numero (ej: 9350)");
    setPaso(id, "comercio", { ...cur.d, monto, moneda: "CLP" });
    return bot.sendMessage(id, "Cual es el COMERCIO o lugar de compra?\n(o toca Omitir)", KB.omitir);
  }

  if (cur.paso === "comercio") {
    const comercio = txt === "Omitir" ? "" : txt;
    await pedirMotivo(id, { ...cur.d, comercio });
    return;
  }

  // ── Flujo común (motivo → destino → detalle → confirmar) ─
  if (cur.paso === "motivo") {
    if (!txt) return;
    setPaso(id, "destino", { ...cur.d, motivo: txt });
    return bot.sendMessage(id,
      "Cual es el DESTINO o area del gasto?\n(ej: Administrativo, Proyecto X, RRHH, Ventas...)",
      KB.cancelar
    );
  }

  if (cur.paso === "destino") {
    if (!txt) return;
    setPaso(id, "detalle", { ...cur.d, destino: txt });
    return bot.sendMessage(id, "Agrega un DETALLE o comentario adicional\n(o toca Omitir)", KB.omitir);
  }

  if (cur.paso === "detalle") {
    const d = { ...cur.d, detalle: txt === "Omitir" ? "" : txt };
    setPaso(id, "confirmar", d);
    const corr = String(d.corr).padStart(4, "0");
    const mFmt = d.monto != null
      ? `${d.moneda} $${Number(d.monto).toLocaleString("es-CL")}` : "-";
    return bot.sendMessage(id,
      `Resumen - Gasto #${corr}\n\n` +
      `Fecha:    ${d.fecha}\n` +
      `Monto:    ${mFmt}\n` +
      `Comercio: ${d.comercio || "-"}\n` +
      `Motivo:   ${d.motivo}\n` +
      `Destino:  ${d.destino}\n` +
      `Detalle:  ${d.detalle || "-"}\n\n` +
      `Confirmar y guardar en Excel?`,
      KB.confirmar
    );
  }

  if (cur.paso === "confirmar") {
    if (txt !== "Confirmar") return;
    try {
      const d     = cur.d;
      const total = await guardarExcel(d);
      avanzarCorr(d.corr);
      reset(id);
      return bot.sendMessage(id,
        `Gasto #${String(d.corr).padStart(4,"0")} guardado en la planilla!\n\n` +
        `La planilla tiene ${total} registro${total > 1 ? "s" : ""} en total.\n\n` +
        `Envia otra foto para seguir registrando.\n` +
        `Escribe /planilla para recibir el Excel.`,
        KB.inicio
      );
    } catch (err) {
      console.error("Error guardando:", err.message);
      return bot.sendMessage(id, "Error al guardar en Excel. Intenta de nuevo o escribe /cancelar.");
    }
  }

  // Sin contexto
  if (cur.paso === "inicio") {
    bot.sendMessage(id,
      "Envia una foto de boleta para registrar un gasto.\n/planilla - recibir Excel",
      KB.inicio
    );
  }
});

bot.on("polling_error", err => console.error("Polling error:", err.message));
console.log("Bot de gastos iniciado correctamente!");
