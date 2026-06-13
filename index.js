require("dotenv").config();
const TelegramBot = require("node-telegram-bot-api");
const Anthropic   = require("@anthropic-ai/sdk");
const ExcelJS     = require("exceljs");
const axios       = require("axios");
const fs          = require("fs");
const path        = require("path");

// ── Configuración ────────────────────────────────────────
const bot    = new TelegramBot(process.env.TELEGRAM_TOKEN, { polling: true });
const claude = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const DATOS_DIR    = path.resolve("./datos");
const EXCEL_PATH   = path.join(DATOS_DIR, "Registro_Gastos.xlsx");
const COUNTER_PATH = path.join(DATOS_DIR, "correlativo.json");

fs.mkdirSync(DATOS_DIR, { recursive: true });

// ── Estado de conversación por usuario ──────────────────
const sesiones = {};

function sesion(chatId) {
  if (!sesiones[chatId]) sesiones[chatId] = { paso: "inicio", draft: {} };
  return sesiones[chatId];
}
function setPaso(chatId, paso, draft = null) {
  sesiones[chatId] = { paso, draft: draft ?? sesiones[chatId]?.draft ?? {} };
}
function resetear(chatId) {
  sesiones[chatId] = { paso: "inicio", draft: {} };
}

// ── Correlativo persistente ──────────────────────────────
function siguienteCorr() {
  try { return JSON.parse(fs.readFileSync(COUNTER_PATH)).next || 1; } catch { return 1; }
}
function avanzarCorr(n) {
  fs.writeFileSync(COUNTER_PATH, JSON.stringify({ next: n + 1 }));
}

// ── Excel acumulativo ────────────────────────────────────
async function agregarGasto(gasto) {
  const wb = new ExcelJS.Workbook();
  const HEADER = "FF128C7E";

  if (fs.existsSync(EXCEL_PATH)) {
    await wb.xlsx.readFile(EXCEL_PATH);
  } else {
    // Crear libro nuevo con estructura
    const ws = wb.addWorksheet("Gastos");
    ws.columns = [
      { header: "Correlativo", key: "correlativo", width: 13 },
      { header: "Fecha",       key: "fecha",        width: 12 },
      { header: "Monto",       key: "monto",        width: 14 },
      { header: "Moneda",      key: "moneda",        width: 8  },
      { header: "Comercio",    key: "comercio",      width: 22 },
      { header: "Descripción", key: "descripcion",   width: 32 },
      { header: "Motivo",      key: "motivo",        width: 20 },
      { header: "Destino",     key: "destino",       width: 20 },
      { header: "Detalle",     key: "detalle",       width: 32 },
    ];
    ws.getRow(1).eachCell(cell => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEADER } };
      cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
      cell.alignment = { horizontal: "center", vertical: "middle" };
    });
    ws.getRow(1).height = 22;
    wb.addWorksheet("Por Motivo");
    wb.addWorksheet("Por Destino");
    await wb.xlsx.writeFile(EXCEL_PATH);
    await wb.xlsx.readFile(EXCEL_PATH);
  }

  const ws  = wb.getWorksheet("Gastos");
  const num = ws.lastRow ? ws.lastRow.number : 1;

  const fila = ws.addRow({
    correlativo: `GASTO_${String(gasto.correlativo).padStart(4, "0")}`,
    fecha:       gasto.fecha,
    monto:       gasto.monto,
    moneda:      gasto.moneda || "CLP",
    comercio:    gasto.comercio || "",
    descripcion: gasto.descripcion || "",
    motivo:      gasto.motivo || "",
    destino:     gasto.destino || "",
    detalle:     gasto.detalle || "",
  });
  fila.getCell("monto").numFmt = "#,##0";
  if (fila.number % 2 === 0) {
    fila.eachCell({ includeEmpty: true }, cell => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F8F6" } };
    });
  }

  // Resúmenes
  for (const [hoja, campo] of [["Por Motivo","motivo"],["Por Destino","destino"]]) {
    let wsr = wb.getWorksheet(hoja);
    if (!wsr) wsr = wb.addWorksheet(hoja);
    wsr.spliceRows(1, wsr.lastRow?.number || 0);
    wsr.columns = [{ header: campo === "motivo" ? "Motivo" : "Destino", key: "cat", width: 26 }, { header: "Total", key: "total", width: 16 }];
    wsr.getRow(1).eachCell(cell => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEADER } };
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    });
    const mapa = {};
    ws.eachRow((r, i) => {
      if (i === 1) return;
      const k = r.getCell(campo).value || "Sin clasificar";
      mapa[k] = (mapa[k] || 0) + (Number(r.getCell("monto").value) || 0);
    });
    Object.entries(mapa).sort((a,b) => b[1]-a[1]).forEach(([k,v]) => {
      const r = wsr.addRow({ cat: k, total: v });
      r.getCell("total").numFmt = "#,##0";
    });
  }

  await wb.xlsx.writeFile(EXCEL_PATH);
  return num; // total registros
}

// ── Analizar imagen con Claude ───────────────────────────
async function analizarBoleta(fileUrl) {
  const res  = await axios.get(fileUrl, { responseType: "arraybuffer" });
  const b64  = Buffer.from(res.data).toString("base64");
  const mime = "image/jpeg";

  const result = await claude.messages.create({
    model: "claude-sonnet-4-6",
    max_tokens: 800,
    messages: [{
      role: "user",
      content: [
        { type: "image", source: { type: "base64", media_type: mime, data: b64 } },
        { type: "text", text: `Analiza esta imagen. Si es una boleta, recibo o comprobante de pago responde SOLO con JSON sin markdown:
{"fecha":"DD/MM/YYYY","monto":número,"moneda":"CLP","comercio":"nombre","descripcion":"qué se compró en máx 15 palabras"}
Usa null si no ves el dato. Si NO es un comprobante responde: {"error":"no es boleta"}` }
      ]
    }]
  });

  const raw = result.content.map(c => c.text||"").join("").replace(/```json|```/g,"").trim();
  try { return JSON.parse(raw); } catch { return { error: "parse" }; }
}

// ── Teclado de opciones ──────────────────────────────────
const btnCancelar = { reply_markup: { keyboard: [["❌ Cancelar"]], resize_keyboard: true } };
const btnConfirmar = {
  reply_markup: {
    keyboard: [["✅ Confirmar"], ["❌ Cancelar"]],
    resize_keyboard: true,
    one_time_keyboard: true
  }
};
const btnInicio = { reply_markup: { keyboard: [["📊 Ver planilla"]], resize_keyboard: true } };

// ── Manejador principal ──────────────────────────────────
bot.on("message", async (msg) => {
  const chatId = msg.chat.id;
  const texto  = msg.text?.trim() || "";
  const s      = sesion(chatId);

  // Cancelar en cualquier momento
  if (texto === "❌ Cancelar" || texto.toLowerCase() === "/cancelar") {
    resetear(chatId);
    return bot.sendMessage(chatId, "❌ Operación cancelada.\n\nEnvíame una foto de boleta para empezar.", btnInicio);
  }

  // Comando /start
  if (texto === "/start") {
    resetear(chatId);
    return bot.sendMessage(chatId,
      "👋 ¡Hola! Soy tu *bot de gastos*.\n\n" +
      "📷 Envíame una foto de tu boleta o recibo y lo registraré automáticamente.\n\n" +
      "Comandos:\n• /planilla — recibir el Excel actualizado\n• /cancelar — cancelar operación",
      { parse_mode: "Markdown", ...btnInicio }
    );
  }

  // Pedir planilla
  if (texto === "/planilla" || texto === "📊 Ver planilla") {
    if (!fs.existsSync(EXCEL_PATH)) {
      return bot.sendMessage(chatId, "📭 Aún no tienes gastos registrados.\n\nEnvíame una foto de boleta para empezar.");
    }
    await bot.sendMessage(chatId, "📤 Enviando tu planilla...");
    return bot.sendDocument(chatId, EXCEL_PATH, { caption: "📊 Tu planilla de gastos actualizada" });
  }

  // ── Recibir foto ──
  if (msg.photo || msg.document?.mime_type?.startsWith("image/")) {
    if (s.paso !== "inicio" && s.paso !== "esperando_foto") {
      return bot.sendMessage(chatId, "⚠️ Estoy esperando tu respuesta. Escribe *❌ Cancelar* si quieres empezar de nuevo.", { parse_mode: "Markdown" });
    }

    setPaso(chatId, "procesando");
    const procesando = await bot.sendMessage(chatId, "🔍 Analizando tu boleta con IA...");

    try {
      // Obtener URL de la foto (mayor resolución)
      const fileId  = msg.photo ? msg.photo[msg.photo.length - 1].file_id : msg.document.file_id;
      const file    = await bot.getFile(fileId);
      const fileUrl = `https://api.telegram.org/file/bot${process.env.TELEGRAM_TOKEN}/${file.file_path}`;

      const datos = await analizarBoleta(fileUrl);
      await bot.deleteMessage(chatId, procesando.message_id).catch(() => {});

      if (datos.error) {
        setPaso(chatId, "inicio");
        return bot.sendMessage(chatId, "⚠️ No pude detectar una boleta en esa imagen.\nEnvía una foto más clara del comprobante.", btnInicio);
      }

      const corr     = siguienteCorr();
      const fecha    = datos.fecha || new Date().toLocaleDateString("es-CL");
      const montoFmt = datos.monto != null
        ? `${datos.moneda || "CLP"} $${Number(datos.monto).toLocaleString("es-CL")}`
        : "no detectado";

      setPaso(chatId, "motivo", {
        correlativo: corr, fecha,
        monto: datos.monto, moneda: datos.moneda || "CLP",
        comercio: datos.comercio || "", descripcion: datos.descripcion || "",
      });

      await bot.sendMessage(chatId,
        `✅ *Datos detectados:*\n\n` +
        `📅 Fecha: ${fecha}\n` +
        `💵 Monto: ${montoFmt}` +
        `${datos.comercio ? `\n🏪 Comercio: ${datos.comercio}` : ""}` +
        `${datos.descripcion ? `\n🧾 ${datos.descripcion}` : ""}\n\n` +
        `¿Cuál es el *motivo* del gasto?\n_(ej: Alimentación, Transporte, Útiles...)_`,
        { parse_mode: "Markdown", ...btnCancelar }
      );

    } catch (err) {
      console.error(err);
      setPaso(chatId, "inicio");
      bot.sendMessage(chatId, "❌ Error al procesar la imagen. Intenta de nuevo.", btnInicio);
    }
    return;
  }

  // ── Flujo de preguntas ──
  if (s.paso === "motivo") {
    if (!texto) return;
    setPaso(chatId, "destino", { ...s.draft, motivo: texto });
    return bot.sendMessage(chatId,
      `¿Cuál es el *destino* o área?\n_(ej: Administrativo, Proyecto X, RRHH, Ventas...)_`,
      { parse_mode: "Markdown", ...btnCancelar }
    );
  }

  if (s.paso === "destino") {
    if (!texto) return;
    setPaso(chatId, "detalle", { ...s.draft, destino: texto });
    return bot.sendMessage(chatId,
      `Agrega un *detalle* adicional o toca _Omitir_:`,
      {
        parse_mode: "Markdown",
        reply_markup: {
          keyboard: [["Omitir"], ["❌ Cancelar"]],
          resize_keyboard: true, one_time_keyboard: true
        }
      }
    );
  }

  if (s.paso === "detalle") {
    if (!texto) return;
    const detalle = texto === "Omitir" ? "" : texto;
    const d = { ...s.draft, detalle };
    setPaso(chatId, "confirmar", d);

    const corr     = String(d.correlativo).padStart(4, "0");
    const montoFmt = d.monto != null ? `${d.moneda} $${Number(d.monto).toLocaleString("es-CL")}` : "—";

    return bot.sendMessage(chatId,
      `📋 *Resumen — Gasto #${corr}*\n\n` +
      `📅 Fecha: ${d.fecha}\n` +
      `💵 Monto: ${montoFmt}\n` +
      `🏪 Comercio: ${d.comercio || "—"}\n` +
      `🧾 Descripción: ${d.descripcion || "—"}\n` +
      `🏷️ Motivo: ${d.motivo}\n` +
      `📍 Destino: ${d.destino}\n` +
      `📝 Detalle: ${d.detalle || "—"}\n\n` +
      `¿Confirmar y guardar?`,
      { parse_mode: "Markdown", ...btnConfirmar }
    );
  }

  if (s.paso === "confirmar") {
    if (texto === "✅ Confirmar") {
      try {
        const d     = s.draft;
        const total = await agregarGasto(d);
        avanzarCorr(d.correlativo);
        const corr  = String(d.correlativo).padStart(4, "0");
        resetear(chatId);
        return bot.sendMessage(chatId,
          `✅ *Gasto #${corr} guardado.*\n\n` +
          `📊 Planilla con *${total} registro${total > 1 ? "s" : ""}* en total.\n\n` +
          `Envía otra foto o escribe /planilla para recibir el Excel.`,
          { parse_mode: "Markdown", ...btnInicio }
        );
      } catch (err) {
        console.error(err);
        return bot.sendMessage(chatId, "❌ Error al guardar. Intenta de nuevo.");
      }
    }
    return;
  }

  // Mensaje de texto sin contexto
  if (s.paso === "inicio") {
    bot.sendMessage(chatId,
      "📷 Envíame una *foto de tu boleta* para registrar un gasto.\n\nO escribe /planilla para recibir el Excel.",
      { parse_mode: "Markdown", ...btnInicio }
    );
  }
});

bot.on("polling_error", err => console.error("Polling error:", err));
console.log("🚀 Bot de gastos iniciado");
