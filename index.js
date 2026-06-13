require("dotenv").config();
const TelegramBot = require("node-telegram-bot-api");
const axios       = require("axios");
const fs          = require("fs");
const { guardarGasto, siguienteCorr, avanzarCorr, EXCEL } = require("./excel");
const { subirFoto } = require("./drive");

const bot = new TelegramBot(process.env.TELEGRAM_TOKEN, { polling: true });

// ── Sesiones ─────────────────────────────────────────────
const S = {};
function ses(id) { if (!S[id]) S[id] = { paso: "inicio", d: {} }; return S[id]; }
function set(id, paso, d) { S[id] = { paso, d: d !== undefined ? d : (S[id]?.d || {}) }; }
function reset(id) { S[id] = { paso: "inicio", d: {} }; }

// ── Analizar imagen con Claude ────────────────────────────
async function analizarImagen(fileId) {
  const fileInfo = await bot.getFile(fileId);
  const fileUrl  = `https://api.telegram.org/file/bot${process.env.TELEGRAM_TOKEN}/${fileInfo.file_path}`;
  const imgRes   = await axios.get(fileUrl, { responseType: "arraybuffer", timeout: 20000 });
  const buffer   = Buffer.from(imgRes.data);
  const b64      = buffer.toString("base64");
  const ext      = fileInfo.file_path.split(".").pop().toLowerCase();
  const mime     = { jpg:"image/jpeg", jpeg:"image/jpeg", png:"image/png", webp:"image/webp" }[ext] || "image/jpeg";

  const resp = await axios.post(
    "https://api.anthropic.com/v1/messages",
    {
      model: "claude-haiku-4-5-20251001",
      max_tokens: 300,
      messages: [{
        role: "user",
        content: [
          { type: "image", source: { type: "base64", media_type: mime, data: b64 } },
          { type: "text", text: `Analiza esta boleta o recibo. Responde SOLO con JSON sin markdown:
{"fecha":"DD/MM/YYYY","monto":numero,"moneda":"CLP","comercio":"nombre"}
Usa null si no ves el dato. Si no es boleta: {"error":"no es boleta"}` }
        ]
      }]
    },
    {
      headers: {
        "x-api-key": process.env.ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
      },
      timeout: 30000
    }
  );

  const texto = resp.data?.content?.[0]?.text || "";
  const clean = texto.replace(/```json|```/g, "").trim();
  try { return { datos: JSON.parse(clean), buffer }; }
  catch { return { datos: { error: "parse" }, buffer }; }
}

// ── Teclados ─────────────────────────────────────────────
const KB = {
  cancelar:  { reply_markup: { keyboard: [["Cancelar"]], resize_keyboard: true } },
  confirmar: { reply_markup: { keyboard: [["Confirmar", "Cancelar"]], resize_keyboard: true, one_time_keyboard: true } },
  omitir:    { reply_markup: { keyboard: [["Omitir", "Cancelar"]], resize_keyboard: true, one_time_keyboard: true } },
  inicio:    { reply_markup: { keyboard: [["Ver planilla"]], resize_keyboard: true } },
};

async function pedirMotivo(id, draft) {
  const mFmt = draft.monto != null
    ? `${draft.moneda || "CLP"} $${Number(draft.monto).toLocaleString("es-CL")}`
    : "no detectado";
  set(id, "motivo", draft);
  await bot.sendMessage(id,
    `Datos detectados:\n\nFecha: ${draft.fecha || "no detectada"}\nMonto: ${mFmt}${draft.comercio ? `\nComercio: ${draft.comercio}` : ""}\n\nCual es el MOTIVO del gasto?\n(ej: Alimentacion, Transporte, Utiles...)`,
    KB.cancelar
  );
}

// ── Manejador ─────────────────────────────────────────────
bot.on("message", async (msg) => {
  const id  = msg.chat.id;
  const txt = (msg.text || "").trim();
  const cur = ses(id);

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
    if (cur.paso !== "inicio")
      return bot.sendMessage(id, "Escribe Cancelar primero.", KB.cancelar);

    const msj = await bot.sendMessage(id, "Leyendo boleta con IA...");

    try {
      const fileId = msg.photo
        ? msg.photo[msg.photo.length - 1].file_id
        : msg.document.file_id;

      const { datos, buffer } = await analizarImagen(fileId);
      await bot.deleteMessage(id, msj.message_id).catch(() => {});

      const corr       = siguienteCorr();
      const fotoNombre = `GASTO_${String(corr).padStart(4, "0")}.jpg`;

      // Subir foto a Google Drive
      let fotoUrl = "";
      try {
        const foto = await subirFoto(buffer, fotoNombre);
        fotoUrl = foto.url;
        console.log("Foto subida a Drive:", fotoNombre);
      } catch (e) {
        console.error("Error subiendo foto:", e.message);
      }

      if (datos.error) {
        // No es boleta — pedir datos manualmente
        set(id, "fecha", { corr, fotoNombre, fotoUrl });
        return bot.sendMessage(id,
          "No pude leer los datos automaticamente.\n\nIngresemos manualmente.\n\nCual es la FECHA?\n(ej: 13/06/2026 o escribe 'hoy')",
          KB.cancelar
        );
      }

      await pedirMotivo(id, {
        corr,
        fotoNombre,
        fotoUrl,
        fecha:    datos.fecha || new Date().toLocaleDateString("es-CL"),
        monto:    datos.monto,
        moneda:   datos.moneda || "CLP",
        comercio: datos.comercio || "",
      });

    } catch (err) {
      await bot.deleteMessage(id, msj.message_id).catch(() => {});
      console.error("Error:", err.response?.data || err.message);

      const corr = siguienteCorr();
      set(id, "fecha", { corr });
      await bot.sendMessage(id,
        "Error conectando con la IA. Ingresemos los datos manualmente.\n\nCual es la FECHA?\n(ej: 13/06/2026 o 'hoy')",
        KB.cancelar
      );
    }
    return;
  }

  // ── Flujo manual ──────────────────────────────────────
  if (cur.paso === "fecha") {
    const fecha = txt.toLowerCase() === "hoy" ? new Date().toLocaleDateString("es-CL") : txt;
    set(id, "monto", { ...cur.d, fecha });
    return bot.sendMessage(id, "Cual es el MONTO? (ej: 9350)", KB.cancelar);
  }

  if (cur.paso === "monto") {
    const monto = Number(txt.replace(/\./g, "").replace(/,/g, "."));
    if (isNaN(monto) || monto <= 0)
      return bot.sendMessage(id, "Escribe solo el numero (ej: 9350)");
    set(id, "comercio", { ...cur.d, monto, moneda: "CLP" });
    return bot.sendMessage(id, "Cual es el COMERCIO? (o Omitir)", KB.omitir);
  }

  if (cur.paso === "comercio") {
    await pedirMotivo(id, { ...cur.d, comercio: txt === "Omitir" ? "" : txt });
    return;
  }

  // ── Flujo común ───────────────────────────────────────
  if (cur.paso === "motivo") {
    if (!txt) return;
    set(id, "destino", { ...cur.d, motivo: txt });
    return bot.sendMessage(id, "Cual es el DESTINO o area?\n(ej: Administrativo, Proyecto X, RRHH...)", KB.cancelar);
  }

  if (cur.paso === "destino") {
    if (!txt) return;
    set(id, "detalle", { ...cur.d, destino: txt });
    return bot.sendMessage(id, "Agrega un DETALLE adicional (o Omitir)", KB.omitir);
  }

  if (cur.paso === "detalle") {
    const d = { ...cur.d, detalle: txt === "Omitir" ? "" : txt };
    set(id, "confirmar", d);
    const corr = String(d.corr).padStart(4, "0");
    const mFmt = d.monto != null ? `${d.moneda} $${Number(d.monto).toLocaleString("es-CL")}` : "-";
    return bot.sendMessage(id,
      `Resumen - Gasto #${corr}\n\n` +
      `Fecha:    ${d.fecha}\n` +
      `Monto:    ${mFmt}\n` +
      `Comercio: ${d.comercio || "-"}\n` +
      `Motivo:   ${d.motivo}\n` +
      `Destino:  ${d.destino}\n` +
      `Detalle:  ${d.detalle || "-"}\n` +
      `Foto:     ${d.fotoNombre || "-"}\n\n` +
      `Confirmar y guardar?`,
      KB.confirmar
    );
  }

  if (cur.paso === "confirmar" && txt === "Confirmar") {
    try {
      const d     = cur.d;
      const total = await guardarGasto(d);
      avanzarCorr(d.corr);
      reset(id);
      return bot.sendMessage(id,
        `Gasto #${String(d.corr).padStart(4,"0")} guardado!\n\n` +
        `Planilla con ${total} registro${total > 1 ? "s" : ""} en total.\n` +
        (d.fotoUrl ? `Foto en Drive: ${d.fotoUrl}\n` : "") +
        `\nEnvia otra foto o escribe /planilla para el Excel.`,
        KB.inicio
      );
    } catch (err) {
      console.error("Error guardando:", err.message);
      return bot.sendMessage(id, "Error al guardar. Intenta de nuevo o escribe /cancelar.");
    }
  }

  if (cur.paso === "inicio") {
    bot.sendMessage(id, "Envia una foto de boleta para registrar un gasto.\n/planilla - recibir Excel", KB.inicio);
  }
});

bot.on("polling_error", err => console.error("Polling error:", err.message));
console.log("Bot de gastos iniciado!");
