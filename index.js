require("dotenv").config();
const TelegramBot = require("node-telegram-bot-api");
const axios       = require("axios");
const crypto      = require("crypto");

const bot = new TelegramBot(process.env.TELEGRAM_TOKEN, { polling: true });

// ── Sesiones ─────────────────────────────────────────────
const S = {};
function ses(id) { if (!S[id]) S[id] = { paso: "inicio", d: {} }; return S[id]; }
function set(id, paso, d) { S[id] = { paso, d: d !== undefined ? d : (S[id]?.d || {}) }; }
function reset(id) { S[id] = { paso: "inicio", d: {} }; }

let corrActual = 1;

// ── Access Token para Google ──────────────────────────────
async function getAccessToken(scope) {
  const creds   = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const now     = Math.floor(Date.now() / 1000);
  const header  = Buffer.from(JSON.stringify({ alg: "RS256", typ: "JWT" })).toString("base64url");
  const payload = Buffer.from(JSON.stringify({
    iss: creds.client_email,
    scope,
    aud: "https://oauth2.googleapis.com/token",
    exp: now + 3600,
    iat: now,
  })).toString("base64url");

  const sign = crypto.createSign("RSA-SHA256");
  sign.update(`${header}.${payload}`);
  const sig = sign.sign(creds.private_key, "base64url");
  const jwt = `${header}.${payload}.${sig}`;

  const res = await axios.post("https://oauth2.googleapis.com/token", {
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    assertion:  jwt,
  });
  return res.data.access_token;
}

// ── Guardar en Google Sheets ──────────────────────────────
async function agregarFila(g, fotoUrl) {
  const token   = await getAccessToken("https://www.googleapis.com/auth/spreadsheets");
  const sheetId = process.env.GOOGLE_SHEET_ID;
  const corr    = `GASTO_${String(g.corr).padStart(4, "0")}`;
  const values  = [[
    corr,
    g.fecha      || "",
    g.monto      != null ? Number(g.monto) : "",
    g.moneda     || "CLP",
    g.comercio   || "",
    g.motivo     || "",
    g.destino    || "",
    g.detalle    || "",
    fotoUrl      || "",
    new Date().toLocaleString("es-CL"),
  ]];

  await axios.post(
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/Gastos!A:J:append?valueInputOption=USER_ENTERED`,
    { values },
    { headers: { Authorization: `Bearer ${token}` } }
  );
  return corr;
}

// ── Subir foto a Google Drive ─────────────────────────────
async function subirFotoDrive(buffer, nombre) {
  const token    = await getAccessToken("https://www.googleapis.com/auth/drive");
  const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;

  // Subir archivo
  const metadata = JSON.stringify({ name: nombre, parents: [folderId] });
  const boundary = "foo_bar_baz";
  const body     = Buffer.concat([
    Buffer.from(`--${boundary}\r\nContent-Type: application/json\r\n\r\n${metadata}\r\n--${boundary}\r\nContent-Type: image/jpeg\r\n\r\n`),
    buffer,
    Buffer.from(`\r\n--${boundary}--`),
  ]);

  const res = await axios.post(
    "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink",
    body,
    {
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": `multipart/related; boundary=${boundary}`,
      },
    }
  );

  const fileId = res.data.id;

  // Hacer público
  await axios.post(
    `https://www.googleapis.com/drive/v3/files/${fileId}/permissions`,
    { role: "reader", type: "anyone" },
    { headers: { Authorization: `Bearer ${token}` } }
  );

  return res.data.webViewLink;
}

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
    const sheetId = process.env.GOOGLE_SHEET_ID;
    const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;
    return bot.sendMessage(id,
      `Tu planilla de gastos:\nhttps://docs.google.com/spreadsheets/d/${sheetId}\n\nFotos de boletas:\nhttps://drive.google.com/drive/folders/${folderId}`,
      KB.inicio
    );
  }

  if (txt === "/start") {
    reset(id);
    return bot.sendMessage(id,
      "Hola! Soy tu bot de gastos.\n\nEnvia una FOTO de boleta y registrare el gasto en Google Sheets y guardare la foto en Google Drive automaticamente.\n\nComandos:\n/planilla - ver planilla y fotos\n/cancelar - cancelar",
      KB.inicio
    );
  }

  // ── Foto ──────────────────────────────────────────────
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

      if (datos.error) {
        set(id, "fecha", { corr: corrActual, buffer });
        return bot.sendMessage(id,
          "No pude leer los datos automaticamente.\n\nIngresemos manualmente.\n\nCual es la FECHA?\n(ej: 13/06/2026 o escribe 'hoy')",
          KB.cancelar
        );
      }

      await pedirMotivo(id, {
        corr:     corrActual,
        buffer,
        fecha:    datos.fecha || new Date().toLocaleDateString("es-CL"),
        monto:    datos.monto,
        moneda:   datos.moneda || "CLP",
        comercio: datos.comercio || "",
      });

    } catch (err) {
      await bot.deleteMessage(id, msj.message_id).catch(() => {});
      console.error("Error:", err.response?.data || err.message);
      set(id, "fecha", { corr: corrActual });
      await bot.sendMessage(id,
        "Error con la IA. Ingresemos manualmente.\n\nCual es la FECHA?\n(ej: 13/06/2026 o 'hoy')",
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
    if (isNaN(monto) || monto <= 0) return bot.sendMessage(id, "Escribe solo el numero (ej: 9350)");
    set(id, "comercio", { ...cur.d, monto, moneda: "CLP" });
    return bot.sendMessage(id, "Cual es el COMERCIO? (o Omitir)", KB.omitir);
  }

  if (cur.paso === "comercio") {
    await pedirMotivo(id, { ...cur.d, comercio: txt === "Omitir" ? "" : txt });
    return;
  }

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
      `Detalle:  ${d.detalle || "-"}\n\n` +
      `Confirmar y guardar?`,
      KB.confirmar
    );
  }

  if (cur.paso === "confirmar" && txt === "Confirmar") {
    try {
      const d    = cur.d;
      const corr = String(d.corr).padStart(4, "0");
      const nombre = `GASTO_${corr}.jpg`;

      // Subir foto a Drive
      let fotoUrl = "";
      try {
        fotoUrl = await subirFotoDrive(d.buffer, nombre);
        console.log("Foto subida:", nombre);
      } catch (e) {
        console.error("Error subiendo foto:", e.message);
      }

      // Guardar en Sheets
      const corrStr = await agregarFila(d, fotoUrl);
      corrActual++;
      reset(id);

      const sheetId  = process.env.GOOGLE_SHEET_ID;
      return bot.sendMessage(id,
        `${corrStr} guardado!\n\n` +
        `Foto: ${nombre}${fotoUrl ? `\n${fotoUrl}` : ""}\n\n` +
        `Ver planilla:\nhttps://docs.google.com/spreadsheets/d/${sheetId}\n\n` +
        `Envia otra foto para seguir registrando.`,
        KB.inicio
      );
    } catch (err) {
      console.error("Error guardando:", err.response?.data || err.message);
      return bot.sendMessage(id, "Error al guardar. Intenta de nuevo o escribe /cancelar.");
    }
  }

  if (cur.paso === "inicio") {
    bot.sendMessage(id, "Envia una foto de boleta.\n/planilla - ver Google Sheets y Drive", KB.inicio);
  }
});

bot.on("polling_error", err => console.error("Polling error:", err.message));
console.log("Bot iniciado con Google Sheets + Drive!");
