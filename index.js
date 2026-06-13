require("dotenv").config();
const TelegramBot = require("node-telegram-bot-api");
const axios       = require("axios");
const crypto      = require("crypto");
const FormData    = require("form-data");

const bot = new TelegramBot(process.env.TELEGRAM_TOKEN, { polling: true });

// ── Usuarios autorizados ──────────────────────────────────
// Para agregar/cambiar usuarios edita este objeto
// "CODIGO": "Nombre"
const USUARIOS = {
  "2691": "Francisco",
  "1430": "Nicolas",
  // Agrega más usuarios aquí:
  // "CODIGO3": "Nombre3",
  // "CODIGO4": "Nombre4",
  // "CODIGO5": "Nombre5",
  // "CODIGO6": "Nombre6",
  // "CODIGO7": "Nombre7",
  // "CODIGO8": "Nombre8",
};

// ── Sesiones ─────────────────────────────────────────────
const S = {};
function ses(id) { if (!S[id]) S[id] = { paso: "login", d: {}, usuario: null, corr: 1 }; return S[id]; }
function set(id, paso, d) { S[id] = { ...ses(id), paso, d: d !== undefined ? d : (S[id]?.d || {}) }; }
function reset(id) { const u = ses(id); S[id] = { paso: "inicio", d: {}, usuario: u.usuario, corr: u.corr }; }

// ── Cloudinary ────────────────────────────────────────────
async function subirCloudinary(buffer, nombre) {
  const form = new FormData();
  form.append("file", buffer, { filename: nombre, contentType: "image/jpeg" });
  form.append("upload_preset", process.env.CLOUDINARY_UPLOAD_PRESET);
  form.append("public_id", nombre.replace(".jpg", ""));
  const res = await axios.post(
    `https://api.cloudinary.com/v1_1/${process.env.CLOUDINARY_CLOUD_NAME}/image/upload`,
    form, { headers: form.getHeaders(), timeout: 30000 }
  );
  return res.data.secure_url;
}

// ── Google Sheets ─────────────────────────────────────────
async function getToken() {
  const creds = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const now   = Math.floor(Date.now() / 1000);
  const hdr   = Buffer.from(JSON.stringify({ alg: "RS256", typ: "JWT" })).toString("base64url");
  const pay   = Buffer.from(JSON.stringify({
    iss: creds.client_email,
    scope: "https://www.googleapis.com/auth/spreadsheets",
    aud: "https://oauth2.googleapis.com/token",
    exp: now + 3600, iat: now,
  })).toString("base64url");
  const sign = crypto.createSign("RSA-SHA256");
  sign.update(`${hdr}.${pay}`);
  const sig = sign.sign(creds.private_key, "base64url");
  const res = await axios.post("https://oauth2.googleapis.com/token", {
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    assertion: `${hdr}.${pay}.${sig}`,
  });
  return res.data.access_token;
}

async function asegurarHoja(token, sheetId, nombreHoja) {
  // Obtener hojas existentes
  const res = await axios.get(
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=sheets.properties.title`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const hojas = res.data.sheets.map(h => h.properties.title);

  if (!hojas.includes(nombreHoja)) {
    // Crear hoja nueva
    await axios.post(
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}:batchUpdate`,
      { requests: [{ addSheet: { properties: { title: nombreHoja } } }] },
      { headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" } }
    );
    // Agregar cabeceras
    await axios.put(
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${nombreHoja}!A1:G1?valueInputOption=USER_ENTERED`,
      { values: [["Correlativo", "Fecha", "Motivo", "Destino", "Detalle", "Monto", "URL Foto"]] },
      { headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" } }
    );
  }
}

async function agregarFila(g, fotoUrl, nombreUsuario) {
  const token   = await getToken();
  const sheetId = process.env.GOOGLE_SHEET_ID;
  const hoja    = nombreUsuario; // Una hoja por usuario

  await asegurarHoja(token, sheetId, hoja);

  const corr   = `GASTO_${String(g.corr).padStart(4, "0")}`;
  const values = [[
    corr,
    g.fecha    || "",
    g.motivo   || "",
    g.destino  || "",
    g.detalle  || "",
    g.monto    != null ? Number(g.monto) : "",
    fotoUrl    || "",
  ]];

  await axios.post(
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${hoja}!A:G:append?valueInputOption=USER_ENTERED`,
    { values },
    { headers: { Authorization: `Bearer ${token}` } }
  );
  return corr;
}

// ── Claude ────────────────────────────────────────────────
async function analizarImagen(fileId) {
  const fi   = await bot.getFile(fileId);
  const url  = `https://api.telegram.org/file/bot${process.env.TELEGRAM_TOKEN}/${fi.file_path}`;
  const ir   = await axios.get(url, { responseType: "arraybuffer", timeout: 20000 });
  const buf  = Buffer.from(ir.data);
  const ext  = fi.file_path.split(".").pop().toLowerCase();
  const mime = { jpg:"image/jpeg", jpeg:"image/jpeg", png:"image/png", webp:"image/webp" }[ext] || "image/jpeg";

  const resp = await axios.post(
    "https://api.anthropic.com/v1/messages",
    { model:"claude-haiku-4-5-20251001", max_tokens:300, messages:[{ role:"user", content:[
      { type:"image", source:{ type:"base64", media_type:mime, data:buf.toString("base64") } },
      { type:"text", text:`Analiza esta boleta. Responde SOLO con JSON sin markdown:
{"fecha":"DD/MM/YYYY","monto":numero,"moneda":"CLP","comercio":"nombre"}
Usa null si no ves el dato. Si no es boleta: {"error":"no es boleta"}` }
    ]}] },
    { headers:{ "x-api-key":process.env.ANTHROPIC_API_KEY, "anthropic-version":"2023-06-01", "content-type":"application/json" }, timeout:30000 }
  );
  const texto = resp.data?.content?.[0]?.text || "";
  try { return { datos: JSON.parse(texto.replace(/```json|```/g,"").trim()), buffer: buf }; }
  catch { return { datos: { error:"parse" }, buffer: buf }; }
}

// ── Teclados ─────────────────────────────────────────────
const KB = {
  cancelar:  { reply_markup: { keyboard: [["Cancelar"]], resize_keyboard: true } },
  confirmar: { reply_markup: { keyboard: [["Confirmar","Cancelar"]], resize_keyboard: true, one_time_keyboard: true } },
  omitir:    { reply_markup: { keyboard: [["Omitir","Cancelar"]], resize_keyboard: true, one_time_keyboard: true } },
  inicio:    { reply_markup: { keyboard: [["Ver planilla"]], resize_keyboard: true } },
};

async function pedirMotivo(id, draft) {
  const mFmt = draft.monto != null ? `${draft.moneda||"CLP"} $${Number(draft.monto).toLocaleString("es-CL")}` : "no detectado";
  set(id, "motivo", draft);
  await bot.sendMessage(id,
    `Datos detectados:\n\nFecha: ${draft.fecha||"no detectada"}\nMonto: ${mFmt}${draft.comercio?`\nComercio: ${draft.comercio}`:""}\n\nCual es el MOTIVO del gasto?\n(ej: Alimentacion, Transporte, Utiles...)`,
    KB.cancelar
  );
}

// ── Manejador ─────────────────────────────────────────────
bot.on("message", async (msg) => {
  const id  = msg.chat.id;
  const txt = (msg.text || "").trim();
  const cur = ses(id);

  // ── Login ────────────────────────────────────────────
  if (!cur.usuario) {
    if (txt === "/start") {
      return bot.sendMessage(id,
        "Bienvenido al Bot de Gastos.\n\nIngresa tu codigo de acceso para continuar:",
        { reply_markup: { remove_keyboard: true } }
      );
    }
    // Verificar código
    if (USUARIOS[txt]) {
      S[id] = { paso: "inicio", d: {}, usuario: USUARIOS[txt], corr: 1 };
      return bot.sendMessage(id,
        `Bienvenido ${USUARIOS[txt]}!\n\nEnvia una FOTO de boleta para registrar un gasto.\n\n/planilla - ver tu planilla\n/salir - cerrar sesion`,
        KB.inicio
      );
    } else {
      return bot.sendMessage(id, "Codigo incorrecto. Intenta nuevamente:");
    }
  }

  // ── Comandos globales ────────────────────────────────
  if (txt === "/salir") {
    S[id] = { paso: "login", d: {}, usuario: null, corr: 1 };
    return bot.sendMessage(id, "Sesion cerrada. Escribe /start para volver a entrar.", { reply_markup: { remove_keyboard: true } });
  }

  if (txt === "Cancelar" || txt === "/cancelar") {
    reset(id);
    return bot.sendMessage(id, "Cancelado.", KB.inicio);
  }

  if (txt === "/planilla" || txt === "Ver planilla") {
    return bot.sendMessage(id,
      `Tu planilla (hoja: ${cur.usuario}):\nhttps://docs.google.com/spreadsheets/d/${process.env.GOOGLE_SHEET_ID}`,
      KB.inicio
    );
  }

  // ── Foto ─────────────────────────────────────────────
  if (msg.photo || msg.document?.mime_type?.startsWith("image/")) {
    if (cur.paso !== "inicio") return bot.sendMessage(id, "Escribe Cancelar primero.", KB.cancelar);
    const msj = await bot.sendMessage(id, "Leyendo boleta con IA...");
    try {
      const fid = msg.photo ? msg.photo[msg.photo.length-1].file_id : msg.document.file_id;
      const { datos, buffer } = await analizarImagen(fid);
      await bot.deleteMessage(id, msj.message_id).catch(()=>{});
      if (datos.error) {
        set(id, "fecha", { corr: cur.corr, buffer });
        return bot.sendMessage(id, "No pude leer los datos.\n\nCual es la FECHA?\n(ej: 13/06/2026 o 'hoy')", KB.cancelar);
      }
      await pedirMotivo(id, { corr:cur.corr, buffer, fecha:datos.fecha||new Date().toLocaleDateString("es-CL"), monto:datos.monto, moneda:datos.moneda||"CLP", comercio:datos.comercio||"" });
    } catch(err) {
      await bot.deleteMessage(id, msj.message_id).catch(()=>{});
      console.error("Error IA:", err.response?.data||err.message);
      set(id, "fecha", { corr: cur.corr });
      await bot.sendMessage(id, "Error con la IA. Ingresemos manualmente.\n\nCual es la FECHA?\n(ej: 13/06/2026 o 'hoy')", KB.cancelar);
    }
    return;
  }

  // ── Flujo manual ─────────────────────────────────────
  if (cur.paso==="fecha") {
    const f = txt.toLowerCase()==="hoy" ? new Date().toLocaleDateString("es-CL") : txt;
    set(id, "monto", {...cur.d, fecha:f});
    return bot.sendMessage(id, "Cual es el MONTO? (ej: 9350)", KB.cancelar);
  }
  if (cur.paso==="monto") {
    const m = Number(txt.replace(/\./g,"").replace(/,/g,"."));
    if (isNaN(m)||m<=0) return bot.sendMessage(id, "Escribe solo el numero (ej: 9350)");
    set(id, "comercio", {...cur.d, monto:m, moneda:"CLP"});
    return bot.sendMessage(id, "Cual es el COMERCIO? (o Omitir)", KB.omitir);
  }
  if (cur.paso==="comercio") { await pedirMotivo(id, {...cur.d, comercio:txt==="Omitir"?"":txt}); return; }
  if (cur.paso==="motivo") {
    if (!txt) return;
    set(id, "destino", {...cur.d, motivo:txt});
    return bot.sendMessage(id, "Cual es el DESTINO?\n(ej: Administrativo, Proyecto X, RRHH...)", KB.cancelar);
  }
  if (cur.paso==="destino") {
    if (!txt) return;
    set(id, "detalle", {...cur.d, destino:txt});
    return bot.sendMessage(id, "Agrega un DETALLE (o Omitir)", KB.omitir);
  }
  if (cur.paso==="detalle") {
    const d = {...cur.d, detalle:txt==="Omitir"?"":txt};
    set(id, "confirmar", d);
    const corr = String(d.corr).padStart(4,"0");
    const mFmt = d.monto!=null ? `${d.moneda} $${Number(d.monto).toLocaleString("es-CL")}` : "-";
    return bot.sendMessage(id,
      `Resumen - Gasto #${corr}\nUsuario: ${cur.usuario}\n\nFecha:    ${d.fecha}\nMonto:    ${mFmt}\nComercio: ${d.comercio||"-"}\nMotivo:   ${d.motivo}\nDestino:  ${d.destino}\nDetalle:  ${d.detalle||"-"}\n\nConfirmar y guardar?`,
      KB.confirmar
    );
  }
  if (cur.paso==="confirmar" && txt==="Confirmar") {
    try {
      const d      = cur.d;
      const corr   = String(d.corr).padStart(4,"0");
      const nombre = `${cur.usuario}_GASTO_${corr}.jpg`;
      let fotoUrl  = "";
      try { fotoUrl = await subirCloudinary(d.buffer, nombre); console.log("Foto subida:", nombre); }
      catch(e) { console.error("Error Cloudinary:", e.response?.data||e.message); }
      const corrStr = await agregarFila(d, fotoUrl, cur.usuario);
      S[id].corr++;
      reset(id);
      return bot.sendMessage(id,
        `${corrStr} guardado en tu planilla!\n\n${fotoUrl?`Foto guardada.\n`:""}Ver planilla:\nhttps://docs.google.com/spreadsheets/d/${process.env.GOOGLE_SHEET_ID}\n\nEnvia otra foto para seguir registrando.`,
        KB.inicio
      );
    } catch(err) {
      console.error("Error guardando:", err.response?.data||err.message);
      return bot.sendMessage(id, "Error al guardar. Escribe /cancelar e intenta de nuevo.");
    }
  }

  if (cur.paso==="inicio") bot.sendMessage(id, "Envia una foto de boleta.\n/planilla - ver tu planilla", KB.inicio);
});

bot.on("polling_error", err => console.error("Polling error:", err.message));
console.log("Bot iniciado con usuarios y hojas separadas!");
