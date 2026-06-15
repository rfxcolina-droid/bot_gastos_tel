require("dotenv").config();
const TelegramBot = require("node-telegram-bot-api");
const axios       = require("axios");
const crypto      = require("crypto");
const FormData    = require("form-data");

const bot = new TelegramBot(process.env.TELEGRAM_TOKEN, { polling: true });

// ── Usuarios ──────────────────────────────────────────────
const USUARIOS = {
  "2691": "Francisco",
  "1430": "Nicolas",
  // "CODIGO3": "Nombre3",
  // "CODIGO4": "Nombre4",
  // "CODIGO5": "Nombre5",
  // "CODIGO6": "Nombre6",
  // "CODIGO7": "Nombre7",
  // "CODIGO8": "Nombre8",
};

// ── Motivos ───────────────────────────────────────────────
const MOTIVOS = [
  "Alimentacion",
  "Combustible",
  "Materiales",
  "Herramientas",
  "Peaje",
  // "Motivo6",
  // "Motivo7",
  // "Motivo8",
  // "Motivo9",
  // "Motivo10",
];

const KB_MOTIVOS = {
  reply_markup: {
    keyboard: [
      [MOTIVOS[0], MOTIVOS[1]],
      [MOTIVOS[2], MOTIVOS[3]],
      [MOTIVOS[4], "Otro motivo..."],
      ["Cancelar"]
    ],
    resize_keyboard: true,
    one_time_keyboard: true
  }
};

// ── Sesiones ──────────────────────────────────────────────
const S = {};
function ses(id) { if (!S[id]) S[id] = { paso:"login", d:{}, usuario:null, corr:1 }; return S[id]; }
function set(id, paso, d) { S[id] = { ...ses(id), paso, d: d !== undefined ? d : ses(id).d }; }
function reset(id) { const u=ses(id); S[id]={ paso:"inicio", d:{}, usuario:u.usuario, corr:u.corr }; }

// ── Cloudinary ────────────────────────────────────────────
async function subirCloudinary(buffer, nombre) {
  const form = new FormData();
  form.append("file", buffer, { filename:nombre, contentType:"image/jpeg" });
  form.append("upload_preset", process.env.CLOUDINARY_UPLOAD_PRESET);
  form.append("public_id", nombre.replace(".jpg",""));
  const res = await axios.post(
    `https://api.cloudinary.com/v1_1/${process.env.CLOUDINARY_CLOUD_NAME}/image/upload`,
    form, { headers:form.getHeaders(), timeout:30000 }
  );
  return res.data.secure_url;
}

// ── Google Sheets ─────────────────────────────────────────
async function getToken() {
  const creds = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const now   = Math.floor(Date.now()/1000);
  const hdr   = Buffer.from(JSON.stringify({alg:"RS256",typ:"JWT"})).toString("base64url");
  const pay   = Buffer.from(JSON.stringify({
    iss:creds.client_email, scope:"https://www.googleapis.com/auth/spreadsheets",
    aud:"https://oauth2.googleapis.com/token", exp:now+3600, iat:now,
  })).toString("base64url");
  const sign = crypto.createSign("RSA-SHA256");
  sign.update(`${hdr}.${pay}`);
  const sig = sign.sign(creds.private_key,"base64url");
  const res = await axios.post("https://oauth2.googleapis.com/token",{
    grant_type:"urn:ietf:params:oauth:grant-type:jwt-bearer", assertion:`${hdr}.${pay}.${sig}`,
  });
  return res.data.access_token;
}

async function asegurarHoja(token, sheetId, hoja) {
  const res = await axios.get(
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=sheets.properties.title`,
    { headers:{ Authorization:`Bearer ${token}` } }
  );
  const hojas = res.data.sheets.map(h=>h.properties.title);
  if (!hojas.includes(hoja)) {
    await axios.post(
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}:batchUpdate`,
      { requests:[{ addSheet:{ properties:{ title:hoja } } }] },
      { headers:{ Authorization:`Bearer ${token}`, "Content-Type":"application/json" } }
    );
    await axios.put(
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${hoja}!A1:G1?valueInputOption=USER_ENTERED`,
      { values:[["Correlativo","Fecha","Motivo","Destino","Detalle","Monto","URL Foto"]] },
      { headers:{ Authorization:`Bearer ${token}`, "Content-Type":"application/json" } }
    );
  }
}

async function actualizarFilaPendiente(g, fotoUrl, usuario) {
  const token   = await getToken();
  const sheetId = process.env.GOOGLE_SHEET_ID;
  const corrStr = `GASTO_${String(g.corr).padStart(4,"0")}`;

  // Buscar la fila con ese correlativo
  const res = await axios.get(
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${usuario}!A:A`,
    { headers:{ Authorization:`Bearer ${token}` } }
  );
  const filas = res.data.values || [];
  let filaIdx = -1;
  for (let i=0; i<filas.length; i++) {
    if (filas[i][0] === corrStr) { filaIdx = i+1; break; }
  }

  const valores = [corrStr, g.fecha||"", g.motivo||"", g.destino||"", g.detalle||"", g.monto!=null?Number(g.monto):"", fotoUrl||""];

  if (filaIdx > 0) {
    // Actualizar fila existente
    await axios.put(
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${usuario}!A${filaIdx}:G${filaIdx}?valueInputOption=USER_ENTERED`,
      { values:[valores] },
      { headers:{ Authorization:`Bearer ${token}`, "Content-Type":"application/json" } }
    );
  } else {
    // No encontró la fila, agregar nueva
    await axios.post(
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${usuario}!A:G:append?valueInputOption=USER_ENTERED`,
      { values:[valores] },
      { headers:{ Authorization:`Bearer ${token}` } }
    );
  }
  return corrStr;
}

async function agregarFila(g, fotoUrl, usuario) {
  const token   = await getToken();
  const sheetId = process.env.GOOGLE_SHEET_ID;
  await asegurarHoja(token, sheetId, usuario);
  const corr = `GASTO_${String(g.corr).padStart(4,"0")}`;
  await axios.post(
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${usuario}!A:G:append?valueInputOption=USER_ENTERED`,
    { values:[[corr, g.fecha||"", g.motivo||"", g.destino||"", g.detalle||"", g.monto!=null?Number(g.monto):"", fotoUrl||""]] },
    { headers:{ Authorization:`Bearer ${token}` } }
  );
  return corr;
}

// ── Claude ────────────────────────────────────────────────
async function analizarImagen(fileId) {
  const fi   = await bot.getFile(fileId);
  const url  = `https://api.telegram.org/file/bot${process.env.TELEGRAM_TOKEN}/${fi.file_path}`;
  const ir   = await axios.get(url,{ responseType:"arraybuffer", timeout:20000 });
  const buf  = Buffer.from(ir.data);
  const ext  = fi.file_path.split(".").pop().toLowerCase();
  const mime = {jpg:"image/jpeg",jpeg:"image/jpeg",png:"image/png",webp:"image/webp"}[ext]||"image/jpeg";
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
  const texto = resp.data?.content?.[0]?.text||"";
  try { return { datos:JSON.parse(texto.replace(/```json|```/g,"").trim()), buffer:buf }; }
  catch { return { datos:{ error:"parse" }, buffer:buf }; }
}

// ── Mostrar resumen ───────────────────────────────────────
async function mostrarResumen(id) {
  const cur  = ses(id);
  const d    = cur.d;
  const corr = String(d.corr).padStart(4,"0");
  const mFmt = d.monto!=null ? `${d.moneda||"CLP"} $${Number(d.monto).toLocaleString("es-CL")}` : "-";

  // Enviar foto con link
  if (d.buffer) {
    await bot.sendPhoto(id, d.buffer, {
      caption: `Foto del gasto #${corr}\n\nVerifica que sea la foto correcta antes de confirmar.`
    });
  }

  await bot.sendMessage(id,
    `Resumen - Gasto #${corr}\nUsuario: ${cur.usuario}\n\n` +
    `📅 Fecha:    ${d.fecha||"-"}\n` +
    `💵 Monto:    ${mFmt}\n` +
    `🏪 Comercio: ${d.comercio||"-"}\n` +
    `🏷️ Motivo:   ${d.motivo||"-"}\n` +
    `📍 Destino:  ${d.destino||"-"}\n` +
    `📝 Detalle:  ${d.detalle||"-"}\n\n` +
    `Confirma o toca el campo que quieres editar:`,
    {
      reply_markup: {
        keyboard: [
          ["✅ Confirmar"],
          ["✏️ Fecha",    "✏️ Monto"],
          ["✏️ Comercio", "✏️ Motivo"],
          ["✏️ Destino",  "✏️ Detalle"],
          ["❌ Cancelar"]
        ],
        resize_keyboard: true,
        one_time_keyboard: true
      }
    }
  );
  set(id, "confirmar", d);
}

// ── Teclados ──────────────────────────────────────────────
const KB = {
  cancelar: { reply_markup:{ keyboard:[["❌ Cancelar"]], resize_keyboard:true } },
  omitir:   { reply_markup:{ keyboard:[["Omitir","❌ Cancelar"]], resize_keyboard:true, one_time_keyboard:true } },
  inicio:   { reply_markup:{ keyboard:[["📊 Ver planilla"]], resize_keyboard:true } },
};

// ── Manejador ─────────────────────────────────────────────
bot.on("message", async (msg) => {
  const id  = msg.chat.id;
  const txt = (msg.text||"").trim();
  const cur = ses(id);

  // ── Login ────────────────────────────────────────────
  if (!cur.usuario) {
    if (txt==="/start") return bot.sendMessage(id,"Bienvenido al Bot de Gastos.\n\nIngresa tu codigo de acceso:",{ reply_markup:{ remove_keyboard:true } });
    if (USUARIOS[txt]) {
      S[id] = { paso:"inicio", d:{}, usuario:USUARIOS[txt], corr:1 };
      return bot.sendMessage(id,`Bienvenido ${USUARIOS[txt]}!\n\nEnvia una FOTO de boleta para registrar un gasto.\n\nComandos:\n/planilla - ver tu planilla\n/saltar - saltar un correlativo (queda PENDIENTE)\n/agregar 0003 - completar un gasto pendiente\n/salir - cerrar sesion`,KB.inicio);
    }
    return bot.sendMessage(id,"Codigo incorrecto. Intenta nuevamente:");
  }

  // ── Comandos globales ────────────────────────────────
  if (txt==="/salir") { S[id]={ paso:"login", d:{}, usuario:null, corr:1 }; return bot.sendMessage(id,"Sesion cerrada.",{ reply_markup:{ remove_keyboard:true } }); }
  if (txt==="❌ Cancelar"||txt==="Cancelar"||txt==="/cancelar") { reset(id); return bot.sendMessage(id,"Cancelado.",KB.inicio); }
  if (txt==="/planilla"||txt==="📊 Ver planilla") return bot.sendMessage(id,`Tu planilla (hoja: ${cur.usuario}):\nhttps://docs.google.com/spreadsheets/d/${process.env.GOOGLE_SHEET_ID}`,KB.inicio);

  // Agregar gasto pendiente: /agregar 0003
  if (txt.startsWith("/agregar")) {
    if (cur.paso !== "inicio") return bot.sendMessage(id,"Cancela la operacion actual primero con /cancelar",KB.cancelar);
    const partes = txt.split(" ");
    if (partes.length < 2 || isNaN(Number(partes[1]))) {
      return bot.sendMessage(id,"Uso: /agregar NUMERO\nEjemplo: /agregar 0003\n\nEsto te permite completar un gasto pendiente o anterior.");
    }
    const numPendiente = Number(partes[1]);
    set(id, "foto_pendiente", { ...cur.d, corrPendiente: numPendiente });
    return bot.sendMessage(id,
      `Vas a completar el gasto GASTO_${String(numPendiente).padStart(4,"0")}.\n\nEnvia la foto de esa boleta:`,
      KB.cancelar
    );
  }

  // Saltar correlativo
  if (txt==="/saltar") {
    if (cur.paso !== "inicio") return bot.sendMessage(id,"Cancela la operacion actual primero con /cancelar",KB.cancelar);
    const corrSaltado = cur.corr;
    S[id].corr++;
    try {
      const token   = await getToken();
      const sheetId = process.env.GOOGLE_SHEET_ID;
      await asegurarHoja(token, sheetId, cur.usuario);
      const corrStr = `GASTO_${String(corrSaltado).padStart(4,"0")}`;
      await axios.post(
        `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${cur.usuario}!A:G:append?valueInputOption=USER_ENTERED`,
        { values:[[corrStr, "", "", "", "*** PENDIENTE ***", "", ""]] },
        { headers:{ Authorization:`Bearer ${token}` } }
      );
      return bot.sendMessage(id,
        `⏭️ ${corrStr} reservado como PENDIENTE en la planilla.\n\nCompleta los datos despues editando directamente en Google Sheets.\n\nEl proximo gasto sera GASTO_${String(S[id].corr).padStart(4,"0")}.\n\nEnvia una foto para continuar.`,
        KB.inicio
      );
    } catch(err) {
      S[id].corr--;
      console.error("Error saltando:", err.message);
      return bot.sendMessage(id,"Error al reservar el correlativo. Intenta de nuevo.");
    }
  }

  // ── Foto ─────────────────────────────────────────────
  if (msg.photo||msg.document?.mime_type?.startsWith("image/")) {
    // Modo agregar pendiente
    const esPendiente = cur.paso === "foto_pendiente";
    if (!esPendiente && cur.paso!=="inicio") return bot.sendMessage(id,"Escribe Cancelar primero.",KB.cancelar);
    const msj = await bot.sendMessage(id,"Leyendo boleta con IA...");
    try {
      const fid = msg.photo ? msg.photo[msg.photo.length-1].file_id : msg.document.file_id;
      const { datos, buffer } = await analizarImagen(fid);
      await bot.deleteMessage(id,msj.message_id).catch(()=>{});
      // Usar correlativo pendiente o el actual
      const corrUsado = esPendiente ? cur.d.corrPendiente : cur.corr;
      if (datos.error) {
        set(id,"fecha",{ corr:corrUsado, buffer, esPendiente });
        return bot.sendMessage(id,"No pude leer los datos automaticamente.\n\nCual es la FECHA?\n(ej: 13/06/2026 o 'hoy')",KB.cancelar);
      }
      // Ir directo a motivo con menú
      set(id,"motivo",{ corr:corrUsado, buffer, esPendiente, fecha:datos.fecha||new Date().toLocaleDateString("es-CL"), monto:datos.monto, moneda:datos.moneda||"CLP", comercio:datos.comercio||"" });
      const mFmt = datos.monto!=null ? `${datos.moneda||"CLP"} $${Number(datos.monto).toLocaleString("es-CL")}` : "no detectado";
      const corrStr2 = `GASTO_${String(corrUsado).padStart(4,"0")}`;
      await bot.sendMessage(id,
        `${esPendiente?`Completando ${corrStr2}\n`:""}Datos detectados:\n📅 Fecha: ${datos.fecha||"no detectada"}\n💵 Monto: ${mFmt}${datos.comercio?`\n🏪 Comercio: ${datos.comercio}`:""}\n\nSelecciona el MOTIVO:`,
        KB_MOTIVOS
      );
    } catch(err) {
      await bot.deleteMessage(id,msj.message_id).catch(()=>{});
      console.error("Error IA:",err.response?.data||err.message);
      const corrUsado2 = esPendiente ? cur.d.corrPendiente : cur.corr;
      set(id,"fecha",{ corr:corrUsado2, esPendiente });
      await bot.sendMessage(id,"Error con la IA. Ingresemos manualmente.\n\nCual es la FECHA?\n(ej: 13/06/2026 o 'hoy')",KB.cancelar);
    }
    return;
  }

  // ── Flujo manual ─────────────────────────────────────
  if (cur.paso==="fecha") {
    const f = txt.toLowerCase()==="hoy"?new Date().toLocaleDateString("es-CL"):txt;
    set(id,"monto",{...cur.d,fecha:f});
    return bot.sendMessage(id,"Cual es el MONTO? (ej: 9350)",KB.cancelar);
  }
  if (cur.paso==="monto") {
    const m = Number(txt.replace(/\./g,"").replace(/,/g,"."));
    if(isNaN(m)||m<=0) return bot.sendMessage(id,"Escribe solo el numero (ej: 9350)");
    set(id,"comercio",{...cur.d,monto:m,moneda:"CLP"});
    return bot.sendMessage(id,"Cual es el COMERCIO? (o Omitir)",KB.omitir);
  }
  if (cur.paso==="comercio") {
    set(id,"motivo",{...cur.d,comercio:txt==="Omitir"?"":txt});
    return bot.sendMessage(id,"Selecciona el MOTIVO:",KB_MOTIVOS);
  }

  // ── Motivo con menú ───────────────────────────────────
  if (cur.paso==="motivo") {
    if (!txt) return;
    if (txt==="Otro motivo...") {
      set(id,"motivo_manual",cur.d);
      return bot.sendMessage(id,"Escribe el motivo:",KB.cancelar);
    }
    set(id,"destino",{...cur.d,motivo:txt});
    return bot.sendMessage(id,"Cual es el DESTINO?\n(ej: Administrativo, Proyecto X, RRHH...)",KB.cancelar);
  }
  if (cur.paso==="motivo_manual") {
    if (!txt) return;
    set(id,"destino",{...cur.d,motivo:txt});
    return bot.sendMessage(id,"Cual es el DESTINO?\n(ej: Administrativo, Proyecto X, RRHH...)",KB.cancelar);
  }

  if (cur.paso==="destino") {
    if (!txt) return;
    set(id,"detalle",{...cur.d,destino:txt});
    return bot.sendMessage(id,
      "Agrega un DETALLE adicional.\n\nEscribe el detalle o toca Omitir.\n(Tip: puedes escribir varias lineas antes de enviar)",
      KB.omitir
    );
  }
  if (cur.paso==="detalle") {
    // Guardar detalle y mostrar resumen — NO borrar
    const d = {...cur.d, detalle:txt==="Omitir"?"":txt};
    set(id,"pre_resumen",d);
    await mostrarResumen(id);
    return;
  }

  // ── Confirmar / Editar ────────────────────────────────
  if (cur.paso==="confirmar") {
    if (txt==="✅ Confirmar") {
      try {
        const d      = cur.d;
        const corr   = String(d.corr).padStart(4,"0");
        const nombre = `${cur.usuario}_GASTO_${corr}.jpg`;
        let fotoUrl  = "";
        try { fotoUrl=await subirCloudinary(d.buffer,nombre); }
        catch(e) { console.error("Error Cloudinary:",e.response?.data||e.message); }
        let corrStr;
        if (d.esPendiente) {
          // Buscar y actualizar la fila pendiente
          corrStr = await actualizarFilaPendiente(d, fotoUrl, cur.usuario);
        } else {
          corrStr = await agregarFila(d, fotoUrl, cur.usuario);
          S[id].corr++;
        }
        reset(id);
        return bot.sendMessage(id,
          `✅ ${corrStr} guardado en tu planilla!\n\n📊 Ver planilla:\nhttps://docs.google.com/spreadsheets/d/${process.env.GOOGLE_SHEET_ID}\n\nEnvia otra foto para seguir registrando.`,
          KB.inicio
        );
      } catch(err) {
        console.error("Error:",err.response?.data||err.message);
        return bot.sendMessage(id,"Error al guardar. Escribe /cancelar e intenta de nuevo.");
      }
    }

    // Botones de edición
    if (txt==="✏️ Fecha")    { set(id,"editando_fecha",cur.d);    return bot.sendMessage(id,"Nueva FECHA:\n(ej: 13/06/2026 o 'hoy')",KB.cancelar); }
    if (txt==="✏️ Monto")    { set(id,"editando_monto",cur.d);    return bot.sendMessage(id,"Nuevo MONTO:\n(ej: 9350)",KB.cancelar); }
    if (txt==="✏️ Comercio") { set(id,"editando_comercio",cur.d); return bot.sendMessage(id,"Nuevo COMERCIO:\n(o Omitir)",KB.omitir); }
    if (txt==="✏️ Motivo")   { set(id,"editando_motivo",cur.d);   return bot.sendMessage(id,"Selecciona el nuevo MOTIVO:",KB_MOTIVOS); }
    if (txt==="✏️ Destino")  { set(id,"editando_destino",cur.d);  return bot.sendMessage(id,"Nuevo DESTINO:",KB.cancelar); }
    if (txt==="✏️ Detalle")  { set(id,"editando_detalle",cur.d);  return bot.sendMessage(id,"Nuevo DETALLE:\n(o Omitir)",KB.omitir); }
  }

  // ── Guardar ediciones ─────────────────────────────────
  if (cur.paso==="editando_fecha")    { const f=txt.toLowerCase()==="hoy"?new Date().toLocaleDateString("es-CL"):txt; set(id,"pre_resumen",{...cur.d,fecha:f}); await mostrarResumen(id); return; }
  if (cur.paso==="editando_monto")    { const m=Number(txt.replace(/\./g,"").replace(/,/g,".")); if(isNaN(m)||m<=0) return bot.sendMessage(id,"Escribe solo el numero (ej: 9350)"); set(id,"pre_resumen",{...cur.d,monto:m}); await mostrarResumen(id); return; }
  if (cur.paso==="editando_comercio") { set(id,"pre_resumen",{...cur.d,comercio:txt==="Omitir"?"":txt}); await mostrarResumen(id); return; }
  if (cur.paso==="editando_motivo")   {
    if (txt==="Otro motivo...") { set(id,"editando_motivo_manual",cur.d); return bot.sendMessage(id,"Escribe el motivo:",KB.cancelar); }
    set(id,"pre_resumen",{...cur.d,motivo:txt}); await mostrarResumen(id); return;
  }
  if (cur.paso==="editando_motivo_manual") { set(id,"pre_resumen",{...cur.d,motivo:txt}); await mostrarResumen(id); return; }
  if (cur.paso==="editando_destino")  { set(id,"pre_resumen",{...cur.d,destino:txt}); await mostrarResumen(id); return; }
  if (cur.paso==="editando_detalle")  { set(id,"pre_resumen",{...cur.d,detalle:txt==="Omitir"?"":txt}); await mostrarResumen(id); return; }

  if (cur.paso==="inicio") bot.sendMessage(id,"Envia una foto de boleta.\n/planilla - ver tu planilla",KB.inicio);
});

bot.on("polling_error", err=>console.error("Polling error:",err.message));
console.log("Bot iniciado!");
