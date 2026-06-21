require("dotenv").config();
const TelegramBot = require("node-telegram-bot-api");
const axios       = require("axios");
const crypto      = require("crypto");
const FormData    = require("form-data");
const ExcelJS     = require("exceljs");
const fs          = require("fs");
const path        = require("path");

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
      ["❌ Cancelar"]
    ],
    resize_keyboard: true,
    one_time_keyboard: true
  }
};

// ── Destinos ──────────────────────────────────────────────
const DESTINOS = [
  "PFV Los Quilos",
  "PFV Las Violetas",
  "PFV San Ramon",
  // "Destino4",
  // "Destino5",
  // "Destino6",
  // "Destino7",
  // "Destino8",
];

const KB_DESTINOS = {
  reply_markup: {
    keyboard: [
      [DESTINOS[0]],
      [DESTINOS[1]],
      [DESTINOS[2]],
      ["Otro destino...", "❌ Cancelar"]
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

// ── Google Auth (solo Drive) ──────────────────────────────
async function getTokenDrive() {
  const creds = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const now   = Math.floor(Date.now()/1000);
  const hdr   = Buffer.from(JSON.stringify({alg:"RS256",typ:"JWT"})).toString("base64url");
  const pay   = Buffer.from(JSON.stringify({
    iss:creds.client_email, scope:"https://www.googleapis.com/auth/drive",
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

// ── Buscar archivo Excel del usuario en Drive ────────────
async function buscarArchivoExcel(usuario, token) {
  const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;
  const nombre   = `${usuario}_Gastos.xlsx`;
  const busq = await axios.get(
    `https://www.googleapis.com/drive/v3/files?q=name='${nombre}' and '${folderId}' in parents and trashed=false&fields=files(id,name)`,
    { headers:{ Authorization:`Bearer ${token}` } }
  );
  return busq.data.files.length > 0 ? busq.data.files[0].id : null;
}

// ── Cargar workbook desde Drive (o crear nuevo) ──────────
async function cargarWorkbook(usuario, token, fileId) {
  const wb = new ExcelJS.Workbook();
  if (fileId) {
    const res = await axios.get(
      `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`,
      { headers:{ Authorization:`Bearer ${token}` }, responseType:"arraybuffer" }
    );
    await wb.xlsx.load(Buffer.from(res.data));
  } else {
    const ws = wb.addWorksheet(usuario);
    ws.columns = [
      { header:"Correlativo", key:"a", width:14 },
      { header:"Fecha",       key:"b", width:13 },
      { header:"Motivo",      key:"c", width:20 },
      { header:"Destino",     key:"d", width:22 },
      { header:"Detalle",     key:"e", width:32 },
      { header:"Monto",       key:"f", width:14 },
      { header:"URL Foto",    key:"g", width:45 },
    ];
    ws.getRow(1).eachCell(c => {
      c.fill = { type:"pattern", pattern:"solid", fgColor:{ argb:"FF128C7E" } };
      c.font = { bold:true, color:{ argb:"FFFFFFFF" }, size:11 };
      c.alignment = { horizontal:"center", vertical:"middle" };
    });
    ws.getRow(1).height = 22;
  }
  return wb;
}

// ── Subir/actualizar workbook en Drive ────────────────────
async function guardarWorkbook(wb, usuario, token, fileId) {
  const nombre  = `${usuario}_Gastos.xlsx`;
  const tmpPath = path.join("/tmp", nombre);
  await wb.xlsx.writeFile(tmpPath);
  const buffer   = fs.readFileSync(tmpPath);
  const mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

  if (fileId) {
    await axios.patch(
      `https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=media`,
      buffer,
      { headers:{ Authorization:`Bearer ${token}`, "Content-Type":mimeType } }
    );
  } else {
    const folderId  = process.env.GOOGLE_DRIVE_FOLDER_ID;
    const boundary  = "excel_boundary_123";
    const metadata  = JSON.stringify({ name:nombre, parents:[folderId] });
    const body = Buffer.concat([
      Buffer.from(`--${boundary}\r\nContent-Type: application/json\r\n\r\n${metadata}\r\n--${boundary}\r\nContent-Type: ${mimeType}\r\n\r\n`),
      buffer,
      Buffer.from(`\r\n--${boundary}--`),
    ]);
    const res = await axios.post(
      `https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id`,
      body,
      { headers:{ Authorization:`Bearer ${token}`, "Content-Type":`multipart/related; boundary=${boundary}` } }
    );
    await axios.post(
      `https://www.googleapis.com/drive/v3/files/${res.data.id}/permissions`,
      { role:"reader", type:"anyone" },
      { headers:{ Authorization:`Bearer ${token}` } }
    );
    fileId = res.data.id;
  }
  fs.unlinkSync(tmpPath);
  return fileId;
}

// ── Agregar gasto al Excel ─────────────────────────────────
async function agregarGastoExcel(usuario, g, fotoUrl) {
  const token  = await getTokenDrive();
  const fileId = await buscarArchivoExcel(usuario, token);
  const wb     = await cargarWorkbook(usuario, token, fileId);
  const ws     = wb.getWorksheet(usuario) || wb.worksheets[0];

  const corr = `GASTO_${String(g.corr).padStart(4,"0")}`;
  const fila = ws.addRow([corr, g.fecha||"", g.motivo||"", g.destino||"", g.detalle||"", g.monto!=null?Number(g.monto):"", fotoUrl||""]);
  fila.getCell(6).numFmt = "#,##0";
  if (fila.number % 2 === 0) {
    fila.eachCell({ includeEmpty:true }, c => {
      c.fill = { type:"pattern", pattern:"solid", fgColor:{ argb:"FFF1F8F6" } };
    });
  }

  await guardarWorkbook(wb, usuario, token, fileId);
  return corr;
}

// ── Actualizar fila pendiente ──────────────────────────────
async function actualizarFilaPendiente(usuario, g, fotoUrl) {
  const token  = await getTokenDrive();
  const fileId = await buscarArchivoExcel(usuario, token);
  const wb     = await cargarWorkbook(usuario, token, fileId);
  const ws     = wb.getWorksheet(usuario) || wb.worksheets[0];

  const corr = `GASTO_${String(g.corr).padStart(4,"0")}`;
  let filaEncontrada = null;
  ws.eachRow((fila, idx) => {
    if (idx===1) return;
    if (fila.getCell(1).value === corr) filaEncontrada = fila;
  });

  if (filaEncontrada) {
    filaEncontrada.getCell(1).value = corr;
    filaEncontrada.getCell(2).value = g.fecha||"";
    filaEncontrada.getCell(3).value = g.motivo||"";
    filaEncontrada.getCell(4).value = g.destino||"";
    filaEncontrada.getCell(5).value = g.detalle||"";
    filaEncontrada.getCell(6).value = g.monto!=null?Number(g.monto):"";
    filaEncontrada.getCell(6).numFmt = "#,##0";
    filaEncontrada.getCell(7).value = fotoUrl||"";
  } else {
    const fila = ws.addRow([corr, g.fecha||"", g.motivo||"", g.destino||"", g.detalle||"", g.monto!=null?Number(g.monto):"", fotoUrl||""]);
    fila.getCell(6).numFmt = "#,##0";
  }

  await guardarWorkbook(wb, usuario, token, fileId);
  return corr;
}

// ── Saltar correlativo (fila PENDIENTE) ────────────────────
async function saltarCorrelativo(usuario, corrNum) {
  const token  = await getTokenDrive();
  const fileId = await buscarArchivoExcel(usuario, token);
  const wb     = await cargarWorkbook(usuario, token, fileId);
  const ws     = wb.getWorksheet(usuario) || wb.worksheets[0];

  const corr = `GASTO_${String(corrNum).padStart(4,"0")}`;
  ws.addRow([corr, "", "", "", "*** PENDIENTE ***", "", ""]);
  await guardarWorkbook(wb, usuario, token, fileId);
  return corr;
}

// ── Obtener último correlativo ────────────────────────────
async function obtenerUltimoCorrelativo(usuario) {
  try {
    const token  = await getTokenDrive();
    const fileId = await buscarArchivoExcel(usuario, token);
    if (!fileId) return 1;
    const wb = await cargarWorkbook(usuario, token, fileId);
    const ws = wb.getWorksheet(usuario) || wb.worksheets[0];
    let maxCorr = 0;
    ws.eachRow((fila, idx) => {
      if (idx===1) return;
      const val = String(fila.getCell(1).value||"");
      const match = val.match(/(\d{4})$/);
      if (match) { const n=parseInt(match[1]); if(n>maxCorr) maxCorr=n; }
    });
    return maxCorr + 1;
  } catch(e) {
    console.error("Error obteniendo correlativo:", e.message);
    return 1;
  }
}

// ── Nueva planilla (renombrar archivo actual) ─────────────
async function nuevaPlanilla(usuario) {
  const token  = await getTokenDrive();
  const fileId = await buscarArchivoExcel(usuario, token);
  if (!fileId) return null;
  const fecha = new Date().toLocaleDateString("es-CL").replace(/\//g,"-");
  const nuevoNombre = `${usuario}_Gastos_${fecha}.xlsx`;
  await axios.patch(
    `https://www.googleapis.com/drive/v3/files/${fileId}`,
    { name: nuevoNombre },
    { headers:{ Authorization:`Bearer ${token}`, "Content-Type":"application/json" } }
  );
  return nuevoNombre;
}

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

// ── Teclados ──────────────────────────────────────────────
const KB = {
  cancelar: { reply_markup:{ keyboard:[["❌ Cancelar"]], resize_keyboard:true } },
  omitir:   { reply_markup:{ keyboard:[["Omitir","❌ Cancelar"]], resize_keyboard:true, one_time_keyboard:true } },
  inicio:   { reply_markup:{ keyboard:[
    ["📊 Ver planilla"],
    ["🔄 Nueva planilla"]
  ], resize_keyboard:true } },
};

// ── Mostrar resumen ───────────────────────────────────────
async function mostrarResumen(id) {
  const cur  = ses(id);
  const d    = cur.d;
  const corr = String(d.corr).padStart(4,"0");
  const mFmt = d.monto!=null ? `${d.moneda||"CLP"} $${Number(d.monto).toLocaleString("es-CL")}` : "-";
  if (d.buffer) {
    await bot.sendPhoto(id, d.buffer, { caption:`Foto del gasto #${corr}\nVerifica que sea la foto correcta.` });
  }
  await bot.sendMessage(id,
    `Resumen - Gasto #${corr}\nUsuario: ${cur.usuario}\n\n` +
    `📅 Fecha:    ${d.fecha||"-"}\n💵 Monto:    ${mFmt}\n🏪 Comercio: ${d.comercio||"-"}\n` +
    `🏷️ Motivo:   ${d.motivo||"-"}\n📍 Destino:  ${d.destino||"-"}\n📝 Detalle:  ${d.detalle||"-"}\n\n` +
    `Confirma o toca el campo que quieres editar:`,
    { reply_markup:{ keyboard:[
      ["✅ Confirmar"],
      ["✏️ Fecha","✏️ Monto"],
      ["✏️ Comercio","✏️ Motivo"],
      ["✏️ Destino","✏️ Detalle"],
      ["❌ Cancelar"]
    ], resize_keyboard:true, one_time_keyboard:true } }
  );
  set(id,"confirmar",d);
}

// ── Manejador ─────────────────────────────────────────────
bot.on("message", async (msg) => {
  const id  = msg.chat.id;
  const txt = (msg.text||"").trim();
  const cur = ses(id);

  // ── Login ────────────────────────────────────────────
  if (!cur.usuario) {
    if (txt==="/start") return bot.sendMessage(id,"Bienvenido al Bot de Gastos.\n\nIngresa tu codigo de acceso:",{ reply_markup:{ remove_keyboard:true } });
    if (USUARIOS[txt]) {
      const nombre = USUARIOS[txt];
      S[id] = { paso:"inicio", d:{}, usuario:nombre, corr:1 };
      await bot.sendMessage(id,`Bienvenido ${nombre}! Cargando tu planilla...`);
      try {
        const ult = await obtenerUltimoCorrelativo(nombre);
        S[id].corr = ult;
      } catch(e) { console.error(e.message); }
      return bot.sendMessage(id,
        `Listo! Tu proximo gasto sera GASTO_${String(S[id].corr).padStart(4,"0")}\n\nEnvia una FOTO de boleta para registrar.\n\nComandos:\n/planilla - ver tu Excel\n/saltar - saltar un correlativo\n/agregar 0003 - completar un gasto pendiente\n/salir - cerrar sesion`,
        KB.inicio
      );
    }
    return bot.sendMessage(id,"Codigo incorrecto. Intenta nuevamente:");
  }

  // ── Comandos globales ────────────────────────────────
  if (txt==="/salir") { S[id]={ paso:"login", d:{}, usuario:null, corr:1 }; return bot.sendMessage(id,"Sesion cerrada.",{ reply_markup:{ remove_keyboard:true } }); }
  if (txt==="❌ Cancelar"||txt==="Cancelar"||txt==="/cancelar") { reset(id); return bot.sendMessage(id,"Cancelado.",KB.inicio); }

  if (txt==="/planilla"||txt==="📊 Ver planilla") {
    const msj = await bot.sendMessage(id,"Buscando tu planilla...");
    try {
      const token  = await getTokenDrive();
      const fileId = await buscarArchivoExcel(cur.usuario, token);
      await bot.deleteMessage(id,msj.message_id).catch(()=>{});
      if (!fileId) return bot.sendMessage(id,"Aun no tienes gastos registrados.",KB.inicio);
      return bot.sendMessage(id,`Tu planilla:\nhttps://docs.google.com/spreadsheets/d/${fileId}/edit\n\n(Se abre con Google Sheets, o descarga el .xlsx desde Drive)`,KB.inicio);
    } catch(e) {
      await bot.deleteMessage(id,msj.message_id).catch(()=>{});
      return bot.sendMessage(id,"Error buscando la planilla.",KB.inicio);
    }
  }

  // ── Saltar correlativo ───────────────────────────────
  if (txt==="/saltar") {
    if (cur.paso!=="inicio") return bot.sendMessage(id,"Cancela la operacion actual primero.",KB.cancelar);
    const corrSaltado = cur.corr;
    S[id].corr++;
    try {
      const corrStr = await saltarCorrelativo(cur.usuario, corrSaltado);
      return bot.sendMessage(id,`⏭️ ${corrStr} reservado como PENDIENTE.\n\nEl proximo gasto sera GASTO_${String(S[id].corr).padStart(4,"0")}.\n\nEnvia una foto para continuar.`,KB.inicio);
    } catch(e) {
      S[id].corr--;
      console.error("Error saltando:",e.message);
      return bot.sendMessage(id,"Error al saltar. Intenta de nuevo.");
    }
  }

  // ── Agregar pendiente ─────────────────────────────────
  if (txt.startsWith("/agregar")) {
    if (cur.paso!=="inicio") return bot.sendMessage(id,"Cancela la operacion actual primero.",KB.cancelar);
    const partes = txt.split(" ");
    if (partes.length<2||isNaN(Number(partes[1]))) return bot.sendMessage(id,"Uso: /agregar NUMERO\nEjemplo: /agregar 0003");
    const numPendiente = Number(partes[1]);
    set(id,"foto_pendiente",{ ...cur.d, corrPendiente:numPendiente });
    return bot.sendMessage(id,`Vas a completar GASTO_${String(numPendiente).padStart(4,"0")}.\n\nEnvia la foto de esa boleta:`,KB.cancelar);
  }

  // ── Nueva planilla ───────────────────────────────────
  if (txt==="🔄 Nueva planilla"||txt==="/nuevaplanilla") {
    if (cur.paso!=="inicio") return bot.sendMessage(id,"Cancela la operacion actual primero.",KB.cancelar);
    return bot.sendMessage(id,
      "⚠️ Esto archivara tu planilla actual y comenzara una nueva.\n\nLa planilla actual quedara guardada con la fecha de hoy en Drive.\n\n¿Confirmas?",
      { reply_markup:{ keyboard:[["✅ Si, nueva planilla","❌ Cancelar"]], resize_keyboard:true, one_time_keyboard:true } }
    );
  }

  if (txt==="✅ Si, nueva planilla") {
    const msj = await bot.sendMessage(id,"Archivando planilla actual...");
    try {
      const nombreArchivado = await nuevaPlanilla(cur.usuario);
      S[id].corr = 1;
      await bot.deleteMessage(id,msj.message_id).catch(()=>{});
      return bot.sendMessage(id,
        `✅ Planilla archivada como "${nombreArchivado||"sin gastos previos"}".\n\nTu nueva planilla esta lista. El proximo gasto sera GASTO_0001.\n\nEnvia una foto para empezar.`,
        KB.inicio
      );
    } catch(e) {
      console.error("Error nueva planilla:",e.message);
      await bot.deleteMessage(id,msj.message_id).catch(()=>{});
      return bot.sendMessage(id,"Error al crear nueva planilla. Intenta de nuevo.",KB.inicio);
    }
  }

  // ── Foto ─────────────────────────────────────────────
  if (msg.photo||msg.document?.mime_type?.startsWith("image/")) {
    const esPendiente = cur.paso==="foto_pendiente";
    if (!esPendiente&&cur.paso!=="inicio") return bot.sendMessage(id,"Escribe Cancelar primero.",KB.cancelar);
    const msj = await bot.sendMessage(id,"Leyendo boleta con IA...");
    try {
      const fid = msg.photo ? msg.photo[msg.photo.length-1].file_id : msg.document.file_id;
      const { datos, buffer } = await analizarImagen(fid);
      await bot.deleteMessage(id,msj.message_id).catch(()=>{});
      const corrUsado = esPendiente ? cur.d.corrPendiente : cur.corr;
      if (datos.error) {
        set(id,"fecha",{ corr:corrUsado, buffer, esPendiente });
        return bot.sendMessage(id,"No pude leer los datos.\n\nCual es la FECHA?\n(ej: 13/06/2026 o 'hoy')",KB.cancelar);
      }
      set(id,"motivo",{ corr:corrUsado, buffer, esPendiente, fecha:datos.fecha||new Date().toLocaleDateString("es-CL"), monto:datos.monto, moneda:datos.moneda||"CLP", comercio:datos.comercio||"" });
      const mFmt = datos.monto!=null?`${datos.moneda||"CLP"} $${Number(datos.monto).toLocaleString("es-CL")}`:"no detectado";
      const corrStr = `GASTO_${String(corrUsado).padStart(4,"0")}`;
      await bot.sendMessage(id,`${esPendiente?`Completando ${corrStr}\n`:""}Datos detectados:\n📅 Fecha: ${datos.fecha||"no detectada"}\n💵 Monto: ${mFmt}${datos.comercio?`\n🏪 Comercio: ${datos.comercio}`:""}\n\nSelecciona el MOTIVO:`,KB_MOTIVOS);
    } catch(e) {
      await bot.deleteMessage(id,msj.message_id).catch(()=>{});
      console.error("Error IA:",e.response?.data||e.message);
      const corrUsado = esPendiente?cur.d.corrPendiente:cur.corr;
      set(id,"fecha",{ corr:corrUsado, esPendiente });
      await bot.sendMessage(id,"Error con la IA. Ingresemos manualmente.\n\nCual es la FECHA?\n(ej: 13/06/2026 o 'hoy')",KB.cancelar);
    }
    return;
  }

  // ── Flujo ─────────────────────────────────────────────
  if (cur.paso==="fecha") { const f=txt.toLowerCase()==="hoy"?new Date().toLocaleDateString("es-CL"):txt; set(id,"monto",{...cur.d,fecha:f}); return bot.sendMessage(id,"Cual es el MONTO? (ej: 9350)",KB.cancelar); }
  if (cur.paso==="monto") { const m=Number(txt.replace(/\./g,"").replace(/,/g,".")); if(isNaN(m)||m<=0) return bot.sendMessage(id,"Escribe solo el numero (ej: 9350)"); set(id,"comercio",{...cur.d,monto:m,moneda:"CLP"}); return bot.sendMessage(id,"Cual es el COMERCIO? (o Omitir)",KB.omitir); }
  if (cur.paso==="comercio") { set(id,"motivo",{...cur.d,comercio:txt==="Omitir"?"":txt}); return bot.sendMessage(id,"Selecciona el MOTIVO:",KB_MOTIVOS); }
  if (cur.paso==="motivo") {
    if (!txt) return;
    if (txt==="Otro motivo...") { set(id,"motivo_manual",cur.d); return bot.sendMessage(id,"Escribe el motivo:",KB.cancelar); }
    set(id,"destino",{...cur.d,motivo:txt});
    return bot.sendMessage(id,"Selecciona el DESTINO:",KB_DESTINOS);
  }
  if (cur.paso==="motivo_manual") { if(!txt)return; set(id,"destino",{...cur.d,motivo:txt}); return bot.sendMessage(id,"Selecciona el DESTINO:",KB_DESTINOS); }
  if (cur.paso==="destino") {
    if (!txt) return;
    if (txt==="Otro destino...") { set(id,"destino_manual",cur.d); return bot.sendMessage(id,"Escribe el destino:",KB.cancelar); }
    set(id,"detalle",{...cur.d,destino:txt});
    return bot.sendMessage(id,"Agrega un DETALLE adicional.\n(o toca Omitir)",KB.omitir);
  }
  if (cur.paso==="destino_manual") { if(!txt)return; set(id,"detalle",{...cur.d,destino:txt}); return bot.sendMessage(id,"Agrega un DETALLE adicional.\n(o toca Omitir)",KB.omitir); }
  if (cur.paso==="detalle") { const d={...cur.d,detalle:txt==="Omitir"?"":txt}; set(id,"pre_resumen",d); await mostrarResumen(id); return; }

  // ── Confirmar / Editar ────────────────────────────────
  if (cur.paso==="confirmar") {
    if (txt==="✅ Confirmar") {
      const msj = await bot.sendMessage(id,"Guardando gasto...");
      try {
        const d      = cur.d;
        const corr   = String(d.corr).padStart(4,"0");
        const nombre = `${cur.usuario}_GASTO_${corr}.jpg`;
        let fotoUrl  = "";
        try { fotoUrl=await subirCloudinary(d.buffer,nombre); } catch(e) { console.error("Cloudinary:",e.message); }

        let corrStr;
        if (d.esPendiente) {
          corrStr = await actualizarFilaPendiente(cur.usuario, d, fotoUrl);
        } else {
          corrStr = await agregarGastoExcel(cur.usuario, d, fotoUrl);
          S[id].corr++;
        }

        reset(id);
        await bot.deleteMessage(id,msj.message_id).catch(()=>{});
        return bot.sendMessage(id,
          `✅ ${corrStr} guardado en tu planilla!\n\nEnvia otra foto para seguir registrando.\n/planilla - ver tu Excel`,
          KB.inicio
        );
      } catch(e) {
        await bot.deleteMessage(id,msj.message_id).catch(()=>{});
        console.error("Error guardando:",e.response?.data||e.message);
        return bot.sendMessage(id,"Error al guardar. Escribe /cancelar e intenta de nuevo.");
      }
    }
    if (txt==="✏️ Fecha")    { set(id,"editando_fecha",cur.d);    return bot.sendMessage(id,"Nueva FECHA:\n(ej: 13/06/2026 o 'hoy')",KB.cancelar); }
    if (txt==="✏️ Monto")    { set(id,"editando_monto",cur.d);    return bot.sendMessage(id,"Nuevo MONTO:\n(ej: 9350)",KB.cancelar); }
    if (txt==="✏️ Comercio") { set(id,"editando_comercio",cur.d); return bot.sendMessage(id,"Nuevo COMERCIO:\n(o Omitir)",KB.omitir); }
    if (txt==="✏️ Motivo")   { set(id,"editando_motivo",cur.d);   return bot.sendMessage(id,"Selecciona el nuevo MOTIVO:",KB_MOTIVOS); }
    if (txt==="✏️ Destino")  { set(id,"editando_destino",cur.d);  return bot.sendMessage(id,"Selecciona el nuevo DESTINO:",KB_DESTINOS); }
    if (txt==="✏️ Detalle")  { set(id,"editando_detalle",cur.d);  return bot.sendMessage(id,"Nuevo DETALLE:\n(o Omitir)",KB.omitir); }
  }

  // ── Ediciones ─────────────────────────────────────────
  if (cur.paso==="editando_fecha")    { const f=txt.toLowerCase()==="hoy"?new Date().toLocaleDateString("es-CL"):txt; set(id,"pre_resumen",{...cur.d,fecha:f}); await mostrarResumen(id); return; }
  if (cur.paso==="editando_monto")    { const m=Number(txt.replace(/\./g,"").replace(/,/g,".")); if(isNaN(m)||m<=0) return bot.sendMessage(id,"Escribe solo el numero"); set(id,"pre_resumen",{...cur.d,monto:m}); await mostrarResumen(id); return; }
  if (cur.paso==="editando_comercio") { set(id,"pre_resumen",{...cur.d,comercio:txt==="Omitir"?"":txt}); await mostrarResumen(id); return; }
  if (cur.paso==="editando_motivo") {
    if (txt==="Otro motivo...") { set(id,"editando_motivo_manual",cur.d); return bot.sendMessage(id,"Escribe el motivo:",KB.cancelar); }
    set(id,"pre_resumen",{...cur.d,motivo:txt}); await mostrarResumen(id); return;
  }
  if (cur.paso==="editando_motivo_manual") { set(id,"pre_resumen",{...cur.d,motivo:txt}); await mostrarResumen(id); return; }
  if (cur.paso==="editando_destino") {
    if (txt==="Otro destino...") { set(id,"editando_destino_manual",cur.d); return bot.sendMessage(id,"Escribe el destino:",KB.cancelar); }
    set(id,"pre_resumen",{...cur.d,destino:txt}); await mostrarResumen(id); return;
  }
  if (cur.paso==="editando_destino_manual") { set(id,"pre_resumen",{...cur.d,destino:txt}); await mostrarResumen(id); return; }
  if (cur.paso==="editando_detalle")  { set(id,"pre_resumen",{...cur.d,detalle:txt==="Omitir"?"":txt}); await mostrarResumen(id); return; }

  if (cur.paso==="inicio") bot.sendMessage(id,"Envia una foto de boleta.\n/planilla - ver tu Excel",KB.inicio);
});

bot.on("polling_error", err=>console.error("Polling error:",err.message));
console.log("Bot iniciado - Solo Excel en Drive por usuario!");
