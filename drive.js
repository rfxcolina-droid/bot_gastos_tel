const { google } = require("googleapis");

let driveClient = null;

function getDrive() {
  if (driveClient) return driveClient;
  const creds = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const auth  = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: ["https://www.googleapis.com/auth/drive"],
  });
  driveClient = google.drive({ version: "v3", auth });
  return driveClient;
}

// Subir foto a Google Drive y devolver URL
async function subirFoto(buffer, nombre) {
  const drive    = getDrive();
  const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;
  const { Readable } = require("stream");

  const stream = new Readable();
  stream.push(buffer);
  stream.push(null);

  const res = await drive.files.create({
    requestBody: {
      name:    nombre,
      parents: [folderId],
    },
    media: {
      mimeType: "image/jpeg",
      body:     stream,
    },
    fields: "id, webViewLink",
  });

  // Hacer el archivo público (solo lectura)
  await drive.permissions.create({
    fileId:      res.data.id,
    requestBody: { role: "reader", type: "anyone" },
  });

  return {
    id:  res.data.id,
    url: res.data.webViewLink,
    nombre,
  };
}

// Subir Excel a Google Drive (reemplaza si existe)
async function subirExcel(rutaLocal, nombre) {
  const drive    = getDrive();
  const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;
  const fs       = require("fs");

  // Buscar si ya existe
  const existing = await drive.files.list({
    q: `name='${nombre}' and '${folderId}' in parents and trashed=false`,
    fields: "files(id)",
  });

  const { Readable } = require("stream");
  const buffer = fs.readFileSync(rutaLocal);
  const stream = new Readable();
  stream.push(buffer);
  stream.push(null);

  const media = {
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    body: stream,
  };

  let fileId;
  if (existing.data.files.length > 0) {
    // Actualizar existente
    fileId = existing.data.files[0].id;
    await drive.files.update({ fileId, media });
  } else {
    // Crear nuevo
    const res = await drive.files.create({
      requestBody: { name: nombre, parents: [folderId] },
      media,
      fields: "id",
    });
    fileId = res.data.id;
    await drive.permissions.create({
      fileId,
      requestBody: { role: "reader", type: "anyone" },
    });
  }

  const file = await drive.files.get({ fileId, fields: "webViewLink" });
  return file.data.webViewLink;
}

module.exports = { subirFoto, subirExcel };
