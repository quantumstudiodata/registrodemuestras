const { Readable } = require('stream');
const { getDriveClient } = require('./auth');

// Reemplaza a getOrCreateCarpetaFotos()/guardarPerfil() de Code.gs. La carpeta hay que
// crearla una vez a mano en Drive, compartirla con el correo de la cuenta de servicio
// (rol Editor) y poner su ID aquí vía variable de entorno.
async function subirFotoPerfil(usuario, fotoBase64) {
  const folderId = process.env.DRIVE_FOLDER_FOTOS_ID;
  if (!folderId) throw new Error('Falta la variable de entorno DRIVE_FOLDER_FOTOS_ID.');

  const partes = String(fotoBase64).split(',');
  const meta = partes[0];
  const contenido = partes[1];
  const match = meta.match(/data:(.*);base64/);
  const mime = match ? match[1] : 'image/jpeg';
  const buffer = Buffer.from(contenido, 'base64');

  const drive = await getDriveClient();
  const res = await drive.files.create({
    requestBody: { name: `${usuario}_foto`, parents: [folderId] },
    media: { mimeType: mime, body: Readable.from(buffer) },
    fields: 'id'
  });
  const fileId = res.data.id;

  await drive.permissions.create({
    fileId,
    requestBody: { role: 'reader', type: 'anyone' }
  });

  return `https://drive.google.com/thumbnail?id=${fileId}&sz=w200`;
}

module.exports = { subirFotoPerfil };
