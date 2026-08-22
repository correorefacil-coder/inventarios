const fs = require('fs');
const path = require('path');

// Cargar variables de entorno obligatoriamente desde la ruta del script
const envPath = path.resolve(__dirname, '../.env');
require('dotenv').config({ path: envPath });

const nodemailer = require('nodemailer');
const { execSync } = require('child_process');

const SMTP_HOST = process.env.SMTP_HOST || 'smtp.gmail.com';
const SMTP_PORT = parseInt(process.env.SMTP_PORT || '465');
const SMTP_USER = process.env.SMTP_USER || '';
const SMTP_PASS = process.env.SMTP_PASS || '';
const BACKUP_TO_EMAIL = process.env.BACKUP_TO_EMAIL || 'inventarios_mascampo@softproductiva.com';

const ROOT_DIR = path.resolve(__dirname, '../../');
const BACKUPS_DIR = path.resolve(__dirname, '../backups');

if (!fs.existsSync(BACKUPS_DIR)) {
  fs.mkdirSync(BACKUPS_DIR, { recursive: true });
}

async function runBackup() {
  const now = new Date();
  const dateStr = now.toISOString().replace(/[:.]/g, '-').slice(0, 10);
  const timestamp = now.toLocaleString('es-CO', { timeZone: 'America/Bogota' });
  const zipName = `backup_inventarios_${dateStr}.zip`;
  const zipPath = path.join(BACKUPS_DIR, zipName);

  console.log(`\n==================================================`);
  console.log(`📦 [BACKUP DIARIO - ${timestamp}] Iniciando proceso...`);
  console.log(`📂 Ruta .env cargada: ${envPath}`);
  console.log(`📧 Usuario SMTP: ${SMTP_USER || 'NO CONFIGURADO'}`);

  try {
    // Archivos a incluir en la copia
    const custJson = path.join(ROOT_DIR, 'all_excel_customers.json');
    const prodJson = path.join(ROOT_DIR, 'all_excel_products.json');
    const extCustJson = path.join(ROOT_DIR, 'extracted_customers.json');
    const logsDir = path.join(ROOT_DIR, 'backend/logs');

    // Intentar comprimir usando zip de Linux
    try {
      const command = `zip -j -r "${zipPath}" "${custJson}" "${prodJson}" "${extCustJson}" "${logsDir}"`;
      execSync(command, { stdio: 'pipe' });
      console.log(`✅ [BACKUP] Archivo comprimido generado en: ${zipPath}`);
    } catch (zipErr) {
      console.warn(`⚠️ [BACKUP] Warning en ejecucion de zip system command: ${zipErr.message}`);
      // Fallback: Si no hay comando zip, copiar el archivo principal de productos como respaldo
      fs.copyFileSync(prodJson, zipPath + '.json');
    }

    const fileToSend = fs.existsSync(zipPath) ? zipPath : zipPath + '.json';
    const fileNameToSend = path.basename(fileToSend);

    if (!SMTP_USER || !SMTP_PASS) {
      console.log('⚠️ [BACKUP] SMTP_USER o SMTP_PASS no están configurados en .env. El backup no pudo enviarse.');
      return;
    }

    console.log(`📧 [BACKUP] Enviando correo a: ${BACKUP_TO_EMAIL}...`);

    const transporter = nodemailer.createTransport({
      host: SMTP_HOST,
      port: SMTP_PORT,
      secure: SMTP_PORT === 465,
      auth: {
        user: SMTP_USER,
        pass: SMTP_PASS,
      },
    });

    const mailOptions = {
      from: `"Sistema Inventarios Más Campo" <${SMTP_USER}>`,
      to: BACKUP_TO_EMAIL,
      subject: `📦 Copia de Seguridad Diaria - Inventarios Más Campo (${dateStr})`,
      text: `Hola,\n\nSe adjunta la copia de seguridad diaria del Sistema de Inventarios Más Campo correspondiente a la fecha: ${timestamp}.\n\nSaludos,\nSistema Automatizado de Inventarios.`,
      attachments: [
        {
          filename: fileNameToSend,
          path: fileToSend,
        },
      ],
    };

    const info = await transporter.sendMail(mailOptions);
    console.log(`🚀 [BACKUP OK - ${timestamp}] Correo enviado exitosamente. ID: ${info.messageId}`);

  } catch (err) {
    console.error(`❌ [BACKUP ERROR - ${timestamp}] Error durante la copia de seguridad:`, err);
  }
}

runBackup();
