const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
const { execSync } = require('child_process');

// Cargar variables de entorno si existen
require('dotenv').config({ path: path.join(__dirname, '../.env') });

const SMTP_HOST = process.env.SMTP_HOST || 'mail.mascampobiotecnologia.com';
const SMTP_PORT = parseInt(process.env.SMTP_PORT || '465');
const SMTP_USER = process.env.SMTP_USER || '';
const SMTP_PASS = process.env.SMTP_PASS || '';
const BACKUP_TO_EMAIL = process.env.BACKUP_TO_EMAIL || SMTP_USER;

const ROOT_DIR = path.join(__dirname, '../../');
const BACKUPS_DIR = path.join(__dirname, '../backups');

if (!fs.existsSync(BACKUPS_DIR)) {
  fs.mkdirSync(BACKUPS_DIR, { recursive: true });
}

async function runBackup() {
  const dateStr = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 10);
  const zipName = `backup_inventarios_${dateStr}.zip`;
  const zipPath = path.join(BACKUPS_DIR, zipName);

  console.log(`📦 [BACKUP] Generando archivo comprimido: ${zipName}...`);

  try {
    // Comprimir archivos de datos y logs usando zip de Linux o PowerShell
    const command = `zip -r "${zipPath}" "${path.join(ROOT_DIR, 'all_excel_customers.json')}" "${path.join(ROOT_DIR, 'all_excel_products.json')}" "${path.join(ROOT_DIR, 'extracted_customers.json')}" "${path.join(ROOT_DIR, 'backend/logs')}" -j`;
    execSync(command, { stdio: 'inherit' });

    console.log(`✅ [BACKUP] Archivo generado exitosamente en ${zipPath}`);

    if (!SMTP_USER || !SMTP_PASS) {
      console.log('⚠️ [BACKUP] SMTP_USER o SMTP_PASS no están configurados en .env. El archivo se guardó localmente pero no se envió por correo.');
      return;
    }

    console.log(`📧 [BACKUP] Enviando copia por correo a ${BACKUP_TO_EMAIL}...`);

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
      from: `"Sistema Inventarios" <${SMTP_USER}>`,
      to: BACKUP_TO_EMAIL,
      subject: `📦 Copia de Seguridad Diaria - Inventarios Más Campo (${dateStr})`,
      text: `Hola,\n\nSe adjunta la copia de seguridad diaria del Sistema de Inventarios Más Campo correspondiente a la fecha ${dateStr}.\n\nSaludos.`,
      attachments: [
        {
          filename: zipName,
          path: zipPath,
        },
      ],
    };

    const info = await transporter.sendMail(mailOptions);
    console.log('🚀 [BACKUP] Correo enviado exitosamente:', info.messageId);

  } catch (err) {
    console.error('❌ [BACKUP] Error durante la copia de seguridad:', err);
  }
}

runBackup();
