#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de Copia de Seguridad Automática Más Campo
Comprime mascampo.db y el catálogo, y lo envía por correo electrónico vía SMTP seguro.
Usa librerías nativas estándar de Python: smtplib, zipfile, email, ssl.
"""

import os
import sys
import zipfile
import smtplib
import ssl
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ROOT_DIR = os.path.dirname(BASE_DIR)
BACKUPS_DIR = os.path.join(BASE_DIR, 'backups')
ENV_PATH = os.path.join(BASE_DIR, '.env')

os.makedirs(BACKUPS_DIR, exist_ok=True)

# 1. Cargar variables de entorno desde .env manualmente sin librerías externas
env_vars = {}
if os.path.exists(ENV_PATH):
    with open(ENV_PATH, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith('#') and '=' in line:
                k, v = line.split('=', 1)
                env_vars[k.strip()] = v.strip().strip('"').strip("'")

SMTP_HOST = env_vars.get('SMTP_HOST') or os.environ.get('SMTP_HOST') or 'smtp.gmail.com'
SMTP_PORT = int(env_vars.get('SMTP_PORT') or os.environ.get('SMTP_PORT') or 465)
SMTP_USER = env_vars.get('SMTP_USER') or os.environ.get('SMTP_USER') or 'gerencia@softproductiva.com'
SMTP_PASS = env_vars.get('SMTP_PASS') or os.environ.get('SMTP_PASS') or ''
BACKUP_TO = env_vars.get('BACKUP_TO_EMAIL') or os.environ.get('BACKUP_TO_EMAIL') or 'gerencia@softproductiva.com'

now = datetime.now()
date_str = now.strftime('%Y-%m-%d_%H%M%S')
timestamp_str = now.strftime('%Y-%m-%d %H:%M:%S')

zip_filename = f"backup_inventarios_{date_str}.zip"
zip_path = os.path.join(BACKUPS_DIR, zip_filename)

print("=" * 60)
print(f"[BACKUP DIARIO - {timestamp_str}] Iniciando proceso...")
print(f"Ruta .env:       {ENV_PATH}")
print(f"Remitente SMTP:  {SMTP_USER}")
print(f"Destinatario:    {BACKUP_TO}")

# 2. Comprimir base de datos SQLite y archivos maestros
files_to_pack = []
db_path = os.path.join(BASE_DIR, 'mascampo.db')
if os.path.exists(db_path):
    files_to_pack.append((db_path, 'mascampo.db'))

for fname in ['all_new_products.json', 'backend/initial_products.json', 'backend/initial_customers.json', 'backend/initial_categories.json']:
    fpath = os.path.join(ROOT_DIR, fname) if not os.path.isabs(fname) else fname
    if os.path.exists(fpath):
        files_to_pack.append((fpath, os.path.basename(fname)))

try:
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file_path, arcname in files_to_pack:
            zf.write(file_path, arcname)
    file_size_kb = round(os.path.getsize(zip_path) / 1024, 2)
    print(f"[OK] Archivo ZIP generado: {zip_filename} ({file_size_kb} KB)")
except Exception as e:
    print(f"[ERROR] Error al generar archivo ZIP: {e}")
    sys.exit(1)

# 3. Validar contraseña
if not SMTP_PASS:
    print("[AVISO] SMTP_PASS no esta configurado en backend/.env. Copia guardada localmente en backups/.")
    print("=" * 60)
    sys.exit(0)

# 4. Enviar correo electrónico vía SMTP
try:
    msg = MIMEMultipart()
    msg['From'] = f"Sistema Inventarios Mas Campo <{SMTP_USER}>"
    msg['To'] = BACKUP_TO
    msg['Subject'] = f"Copia de Seguridad Diaria - Inventarios Mas Campo ({now.strftime('%d/%m/%Y')})"

    body = f"Hola,\n\nSe adjunta la copia de seguridad fisica de la base de datos (SQLite mascampo.db) y catalogo del Sistema de Inventarios Mas Campo.\n\nFecha: {timestamp_str}\nArchivo: {zip_filename}\nTamano: {file_size_kb} KB\n\nSaludos cordiales,\nSistema Automatizado de Inventarios Mas Campo.\n"
    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    with open(zip_path, 'rb') as f:
        part = MIMEBase('application', 'zip')
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{zip_filename}"')
        msg.attach(part)

    if SMTP_PORT == 465:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as server:
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(SMTP_USER, [BACKUP_TO], msg.as_string())
    else:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(SMTP_USER, [BACKUP_TO], msg.as_string())

    print(f"[EXITO] Copia enviada por correo a {BACKUP_TO}.")
    print("=" * 60)
except Exception as e:
    print(f"[ERROR SMTP]: {e}")
    print("=" * 60)
    sys.exit(1)
