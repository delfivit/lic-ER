#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Avisos por mail + salida CSV para Google Sheets.

Las credenciales NUNCA van en el código: se leen de variables de entorno
(en GitHub Actions, de los Secrets del repo; en tu Mac, del archivo .env).
    MAIL_USUARIO   tu-cuenta@gmail.com
    MAIL_PASSWORD  la contraseña de aplicación de Google (16 letras)
    MAIL_PARA      destinatarios separados por coma
"""
import os, csv, smtplib, html
from email.message import EmailMessage
from pathlib import Path

BASE = Path(__file__).parent

def cargar_env():
    """Lee .env si existe (para correr en la Mac). En GitHub usa los Secrets."""
    f = BASE / '.env'
    if not f.exists(): return
    for linea in f.read_text(encoding='utf-8').splitlines():
        linea = linea.strip()
        if not linea or linea.startswith('#') or '=' not in linea: continue
        k, v = linea.split('=', 1)
        os.environ.setdefault(k.strip(), v.strip().strip('"\''))

# ------------------------------------------------------- CSV para Google Sheets
def escribir_csv(filas, cols, ruta=None):
    """La Google Sheet lo lee con =IMPORTDATA(...) — sin credenciales de Google."""
    ruta = ruta or (BASE / 'licitaciones.csv')
    visibles = list(cols)   # incluye el ID: lo necesita el rescate de fuentes caídas
    with open(ruta, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(visibles)
        for fila in filas:
            w.writerow([str(fila.get(c, '') or '').replace('\n', ' ') for c in visibles])
    return ruta

# ------------------------------------------------------------------- mail
def mail_html(nuevas, cols, total_vigentes, alertas):
    e = lambda s: html.escape(str(s or ''))
    td = 'style="padding:6px 9px;border:1px solid #ddd;font-size:13px"'
    filas = ''
    for f in nuevas:
        links = []
        if f.get('Link'):       links.append(f'<a href="{e(f["Link"])}">ver</a>')
        if f.get('Pliego PDF'): links.append(f'<a href="{e(f["Pliego PDF"])}">PDF</a>')
        filas += (
            f'<tr><td {td}><b>{e(f.get("Rubro"))}</b></td>'
            f'<td {td}>{e(f.get("Organismo"))[:60]}</td>'
            f'<td {td}>{e(f.get("N°"))}</td>'
            f'<td {td}>{e(f.get("Objeto"))[:150]}</td>'
            f'<td {td}>{e(f.get("Apertura"))[:60]}</td>'
            f'<td {td}><b>{e(f.get("Venta hasta"))}</b></td>'
            f'<td {td}>{e(f.get("Valor pliego"))[:40]}</td>'
            f'<td {td}>{" · ".join(links)}</td></tr>')
    cabecera = ''.join(f'<th style="padding:6px 9px;text-align:left">{c}</th>' for c in
                       ('Rubro','Organismo','N°','Objeto','Apertura','Venta hasta','Valor pliego',''))
    cuerpo = (f'<p style="font-family:Arial">Hola Delfi, hay <b>{len(nuevas)}</b> '
              f'licitación(es) nueva(s) de tu interés en Entre Ríos '
              f'(de {total_vigentes} vigentes en total):</p>'
              f'<table style="border-collapse:collapse;font-family:Arial">'
              f'<tr style="background:#1F3A5F;color:#fff">{cabecera}</tr>{filas}</table>'
              ) if nuevas else (
              f'<p style="font-family:Arial">Hola Delfi, hoy no hubo licitaciones nuevas. '
              f'Seguís con <b>{total_vigentes}</b> vigentes.</p>')
    if alertas:
        cuerpo += ('<p style="margin-top:16px;padding:10px;background:#FCE4E4;'
                   'border-left:4px solid #C00;font-family:Arial;font-size:13px">'
                   f'<b>Fuentes con problemas ({len(alertas)})</b><br>'
                   + '<br>'.join(e(a) for a in alertas) + '</p>')
    cuerpo += ('<p style="color:#888;font-size:12px;font-family:Arial">'
               'Generado automáticamente. El Excel y la planilla ya están actualizados.</p>')
    return cuerpo

def enviar(nuevas, cols, total_vigentes, alertas, adjunto=None):
    cargar_env()
    usuario = os.environ.get('MAIL_USUARIO', '').strip()
    clave   = os.environ.get('MAIL_PASSWORD', '').strip()
    para    = [x.strip() for x in os.environ.get('MAIL_PARA', usuario).split(',') if x.strip()]
    if not (usuario and clave and para):
        print('  (mail no configurado: falta MAIL_USUARIO / MAIL_PASSWORD / MAIL_PARA)')
        return False
    if not nuevas and not alertas:
        print('  (sin novedades ni alertas: no se manda mail)')
        return False

    msg = EmailMessage()
    msg['Subject'] = (f'🏗️ {len(nuevas)} licitación(es) nueva(s) — Entre Ríos'
                      if nuevas else '⚠️ Licitaciones ER — fuentes con problemas')
    msg['From'], msg['To'] = usuario, ', '.join(para)
    msg.set_content('Abrí este mail en formato HTML para ver la tabla.')
    msg.add_alternative(mail_html(nuevas, cols, total_vigentes, alertas), subtype='html')
    if adjunto and Path(adjunto).exists():
        datos = Path(adjunto).read_bytes()
        msg.add_attachment(datos, maintype='application', filename=Path(adjunto).name,
                           subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    try:
        with smtplib.SMTP('smtp.gmail.com', 587, timeout=30) as sm:
            sm.starttls(); sm.login(usuario, clave); sm.send_message(msg)
        print(f'  mail enviado a {", ".join(para)}')
        return True
    except Exception as ex:
        print(f'  ERROR al mandar el mail: {type(ex).__name__}: {ex}')
        return False


# --------------------------------------------- notificación de macOS (sin claves)
def avisar_mac(nuevas, total_vigentes):
    """Muestra una notificación en la Mac. No necesita ninguna contraseña."""
    import subprocess, shlex
    if nuevas:
        rubros = {}
        for f in nuevas:
            rubros[f.get('Rubro', '?')] = rubros.get(f.get('Rubro', '?'), 0) + 1
        detalle = ', '.join(f'{v} de {k}' for k, v in sorted(rubros.items(), key=lambda x: -x[1]))
        titulo = f'{len(nuevas)} licitación(es) nueva(s)'
        texto = f'{detalle}. En total tenés {total_vigentes} vigentes.'
    else:
        titulo = 'Sin licitaciones nuevas'
        texto = f'Seguís con {total_vigentes} vigentes.'
    guion = (f'display notification {shlex.quote(texto)} '
             f'with title "Licitaciones Entre Ríos" '
             f'subtitle {shlex.quote(titulo)} sound name "Glass"')
    try:
        subprocess.run(['osascript', '-e', guion], timeout=15, capture_output=True)
        return True
    except Exception:
        return False
