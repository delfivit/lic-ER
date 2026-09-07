#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Bajar páginas con un navegador de verdad (Playwright).

Se usa sólo como reintento, cuando el pedido normal falla. Resuelve dos casos
que aparecieron al correr en la nube:

  · Paraná (compras.parana.gob.ar) devuelve 403 a los servidores de datacenter
    (le pasaba a Google Apps Script y también a GitHub). Un navegador real pasa.
  · ENERSA no manda su certificado intermedio de Sectigo: requests corta con
    SSLError, pero el navegador completa la cadena solo (AIA fetching).
"""
import re

import threading
# OJO: la API sync de Playwright NO es thread-safe y el scraper corre las fuentes
# en paralelo. Compartir un navegador entre hilos rompe con
# "greenlet.error: Cannot switch to a different thread". Por eso cada hilo tiene
# el suyo, guardado en almacenamiento local del hilo.
_local = threading.local()

UA = ('Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 '
      '(KHTML, like Gecko) Chrome/124.0 Safari/537.36')

def disponible():
    try:
        import playwright  # noqa: F401
        return True
    except ImportError:
        return False

def _arrancar():
    nav = getattr(_local, 'navegador', None)
    if nav is not None:
        return nav
    from playwright.sync_api import sync_playwright
    _local.pw = sync_playwright().start()
    _local.navegador = _local.pw.chromium.launch(
        args=['--no-sandbox', '--disable-dev-shm-usage'])
    return _local.navegador

def cerrar():
    """Cierra el navegador de ESTE hilo (cada uno cierra el suyo)."""
    try:
        if getattr(_local, 'navegador', None): _local.navegador.close()
        if getattr(_local, 'pw', None): _local.pw.stop()
    except Exception:
        pass
    _local.navegador = _local.pw = None

def bajar(url, espera_ms=2500, timeout_ms=60000, esperar=None, diagnostico=False):
    """Devuelve el HTML ya renderizado, o None.

    `esperar` es un selector CSS que tiene que aparecer antes de leer la página.
    Es más confiable que esperar un tiempo fijo: en los servidores de GitHub la
    tabla de Paraná tarda más en armarse que en una máquina local.
    """
    if not disponible():
        return None
    try:
        nav = _arrancar()
        ctx = nav.new_context(user_agent=UA, locale='es-AR',
                              ignore_https_errors=True,   # cadenas incompletas (ENERSA)
                              viewport={'width': 1400, 'height': 900},
                              extra_http_headers={'Accept-Language': 'es-AR,es;q=0.9'})
        pg = ctx.new_page()
        try:
            pg.goto(url, wait_until='domcontentloaded', timeout=timeout_ms)
            # si aparece el "Just a moment" de Cloudflare, le damos tiempo a resolverse
            for _ in range(8):
                if not re.search(r'just a moment|un momento|aguarde|verificando|checking your browser|attention required|espere',
                                 pg.content()[:4000], re.I):
                    break
                pg.wait_for_timeout(3000)
            if esperar:
                try:
                    pg.wait_for_selector(esperar, timeout=25000, state='attached')
                except Exception:
                    if diagnostico:
                        _informar(pg, url, esperar)
            pg.wait_for_timeout(espera_ms)
            html = pg.content()
            if diagnostico and esperar:
                print(f'    [navegador] {url} -> {len(html)} bytes, '
                      f'"{esperar}": {len(pg.query_selector_all(esperar))} coincidencias')
            return html if html and len(html) > 500 else None
        finally:
            ctx.close()
    except Exception as e:
        if diagnostico:
            print(f'    [navegador] falló {url}: {type(e).__name__}: {e}'[:200])
        return None

def _informar(pg, url, esperar):
    """Cuando no aparece lo que esperábamos, deja pistas en el log."""
    try:
        cont = pg.content()
        print(f'    [navegador] NO apareció "{esperar}" en {url}')
        print(f'      titulo: {pg.title()[:70]}')
        print(f'      bytes : {len(cont)}')
        print(f'      filas <tr>: {len(pg.query_selector_all("tr"))}')
        if re.search(r'just a moment|attention required|cloudflare|blocked|denied', cont, re.I):
            print('      >>> parece un bloqueo de Cloudflare')
        texto = re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', cont))[:220]
        print(f'      texto : {texto}')
    except Exception:
        pass
