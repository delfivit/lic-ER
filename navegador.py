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

_PW = None          # se arranca una sola vez y se reusa
_NAVEGADOR = None

UA = ('Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 '
      '(KHTML, like Gecko) Chrome/124.0 Safari/537.36')

def disponible():
    try:
        import playwright  # noqa: F401
        return True
    except ImportError:
        return False

def _arrancar():
    global _PW, _NAVEGADOR
    if _NAVEGADOR is not None:
        return _NAVEGADOR
    from playwright.sync_api import sync_playwright
    _PW = sync_playwright().start()
    _NAVEGADOR = _PW.chromium.launch(args=['--no-sandbox', '--disable-dev-shm-usage'])
    return _NAVEGADOR

def cerrar():
    global _PW, _NAVEGADOR
    try:
        if _NAVEGADOR: _NAVEGADOR.close()
        if _PW: _PW.stop()
    except Exception:
        pass
    _NAVEGADOR = _PW = None

def bajar(url, espera_ms=2500, timeout_ms=45000):
    """Devuelve el HTML ya renderizado, o None."""
    if not disponible():
        return None
    try:
        nav = _arrancar()
        ctx = nav.new_context(user_agent=UA, locale='es-AR',
                              ignore_https_errors=True,   # cadenas incompletas (ENERSA)
                              viewport={'width': 1400, 'height': 900})
        pg = ctx.new_page()
        try:
            pg.goto(url, wait_until='domcontentloaded', timeout=timeout_ms)
            # si aparece el "Just a moment" de Cloudflare, le damos tiempo a resolverse
            for _ in range(3):
                if not re.search(r'just a moment|verificando|checking your browser',
                                 pg.content()[:3000], re.I):
                    break
                pg.wait_for_timeout(3500)
            pg.wait_for_timeout(espera_ms)
            html = pg.content()
            return html if html and len(html) > 500 else None
        finally:
            ctx.close()
    except Exception:
        return None
