#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BOLETÍN OFICIAL DE ENTRE RÍOS — parser de la sección LICITACIONES.

Es la fuente más completa que hay: por ley TODOS los organismos y municipios
publican acá, incluidas las localidades chicas que no tienen web (Villa Elisa,
Seguí, Villaguay...). Además es la única que trae VENTA DE PLIEGOS y VALOR DEL
PLIEGO, que ninguna web publica.

El PDF diario está en una URL predecible:
  https://www.entrerios.gov.ar/boletin/calendario/Boletin/2026/Septiembre/04-09-26.pdf
"""
import re, json
from datetime import date, timedelta
from pathlib import Path
import requests
from pypdf import PdfReader

BASE = Path(__file__).parent
CACHE = BASE / '.cache_boletin'
MESES_URL = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto',
             'Septiembre','Octubre','Noviembre','Diciembre']
UA = ('Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 '
      '(KHTML, like Gecko) Chrome/124.0 Safari/537.36')

def url_boletin(d):
    return (f'https://www.entrerios.gov.ar/boletin/calendario/Boletin/'
            f'{d.year}/{MESES_URL[d.month-1]}/{d:%d-%m-%y}.pdf')

def desdoblar(s):
    """El PDF renderiza la negrita duplicando cada letra:
       'GGOOBBIIEERRNNOO DDEE EENNTTRREE RRÍÍOOSS' -> 'GOBIERNO DE ENTRE RÍOS'"""
    def fix(m):
        w = m.group(0)
        if len(w) >= 4 and len(w) % 2 == 0 and all(w[i] == w[i+1] for i in range(0, len(w)-1, 2)):
            return w[::2]
        return w
    return re.sub(r'\S+', fix, s)

def texto_de_pdf(pdf_bytes):
    import io
    r = PdfReader(io.BytesIO(pdf_bytes))
    t = '\n'.join((p.extract_text() or '') for p in r.pages)
    t = re.sub(r'[ \t]+', ' ', t)
    # los pies de página parten los avisos al medio
    t = re.sub(r'\n?\s*P\s?ar\s?an\s?á[^\n]{0,80}BOLETIN OFICIAL[^\n]*\n?', '\n', t)
    t = re.sub(r'\n?\s*\d{1,3} BOLETIN OFICIAL[^\n]*\n?', '\n', t)
    return desdoblar(t)

def bajar(d, sesion, forzar=False):
    """Devuelve el texto del boletín de esa fecha, o None. Cachea en disco."""
    CACHE.mkdir(exist_ok=True)
    cache = CACHE / f'{d:%Y-%m-%d}.txt'
    if cache.exists() and not forzar:
        return cache.read_text(encoding='utf-8')
    try:
        r = sesion.get(url_boletin(d), timeout=(10, 40))
        if r.status_code != 200 or not r.content.startswith(b'%PDF'):
            return None
        t = texto_de_pdf(r.content)
        cache.write_text(t, encoding='utf-8')
        return t
    except Exception:
        return None

# --------------------------------------------------------------- parseo
FIN_AVISO = re.compile(r'\s*-{2,}(?:\s*-{2,})+\s*')          # " -- -- -- "
ID_AVISO  = re.compile(r'ID:\s*(\d+)\s*-\s*([A-Z.\-\d]+)')
CAMPOS = {
    'objeto':   r'(?:OBJETO|Objeto|OBRA|Obra)\s*:\s*(.+?)(?=\n\s*[A-ZÁÉÍÓÚ][A-ZÁÉÍÓÚ \.]{3,}\s*:|\Z)',
    'apertura': r'(?:APERTURA(?:\s+DE\s+(?:SOBRES|OFERTAS))?|Apertura(?:\s+de\s+Sobres)?)\s*:?\s*(.+?)(?=\n\s*[A-ZÁÉÍÓÚ][A-ZÁÉÍÓÚ \.]{3,}\s*:|\Z)',
    'venta':    r'(?:VENTA\s+DE\s+PLIEGOS?|ADQUISICI[OÓ]N\s+DE\s+PLIEGOS?|RETIRO\s+DE\s+PLIEGOS?)\s*:?\s*(.+?)(?=\n\s*[A-ZÁÉÍÓÚ][A-ZÁÉÍÓÚ \.]{3,}\s*:|\Z)',
    'valor':    r'(?:VALOR\s+DEL\s+PLIEGO|PRECIO\s+DEL\s+PLIEGO|Valor\s+del\s+Pliego)\s*:?\s*(.+?)(?=\n|\Z)',
    'presup':   r'PRESUPUESTO\s+OFICIAL\s*:?\s*(.+?)(?=\n|\Z)',
}
RE_TIPO_NUM = re.compile(
    r'(Licitaci[oó]n\s+P[uú]blica|Licitaci[oó]n\s+Privada|Licitaci[oó]n|Concurso\s+de\s+Precios|'
    r'Concurso\s+P[uú]blico|Contrataci[oó]n\s+Directa|Compulsa\s+Abreviada|Remate)'
    r'\s*(?:N[°ºo\.]*\s*)?([\dA-Z]+\s*[-/]\s*\d{2,4})?', re.I)

def limpiar(s, maxlen=600):
    s = re.sub(r'\s+', ' ', s or '').strip(' .-–—')
    return s[:maxlen]

def secciones_licitaciones(texto):
    """Recorta el texto a la(s) zona(s) donde están los avisos de licitación."""
    marcas = [m.end() for m in re.finditer(r'\n\s*LICITACIONES\s*\n', texto)]
    if not marcas:
        return []
    # se descarta el índice del principio: la sección real es la última marca
    ini = marcas[-1]
    fin = len(texto)
    for corte in (r'\n\s*COMUNICADOS\s*\n', r'\n\s*CONVOCATORIAS\s*\n', r'\n\s*EDICTOS\s*\n'):
        m = re.search(corte, texto[ini:])
        if m:
            fin = min(fin, ini + m.start())
    return [texto[ini:fin]]

def parsear(texto, fecha_bol):
    """Devuelve la lista de avisos de licitación del boletín."""
    out = []
    for sec in secciones_licitaciones(texto):
        # Cada aviso TERMINA con su "ID:nnnnn - ...". Partir por ahí es fiable;
        # el separador "-- -- --" no aparece en todos y mezclaba avisos distintos
        # (el de ENERSA se quedaba con la fecha de apertura de Villa Elisa).
        cortes = [m.end() for m in re.finditer(r'ID:\s*\d+\s*-\s*[^\n]*', sec)]
        bloques, ant = [], 0
        for c in cortes:
            bloques.append(sec[ant:c]); ant = c
        if ant < len(sec) - 60:
            bloques.append(sec[ant:])
        for bloque in bloques:
            bloque = FIN_AVISO.sub('\n', bloque)
            if len(bloque) < 90 or not re.search(r'licitaci|concurso|compulsa|contrataci', bloque, re.I):
                continue
            lineas = [l.strip() for l in bloque.split('\n') if l.strip()]
            if not lineas: continue

            mid = ID_AVISO.search(bloque)
            # Organismo y localidad: las líneas en MAYÚSCULAS antes del tipo.
            # OJO: el título de la obra también viene en mayúsculas y se colaba
            # como organismo ("CUBIERTA DE TECHOS, REPARACION DE MAMPOSTERIAS").
            # Sólo aceptamos nombres de organismo o localidades cortas.
            ES_ORGANISMO = re.compile(
                r'^\s*(municipalidad|municipio|comuna|gobierno|secretar|direcci[oó]n|'
                r'instituto|ente|comisi[oó]n|consejo|ministerio|junta|tribunal|'
                r'universidad|hospital|caja|banco|empresa|administraci[oó]n|'
                r'unidad|departamento ejecutivo|honorable)', re.I)
            ES_OBRA = re.compile(
                r'(construcci|refacci|remodelaci|ampliaci|reparaci|provisi|adquisici|'
                r'cubierta|mamposter|pavimento|impermeabiliza|pisos|techos|viviendas|'
                r'apertura|objeto|licitaci)', re.I)
            enc = []
            for l in lineas[:7]:
                if RE_TIPO_NUM.match(l): break
                l = l.strip(' .-—–')
                letras = re.sub(r'[^A-Za-zÁÉÍÓÚÑáéíóúñ]', '', l)
                if not letras or letras != letras.upper() or len(letras) < 3:
                    continue
                if ES_ORGANISMO.match(l):
                    enc.append(l[:80]); continue
                # si no arranca como organismo, sólo vale si es corta y no habla de obra
                if len(l) <= 32 and not ES_OBRA.search(l):
                    enc.append(l)
            organismo = ' — '.join(dict.fromkeys(enc[-2:])) if enc else ''
            localidad = enc[0] if enc else ''

            mt = RE_TIPO_NUM.search(bloque)
            tipo = limpiar(mt.group(1), 40).title() if mt else 'Licitación'
            numero = re.sub(r'\s+', '', mt.group(2)) if (mt and mt.group(2)) else ''

            campos = {}
            for k, pat in CAMPOS.items():
                m = re.search(pat, bloque, re.S)
                campos[k] = limpiar(m.group(1)) if m else ''
            # si el "objeto" arrancó con otro rótulo (LUGAR DE APERTURA:, DESTINO:...)
            # entonces el campo se leyó mal y conviene descartarlo
            if re.match(r'(?:LUGAR|DESTINO|APERTURA|VENTA|VALOR|PRESUPUESTO|CONSULTA|PLAZO)\b',
                        campos['objeto'], re.I):
                campos['objeto'] = ''
            if not campos['objeto']:
                # sin "OBJETO:" explícito: la línea más larga que no sea encabezado
                cands = [l for l in lineas if 40 < len(l) < 400 and not l.isupper()]
                campos['objeto'] = limpiar(max(cands, key=len)) if cands else ''
            if not campos['objeto']:
                continue

            # si el campo APERTURA no trae fecha, buscarla en el resto del aviso
            if campos['apertura'] and not re.search(r'\d{1,2}\s*[/\-.]\s*\d{1,2}|\d{1,2}\s+de\s+\w+',
                                                    campos['apertura'], re.I):
                m2 = re.search(r'apertura[^.;]{0,120}?((?:\d{1,2}\s*[/\-.]\s*\d{1,2}\s*[/\-.]\s*\d{2,4})'
                               r'|(?:\d{1,2}\s+de\s+[a-zA-Zá-úÁ-Ú]+\s+de[l]?\s+(?:a[ñn]o\s+)?\d{4}))',
                               bloque, re.I | re.S)
                if m2: campos['apertura'] = limpiar(m2.group(0), 160)

            out.append(dict(
                organismo=organismo or localidad or 'Boletín Oficial',
                localidad=localidad, tipo=tipo, numero=numero,
                objeto=campos['objeto'], apertura=campos['apertura'],
                venta=campos['venta'], valor_pliego=campos['valor'],
                presupuesto=campos['presup'],
                aviso_id=mid.group(1) if mid else '',
                fecha_boletin=fecha_bol, url=url_boletin(fecha_bol),
            ))
    return out

def recolectar(dias=7, sesion=None, hilos=6):
    """Baja los boletines de los últimos N días y devuelve todos los avisos.

    Los PDFs se bajan en paralelo: son ~2 MB cada uno y en secuencia 20 días
    tardaban más que el límite de tiempo por fuente.
    """
    from concurrent.futures import ThreadPoolExecutor
    import threading
    local = threading.local()

    def sesion_del_hilo():
        # requests.Session no es thread-safe: una por hilo
        if not hasattr(local, 's'):
            local.s = requests.Session()
            local.s.headers.update({'User-Agent': UA})
        return local.s

    fechas = []
    for i in range(dias):
        d = date.today() - timedelta(days=i)
        if d.weekday() < 5:          # el boletín sale días hábiles
            fechas.append(d)

    def uno(d):
        t = bajar(d, sesion_del_hilo())
        return (d, parsear(t, d)) if t else (d, None)

    avisos, bajados = [], []
    with ThreadPoolExecutor(max_workers=hilos) as pool:
        for d, got in pool.map(uno, fechas):
            if got is None:
                continue
            bajados.append((d, len(got)))
            avisos += got
    bajados.sort(reverse=True)
    return avisos, bajados

if __name__ == '__main__':
    import sys
    dias = int(sys.argv[1]) if len(sys.argv) > 1 else 7
    avisos, bajados = recolectar(dias)
    print(f'\nBoletines leídos: {len(bajados)}')
    for d, n in bajados: print(f'   {d:%d/%m/%Y}: {n} avisos')
    print(f'\nTOTAL avisos de licitación: {len(avisos)}\n' + '=' * 76)
    for a in avisos:
        print(f'\n[{a["tipo"]} {a["numero"]}] {a["organismo"][:66]}')
        print(f'   OBJETO : {a["objeto"][:112]}')
        if a['apertura']: print(f'   APERT. : {a["apertura"][:92]}')
        if a['venta']:    print(f'   VENTA  : {a["venta"][:92]}')
        if a['valor_pliego']: print(f'   PLIEGO : {a["valor_pliego"][:70]}')
