#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LICITACIONES ENTRE RÍOS — scraper para Granss SRL
Rubros: Arquitectura · Infraestructura · Áridos · Vehículos

Uso:
    ./correr.sh                # baja todo y actualiza el Excel
    ./correr.sh --solo NOMBRE  # una sola fuente (para probar)
"""
import json, re, os, sys, time, unicodedata, hashlib, argparse
from datetime import datetime, date
from urllib.parse import quote
from pathlib import Path
import requests
from bs4 import BeautifulSoup

BASE = Path(__file__).parent
HOY = date.today()
BOLETIN_DIAS = 25   # cuántos días atrás mirar del Boletín Oficial
EN_LA_NUBE = bool(os.environ.get('GITHUB_ACTIONS'))   # ¿corriendo en GitHub?
TODAS_LAS_FUENTES = []
# La hoja "Fuentes" de la planilla de Delfina. Lo que agregue ahí se scrapea
# al día siguiente sin tocar código. Si no se puede leer, se usa fuentes.json.
PLANILLA_FUENTES = ('https://docs.google.com/spreadsheets/d/'
                    '1gFFecC4oR189HbYa04sAzm7Q6tTN3UyRvrXiM_8CeFc/gviz/tq?tqx=out:csv&sheet=Fuentes')


def fuentes_de_la_planilla():
    """Lee las fuentes que el equipo cargó en la planilla. Devuelve [] si no se
    puede (sin internet, planilla movida, hoja renombrada): en ese caso se sigue
    con las del archivo, así una planilla rota nunca deja al scraper sin fuentes."""
    import csv, io
    try:
        r = requests.get(PLANILLA_FUENTES, headers=HEADERS, timeout=(10, 25))
        if r.status_code != 200:
            return []
        r.encoding = 'utf-8'
        filas = list(csv.DictReader(io.StringIO(r.text)))
    except Exception as e:
        print(f'  (no pude leer las fuentes de la planilla: {type(e).__name__})')
        return []
    out = []
    for f in filas:
        cols = {(k or '').strip().lower(): (v or '').strip() for k, v in f.items()}
        url = cols.get('url', '')
        nombre = cols.get('nombre', '')
        if not url.startswith('http') or not nombre:
            continue
        activa = cols.get('activa', '').upper() in ('SI', 'SÍ', 'X', 'TRUE', 'VERDADERO', '1', '✓')
        out.append({'activa': activa, 'nombre': nombre, 'url': url,
                    'parser': (cols.get('parser') or 'generico').lower(),
                    'nota': cols.get('notas', '')})
    return out
UA = ('Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 '
      '(KHTML, like Gecko) Chrome/124.0 Safari/537.36')
# Proxies de lectura para sitios que bloquean datacenters (a Apps Script Paraná
# le devolvía 403). Son servicios gratuitos y se caen seguido — se prueban en
# cadena y es un último intento, no una garantía. Desde tu Mac no hacen falta.
PROXIES = ['https://api.allorigins.win/raw?url=',
           'https://api.codetabs.com/v1/proxy?quest=']
HEADERS = {'User-Agent': UA, 'Accept-Language': 'es-AR,es;q=0.9',
           'Accept': 'text/html,application/xhtml+xml,application/json;q=0.9,*/*;q=0.8'}

# ---------------------------------------------------------------- utilidades
def sin_acentos(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s or '')
                   if unicodedata.category(c) != 'Mn')

def norm(s):
    return re.sub(r'\s+', ' ', (s or '')).strip()

MESES = {'enero':1,'febrero':2,'marzo':3,'abril':4,'mayo':5,'junio':6,'julio':7,
         'agosto':8,'septiembre':9,'setiembre':9,'octubre':10,'noviembre':11,'diciembre':12}

def parse_fecha(s):
    """Acepta 11-09-2026, 23/09/26, 5.9.2026 y '11 de octubre de 2026'."""
    s = str(s or '')
    m = re.search(r'(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{2,4})', s)
    if m:
        d, mes, a = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if a < 100: a += 2000
        try: return date(a, mes, d)
        except ValueError: return None
    # "11 de octubre de 2026", "24 DE SEPTIEMBRE DEL AÑO 2026", "5 de mayo 2026"
    m = re.search(r'(\d{1,2})\s+de\s+([a-z]+)\s+(?:de[l]?\s+)?(?:a[nñ]o\s+)?(\d{4})',
                  sin_acentos(s.lower()))
    if m and m.group(2) in MESES:
        try: return date(int(m.group(3)), MESES[m.group(2)], int(m.group(1)))
        except ValueError: return None
    return None

def estado_de(f):
    if not f: return 'Revisar fecha'
    return 'Vigente' if f >= HOY else 'Vencida'

def uid_hash(s):
    return 'id-' + hashlib.md5(sin_acentos(s.lower()).encode()).hexdigest()[:16]

def limpiar_apertura(txt):
    """El texto de apertura viene con arrastres del boletín
    ('DE OFERTAS: el día 30 de Septiembre...'). Lo dejamos legible."""
    t = norm(txt)
    t = re.sub(r'^(?:DE\s+)?(?:OFERTAS?|SOBRES?|PROPUESTAS?)\s*:?\s*', '', t, flags=re.I)
    t = re.sub(r'^(?:el\s+)?d[ií]a\s+', '', t, flags=re.I)
    t = re.sub(r'\s*ID:\s*\d+.*$', '', t)              # cola del aviso del boletín
    t = re.sub(r'\s*-\s*SI ES DECRETADO INHABIL.*$', '', t, flags=re.I)
    return t[:150].strip(' .-–—')

# ---------------------------------------------------------------- clasificador
RUIDO = ['servicio de limpieza','corte de pasto','desmalezamiento','desmalezado','vigilancia',
    'seguridad privada','catering','comedor','indumentaria','calzado','uniforme','papeler',
    'libreria','toner','cartucho','fotocopiadora','impresora','insumos informatic',
    'equipamiento informatic','telefonia','protesis','material descartable','odontolog',
    'farmac','medicamento','poliza de seguro','alquiler de inmueble','repuesto','neumatic',
    'heladera','aire acondicionado','mobiliario','venta de lotes','lotes de terreno',
    'venta de terreno','servidor de base de datos','locacion de servicio']

RUBROS = [
    ('Vehículos', ['camioneta','camion','vehiculo','pick up','pick-up','pickup','utilitario',
        'automotor','automovil','minibus','colectivo','furgon','tractor','retroexcavadora',
        'motoniveladora','pala cargadora','cargadora frontal','maquinaria vial','excavadora',
        'compactador','acoplado','chasis','autoelevador']),
    ('Áridos', ['arido','ripio','arena','canto rodado','piedra partida','agregado petreo',
        'suelo seleccionado','broza','triturado petreo','provision de suelo',
        'extraccion de deposito','movimiento de suelo','extraccion de suelo']),
    ('Infraestructura', ['vivienda','hospital','centro de salud','sala de salud','obras sanitarias',
        'cloaca','cloacal','agua potable','desague','desagues','pluvial','red de gas','gas natural',
        'gasoducto','electrificacion','linea de media tension','linea de baja tension',
        'estacion transformadora','subestacion','pavimento','pavimentacion','repavimentacion',
        'bacheo','ruta','camino','puente','obra basica','cordon cuneta','alcantarilla','hidraulica',
        'iluminacion','alumbrado publico','red de distribucion','saneamiento','planta depuradora',
        'acueducto','enripiado','calzada']),
    ('Arquitectura', ['escuela','edificio','edilicia','edilicio','refaccion','remodelacion',
        'ampliacion','construccion','reparaciones generales','reparacion','cubierta de techo',
        'techo','banos','aulas','jardin de infantes','arquitectura','obra civil','mamposteria',
        'impermeabilizacion','instalacion electrica','cenefa','losa','portico','pintura',
        'carpinteria','herreria']),
]

def es_ruido(texto):
    t = sin_acentos((texto or '').lower())
    return any(k in t for k in RUIDO)

def clasificar(texto):
    t = sin_acentos((texto or '').lower())
    for rubro, claves in RUBROS:
        if any(k in t for k in claves):
            return rubro
    return None

# ------------------------------------------------ fecha de venta de pliego
VERBO = re.compile(r'(venta|vender|vende|adquisicion|adquirir|adquieren|retiro|retirar|'
                   r'retirarse|retiran|compra|comprar|expenden|expendio|entrega)')
F_NUM = re.compile(r'(\d{1,2}\s*[/\-.]\s*\d{1,2}\s*[/\-.]\s*\d{2,4})')
F_TXT = re.compile(r'(\d{1,2}\s+de\s+[a-z]+\s+(?:de[l]?\s+)?(?:a[nñ]o\s+)?\d{4})')

def venta_pliego(texto):
    """Busca por ventana alrededor de 'pliego'. Lineal: no puede colgarse."""
    if not texto: return ''
    t = norm(texto)
    plano = sin_acentos(t.lower())
    mejor, desde = '', 0
    while True:
        pos = plano.find('pliego', desde)
        if pos == -1: break
        desde = pos + 6
        ini = max(0, pos - 150)
        ctx = plano[ini:pos + 150]
        if not VERBO.search(ctx): continue
        ih = ctx.rfind('hasta')      # el último "hasta" es la fecha límite
        for trozo in ([ctx[ih:]] if ih != -1 else []) + [ctx]:
            m = F_NUM.search(trozo) or F_TXT.search(trozo)
            if m:
                # devolvemos siempre dd/mm/aaaa: uniforme, ordenable y sin los
                # acentos perdidos al normalizar el texto para buscar
                f = parse_fecha(m.group(1))
                val = f.strftime('%d/%m/%Y') if f else norm(m.group(1))
                if ih != -1: return val
                mejor = mejor or val
    return mejor

# ---------------------------------------------------------------- red
class Red:
    def __init__(self):
        from requests.adapters import HTTPAdapter
        self.s = requests.Session()
        self.s.headers.update(HEADERS)
        # sin reintentos y con timeout (connect, read): DPV se quedaba 18 minutos colgada
        ad = HTTPAdapter(max_retries=0)
        self.s.mount('https://', ad); self.s.mount('http://', ad)
        self.log = {}

    def get(self, url, fuente, json_=False, _proxy=False):
        d = self.log.setdefault(fuente, {'http': '', 'bytes': 0, 'error': '', 'crudos': 0, 'relev': 0})
        pedir = (PROXIES[_proxy - 1] + quote(url, safe='')) if _proxy else url
        try:
            r = self.s.get(pedir, timeout=(10, 25), allow_redirects=True)
            d['http'] = r.status_code
            if r.status_code != 200:
                # Paraná bloquea a los datacenters (a Google le daba 403). Si estamos
                # corriendo en la nube, reintentamos a través de un proxy de lectura.
                if not _proxy and r.status_code in (403, 429, 503):
                    alt = self._con_navegador(url, d, json_, f'{r.status_code}')
                    if alt is not None:
                        return alt
                    for i in range(1, len(PROXIES) + 1):
                        alt = self.get(url, fuente, json_, _proxy=i)
                        if alt is not None:
                            d['http'] = f'{r.status_code} → proxy {i}'
                            d['error'] = ''
                            return alt
                d['error'] = f'HTTP {r.status_code}'
                return None
            if json_:
                d['bytes'] = len(r.content)
                return r.json()
            # encoding: el header manda; si dice ISO-8859-1 (caso IAPV) hay que respetarlo
            if not r.encoding or r.encoding.lower() in ('iso-8859-1', 'latin-1'):
                declarado = re.search(rb'charset\s*=\s*["\']?([\w-]+)', r.content[:4000], re.I)
                if declarado:
                    r.encoding = declarado.group(1).decode('ascii', 'ignore')
                elif r.apparent_encoding:
                    r.encoding = r.apparent_encoding
            txt = r.text
            d['bytes'] = len(txt)
            if re.search(r'just a moment|cf-browser-verification|challenge-platform', txt, re.I):
                if not _proxy:
                    alt = self._con_navegador(url, d, json_, 'Cloudflare')
                    if alt is not None:
                        return alt
                    for i in range(1, len(PROXIES) + 1):
                        alt = self.get(url, fuente, json_, _proxy=i)
                        if alt is not None:
                            d['http'] = f'CF → proxy {i}'; d['error'] = ''
                            return alt
                d['error'] = 'Bloqueado por Cloudflare (necesita navegador real)'
                return None
            return txt
        except Exception as e:
            if not _proxy and 'SSL' in type(e).__name__.upper():
                # ENERSA no manda el certificado intermedio: el navegador sí completa la cadena
                alt = self._con_navegador(url, d, json_, 'SSL')
                if alt is not None:
                    return alt
            d['http'] = 'ERROR'
            d['error'] = f'{type(e).__name__}: {e}'[:120]
            return None

    # qué elemento hay que esperar en cada sitio antes de leer la página
    ESPERAR = {'compras.parana.gob.ar': 'a[href*="/uploads/pliegos/"]'}
    # sitios donde el navegador tampoco sirve cuando corremos en la nube
    SIN_SALIDA_EN_CI = ('compras.parana.gob.ar',)

    def _con_navegador(self, url, d, json_, motivo):
        import navegador
        if EN_LA_NUBE and any(h in url for h in self.SIN_SALIDA_EN_CI):
            d['error'] = ('Cloudflare no deja entrar desde la nube. Esta fuente se '
                          'actualiza cuando corre en la Mac de Delfi.')
            return None
        if not navegador.disponible():
            return None
        sel = next((v for k, v in self.ESPERAR.items() if k in url), None)
        html = navegador.bajar(url, esperar=sel, diagnostico=True)
        if html is None:
            return None
        d['http'] = f'{motivo} → navegador'
        d['error'] = ''
        d['bytes'] = len(html)
        if json_:
            m = re.search(r'<pre[^>]*>(.*?)</pre>', html, re.S)
            crudo = m.group(1) if m else html
            try:
                import json as _j, html as _h
                return _j.loads(_h.unescape(re.sub(r'<[^>]+>', '', crudo)))
            except Exception:
                return None
        return html

# ---------------------------------------------------------------- parsers
ORG = {'ENERGIA':'Sec. de Energía','DGAyC':'Arquitectura y Construcciones (DGAyC)',
       'DH':'Dir. de Hidráulica','DPV':'Vialidad (DPV)','IAPV':'IAPV (Vivienda)',
       'OSER':'Obras Sanitarias (OSER)','UEP':'Unidad Ejecutora Provincial',
       'CAFESG':'CAFESG','SMPIyS':'Sec. de Planeamiento e Infraestructura'}

def p_parana(red, url, nombre):
    html = red.get(url, nombre)
    if not html: return []
    soup = BeautifulSoup(html, 'lxml')
    origin = re.match(r'^https?://[^/]+', url).group(0)
    out = []
    for tr in soup.select('tr'):
        tds = tr.find_all('td')
        a = tr.find('a', href=re.compile(r'/uploads/pliegos/'))
        if not a or len(tds) < 4: continue
        # OJO: la cantidad de columnas cambia según cómo se lea la página.
        # Bajada directa: 5 tds (la 1ª es una columna oculta de ordenamiento).
        # Con navegador: 4 tds, porque DataTables elimina esa columna al renderizar.
        # Por eso ubicamos la celda del PDF y contamos a partir de ella.
        i_pdf = next((i for i, td in enumerate(tds) if td.find('a', href=re.compile(r'/uploads/pliegos/'))), None)
        if i_pdf is None or i_pdf + 3 >= len(tds) + 0:
            if i_pdf is None or len(tds) < i_pdf + 4: continue
        apertura = norm(tds[i_pdf + 1].get_text())
        tramite  = norm(tds[i_pdf + 2].get_text())
        objeto   = norm(tds[i_pdf + 3].get_text())
        partes = tramite.split(' - ')
        pdf = origin + a['href']
        out.append(dict(fuente=nombre, organismo='Municipalidad de Paraná',
            tipo=partes[0].strip() if partes else '', numero=partes[1].strip() if len(partes) > 1 else '',
            objeto=objeto, apertura_txt=apertura, fecha=parse_fecha(apertura),
            link=url, pdf=pdf, uid=pdf))
    return out

def p_minplan(red, url, nombre):
    html = red.get(url, nombre)
    if not html: return []
    soup = BeautifulSoup(html, 'lxml')
    base = re.match(r'^(https?://.*/)[^/]*$', url).group(1)
    out, visto = [], set()
    for a in soup.find_all('a', href=re.compile(r'^licitaciones/')):
        href = a['href']
        if href in visto: continue
        txt = norm(a.get_text(' '))
        m = re.search(r'Apertura:\s*(.+?)(?:\s*$)', txt)
        apertura = norm(m.group(1)) if m else ''
        strong = a.find('strong')
        if not strong: continue
        num_full = norm(strong.get_text())
        toks = num_full.split()
        code = toks[-1] if len(toks) > 1 else ''
        numero = ' '.join(toks[:-1]) if len(toks) > 1 else num_full
        # objeto: el párrafo con más texto que no sea el número ni la apertura
        cands = [norm(p.get_text()) for p in a.find_all(['p', 'h4'])]
        cands = [c for c in cands if c and 'Apertura' not in c and c != num_full and len(c) > 12]
        objeto = max(cands, key=len) if cands else txt
        objeto = re.sub(r'^' + re.escape(num_full) + r'\s*', '', objeto).strip()
        tipo = 'Licitación Pública'
        for t in ('Licitación Privada', 'Concurso de Precios', 'Licitación Pública', 'Concurso'):
            if t.lower() in txt.lower(): tipo = t; break
        visto.add(href)
        link = base + href
        out.append(dict(fuente=nombre, organismo=ORG.get(code, f'Min. Planeamiento {code}'.strip()),
            tipo=tipo, numero=numero, objeto=objeto, apertura_txt=apertura,
            fecha=parse_fecha(apertura), link=link, pdf='', uid=link))
    return out

def p_iapv(red, url, nombre):
    html = red.get(url, nombre)
    if not html: return []
    soup = BeautifulSoup(html, 'lxml')
    origin = re.match(r'^https?://[^/]+', url).group(0)
    out = []
    for h3 in soup.find_all('h3'):
        a = h3.find('a', href=True)
        if not a: continue
        titulo = norm(a.get_text())
        if not re.search(r'licitaci|concurso|contrataci', titulo, re.I): continue
        cont = h3.find_next('p')
        detalle = norm(cont.get_text()) if cont else ''
        bloque = h3.find_parent(class_=re.compile('divList')) or h3.parent
        fspan = bloque.find(class_=re.compile('divList_Fecha')) if bloque else None
        publicada = norm(fspan.get_text()) if fspan else ''
        m = re.search(r'n[°ºo]?\s*(\d+\s*/\s*\d{2,4})', titulo, re.I)
        obj = re.search(r'OBRA\s*:\s*([^.]+)', detalle, re.I)
        link = a['href'] if a['href'].startswith('http') else origin + a['href']
        out.append(dict(fuente=nombre, organismo='IAPV (Vivienda)', tipo='Licitación Pública',
            numero=re.sub(r'\s+', '', m.group(1)) if m else '',
            objeto=obj.group(1).strip() if obj else titulo, rubro='Infraestructura',
            apertura_txt=f'Publicada {publicada} — ver pliego' if publicada else 'Ver pliego',
            fecha=None, link=link, pdf='', uid=link,
            venta=venta_pliego(detalle)))
    return out

def p_enersa(red, url, nombre):
    html = red.get(url, nombre)
    if not html: return []
    texto = BeautifulSoup(html, 'lxml').get_text('\n')
    out, visto = [], set()
    for m in re.finditer(r'Licitaci[oó]n\s*N?°?\s*(\d{4}-\d{3,4})\s*[:\-–]\s*["“]?([^"”\n]+)', texto):
        num, obj = m.group(1), norm(m.group(2)).rstrip('”"')
        if int(num.split('-')[0]) < HOY.year or num in visto: continue
        visto.add(num)
        out.append(dict(fuente=nombre, organismo='ENERSA', tipo='Licitación', numero=num,
            objeto=obj, apertura_txt='Ver pliego', fecha=None,
            link=url, pdf='', uid=f'{nombre}|{num}'))
    return out

def p_wpjson(red, url, nombre):
    origin = re.match(r'^https?://[^/]+', url).group(0)
    data = red.get(f'{origin}/wp-json/wp/v2/posts?per_page=40&_fields=date,link,title,excerpt',
                   nombre, json_=True)
    if not isinstance(data, list): return []
    out = []
    for p in data:
        titulo = norm(BeautifulSoup(p.get('title', {}).get('rendered', ''), 'lxml').get_text())
        if not re.match(r'\s*(licitaci|concurso|contrataci|adquisici|provisi|compulsa)', titulo, re.I):
            continue                                            # descarta noticias
        resumen = norm(BeautifulSoup(p.get('excerpt', {}).get('rendered', ''), 'lxml').get_text(' '))
        corte = r'(?=\s*[–—]?\s*(?:APERTURA|CONSULTA|PRESUPUESTO|VALOR DEL PLIEGO|VENTA|LUGAR|PLAZO|GARANT)\s*:|$)'
        mo = re.search(r'(?:OBJETO|PROYECTO A CONTRATAR|OBRA A CONTRATAR|OBRA)\s*:\s*(.{5,400}?)' + corte,
                       resumen, re.I | re.S)
        objeto = norm(mo.group(1)).strip(' .–—-') if mo else titulo
        mf = re.search(r'APERTURA\s*:?\s*(?:El\s+)?(\d{1,2}\s*[/\-.]\s*\d{1,2}\s*[/\-.]\s*\d{2,4})',
                       resumen, re.I)
        apertura = re.sub(r'\s+', '', mf.group(1)) if mf else 'Ver pliego'
        mn = re.search(r'n[°ºo]?\s*(\d+\s*/\s*\d{2,4})', titulo, re.I) or re.search(r'\b(\d{3,5})\b\s*:', titulo)
        tipo = ('Concurso' if re.search(r'concurso', titulo, re.I) else
                'Licitación Privada' if re.search(r'privada', titulo, re.I) else
                'Contratación Directa' if re.search(r'contrataci', titulo, re.I) else
                'Licitación Pública' if re.search(r'licitaci', titulo, re.I) else 'Adquisición/Provisión')
        out.append(dict(fuente=nombre, organismo=nombre, tipo=tipo,
            numero=re.sub(r'\s+', '', mn.group(1)) if mn else '', objeto=objeto,
            apertura_txt=apertura, fecha=parse_fecha(apertura) if mf else None,
            link=p.get('link', url), pdf='', uid=p.get('link', ''),
            venta=venta_pliego(resumen)))
    return out

def p_generico(red, url, nombre):
    html = red.get(url, nombre)
    if not html: return []
    soup = BeautifulSoup(html, 'lxml')
    for t in soup(['script', 'style', 'nav', 'footer']): t.decompose()
    disp = re.compile(r'licitaci[oó]n|concurso de precios|concurso|contrataci[oó]n|obra p[uú]blica|adquisici[oó]n', re.I)
    yy = HOY.year % 100
    # "2026", "2025", "10/26", "09/2025", "10/2.025"
    anios = re.compile(r'\b(?:%d|%d)\b|\b\d{1,3}\s*/\s*(?:%d|%d|%d|%d)\b|\b\d{1,3}\s*/\s*2\.?0(?:%d|%d)\b'
                       % (HOY.year, HOY.year-1, yy, yy-1, HOY.year, HOY.year-1, yy, yy-1))
    out, visto = [], set()
    for linea in soup.get_text('\n').split('\n'):
        l = norm(linea)
        if not (15 < len(l) < 200) or not disp.search(l) or not anios.search(l): continue
        k = sin_acentos(l.lower())
        if k in visto: continue
        visto.add(k)
        mn = re.search(r'n[°ºo]?\s*(\d+\s*/\s*\d{2,4})', l, re.I)
        out.append(dict(fuente=nombre, organismo=nombre, tipo='(según sitio)',
            numero=re.sub(r'\s+', '', mn.group(1)) if mn else '', objeto=l,
            apertura_txt='Ver el sitio', fecha=parse_fecha(l),
            link=url, pdf='', uid=f'{nombre}|{k[:120]}'))
    return out


def p_boletin(red, url, nombre):
    """Boletín Oficial de Entre Ríos. La fuente más completa: por ley publican
    acá TODOS los organismos y municipios, incluidas las comunas chicas sin web.
    Es también la única que trae VENTA DE PLIEGOS y VALOR DEL PLIEGO."""
    import boletin as BO
    d = red.log.setdefault(nombre, {'http': '', 'bytes': 0, 'error': '', 'crudos': 0, 'relev': 0})
    try:
        avisos, bajados = BO.recolectar(dias=BOLETIN_DIAS, sesion=red.s)
    except Exception as e:
        d['error'] = f'{type(e).__name__}: {e}'[:120]
        return []
    if not bajados:
        d['http'] = 'sin boletines'
        d['error'] = f'No se pudo bajar ningún boletín de los últimos {BOLETIN_DIAS} días'
        return []
    d['http'] = f'{len(bajados)} PDFs'
    out, vistos = [], set()
    for a in avisos:
        # el mismo aviso se publica varios días seguidos ("3 v./04/09/2026")
        clave = a['aviso_id'] or f"{a['organismo']}|{a['numero']}|{a['objeto'][:80]}"
        if clave in vistos: continue
        vistos.add(clave)
        apert_txt = a['apertura'] or ''
        f_ap = parse_fecha(apert_txt)
        venta = a['venta'] or ''
        f_venta = venta_pliego(venta) or venta_pliego(apert_txt) or ''
        detalle = []
        if a['valor_pliego']: detalle.append('Pliego: ' + a['valor_pliego'][:60])
        if a['presupuesto']:  detalle.append('Presup.: ' + a['presupuesto'][:50])
        out.append(dict(
            fuente=nombre,
            organismo=a['organismo'][:90] or 'Boletín Oficial',
            tipo=a['tipo'], numero=a['numero'], objeto=a['objeto'],
            apertura_txt=norm(apert_txt)[:120] or 'Ver boletín',
            fecha=f_ap,
            venta=f_venta,
            valor_pliego=a['valor_pliego'][:70] if a['valor_pliego'] else '',
            link=a['url'], pdf=a['url'],
            uid=f"BO|{a['aviso_id']}" if a['aviso_id'] else f"BO|{a['organismo']}|{a['numero']}|{a['objeto'][:60]}",
            extra=' · '.join(detalle),
        ))
    return out

PARSERS = {'parana': p_parana, 'boletin': p_boletin, 'minplan': p_minplan, 'iapv': p_iapv,
           'enersa': p_enersa, 'wpjson': p_wpjson, 'generico': p_generico}
# ---------------------------------------------------------------- salida
COLS = ['Detectada','Estado','Rubro','Organismo','Tipo','N°','Objeto',
        'Apertura','Detalle apertura','Venta hasta','Valor pliego',
        'Link','Pliego PDF','Fuente','ID']
COLORES = {'Arquitectura':'D9E1F2','Infraestructura':'E2EFDA','Áridos':'FCE4D6','Vehículos':'FFF2CC'}

def escribir_excel(filas, ruta):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    wb = Workbook(); ws = wb.active; ws.title = 'Licitaciones'
    ws.append(COLS)
    for c in range(1, len(COLS) + 1):
        cel = ws.cell(1, c)
        cel.font = Font(bold=True, color='FFFFFF')
        cel.fill = PatternFill('solid', fgColor='1F3A5F')
    orden = {'Vigente': 0, 'Revisar fecha': 1, 'Vencida': 2}
    filas = sorted(filas, key=lambda f: (orden.get(f['Estado'], 3), f['Apertura'] or ''))
    for f in filas:
        ws.append([('' if f.get(c) is None else f.get(c, '')) for c in COLS])
        r = ws.max_row
        # el N° va como TEXTO a propósito: si no, Excel convierte "08/2026" en fecha
        for nombre in ('N°','Apertura','Detalle apertura','Venta hasta','Valor pliego','ID'):
            ws.cell(r, COLS.index(nombre) + 1).number_format = '@'
        col_rubro = COLS.index('Rubro') + 1
        if f['Rubro'] in COLORES:
            ws.cell(r, col_rubro).fill = PatternFill('solid', fgColor=COLORES[f['Rubro']])
            ws.cell(r, col_rubro).font = Font(bold=True)
        if f['Estado'] == 'Vigente':
            ws.cell(r, COLS.index('Estado') + 1).font = Font(bold=True, color='1B7F3B')
        for nombre, etiqueta in (('Link','ver licitación'), ('Pliego PDF','PDF')):
            u = f.get(nombre, '')
            if u and u.startswith('http'):
                cel = ws.cell(r, COLS.index(nombre) + 1)
                cel.value = etiqueta; cel.hyperlink = u
                cel.font = Font(color='0563C1', underline='single')
        ws.cell(r, COLS.index('Objeto') + 1).alignment = Alignment(wrap_text=True, vertical='center')
    anchos = {'Detectada':11,'Estado':14,'Rubro':16,'Organismo':32,'Tipo':20,'N°':12,
              'Objeto':70,'Apertura':12,'Detalle apertura':40,'Venta hasta':14,
              'Valor pliego':28,'Link':14,'Pliego PDF':11,'Fuente':26,'ID':1}
    for nombre, w in anchos.items():
        ws.column_dimensions[get_column_letter(COLS.index(nombre) + 1)].width = w
    ws.column_dimensions[get_column_letter(COLS.index('ID') + 1)].hidden = True
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(COLS))}{ws.max_row}'
    wb.save(ruta)
    return len(filas)

DIAG_COLS = ['Fuente','URL','Parser','Estado','Respuesta','Encontradas','De tu rubro',
             'Corridas sin traer nada','Detalle','Última corrida']

def escribir_diagnostico(red, fuentes, estado_previo):
    """Deja en diagnostico.csv cómo le fue a cada fuente. La planilla lo muestra
    para que se vea de un vistazo si alguna dejó de andar y haya que entrar a
    mano — así no se pierde ninguna licitación."""
    import csv
    ceros = estado_previo.get('ceros', {})
    ahora = datetime.now().strftime('%d/%m/%Y %H:%M')
    filas, alertas = [], []
    for f in fuentes:
        n = f['nombre']
        d = red.log.get(n, {})
        crudos = d.get('crudos', 0)
        error = d.get('error', '')
        sin_nada = (ceros.get(n, 0) + 1) if crudos == 0 else 0
        ceros[n] = sin_nada
        if error:
            estado = 'ERROR'
        elif crudos == 0:
            estado = 'Sin resultados'
        else:
            estado = 'OK'
        motivo = error or ('No trajo nada hace %d corrida(s)' % sin_nada if sin_nada >= 3 else '')
        filas.append({
            'Fuente': n, 'URL': f['url'], 'Parser': f['parser'], 'Estado': estado,
            'Respuesta': str(d.get('http', '')), 'Encontradas': crudos,
            'De tu rubro': d.get('relev', 0), 'Corridas sin traer nada': sin_nada,
            'Detalle': motivo, 'Última corrida': ahora,
        })
        if estado == 'ERROR' or sin_nada >= 3:
            alertas.append(f"{n}: {motivo or 'no trajo resultados'} — entrá a {f['url']}")
    # las apagadas también se listan, para que se vean en la planilla
    for f in TODAS_LAS_FUENTES:
        if f.get('activa') or any(x['Fuente'] == f['nombre'] for x in filas):
            continue
        filas.append({'Fuente': f['nombre'], 'URL': f['url'], 'Parser': f.get('parser', ''),
                      'Estado': 'Apagada', 'Respuesta': '', 'Encontradas': '', 'De tu rubro': '',
                      'Corridas sin traer nada': '', 'Detalle': f.get('nota', ''), 'Última corrida': ahora})
    with open(BASE / 'diagnostico.csv', 'w', newline='', encoding='utf-8') as fh:
        w = csv.DictWriter(fh, fieldnames=DIAG_COLS)
        w.writeheader()
        w.writerows(filas)
    estado_previo['ceros'] = ceros
    return alertas


def leer_csv_previo(ruta=None):
    """Lee el CSV de la corrida anterior. Sirve para no perder las licitaciones
    de una fuente que hoy falló (típicamente Paraná cuando corre en la nube)."""
    import csv
    ruta = ruta or (BASE / 'licitaciones.csv')
    if not ruta.exists():
        return []
    try:
        with open(ruta, newline='', encoding='utf-8') as f:
            lector = csv.DictReader(f)
            previas = []
            for fila in lector:
                d = {c: (fila.get(c) or '') for c in COLS if c in fila}
                # el CSV no guarda el ID (es interno): lo recalculamos igual que siempre
                if not d.get('ID'):
                    base = d.get('Link') or f"{d.get('Fuente','')}|{d.get('N°','')}|{d.get('Objeto','')}"
                    d['ID'] = uid_hash(base)
                previas.append(d)
            return previas
    except Exception as e:
        print(f'  (no pude leer el CSV anterior: {e})')
        return []

# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--solo', help='correr una sola fuente (parte del nombre)')
    ap.add_argument('--todo', action='store_true', help='mostrar también las vencidas')
    ap.add_argument('--sin-mail', action='store_true', help='no mandar mail (para probar)')
    args = ap.parse_args()

    cfg = json.loads((BASE / 'fuentes.json').read_text(encoding='utf-8'))
    todas = list(cfg['fuentes'])
    # lo que el equipo cargó en la planilla manda sobre el archivo
    de_planilla = fuentes_de_la_planilla()
    if de_planilla:
        por_url = {f['url'].rstrip('/'): i for i, f in enumerate(todas)}
        sumadas = 0
        for f in de_planilla:
            k = f['url'].rstrip('/')
            if k in por_url:
                todas[por_url[k]].update(f)
            else:
                todas.append(f); sumadas += 1
        print(f'  ({len(de_planilla)} fuentes leídas de la planilla'
              + (f', {sumadas} nueva(s)' if sumadas else '') + ')')
    global TODAS_LAS_FUENTES
    TODAS_LAS_FUENTES = todas
    fuentes = [f for f in todas if f.get('activa')]
    if args.solo:
        fuentes = [f for f in fuentes if args.solo.lower() in f['nombre'].lower()]

    est_path = BASE / 'estado.json'
    estado = json.loads(est_path.read_text(encoding='utf-8')) if est_path.exists() else {'vistos': {}}
    vistos = estado.get('vistos', {})

    red = Red()
    items = []
    print(f'\n{"="*78}\nLICITACIONES ENTRE RÍOS · {HOY.strftime("%d/%m/%Y")} · {len(fuentes)} fuentes\n{"="*78}')

    # En paralelo y con LÍMITE DURO por fuente. El timeout de requests no alcanza:
    # Oro Verde se colgó 930s y DPV 1119s pese a tener timeout puesto, porque el
    # problema estaba a nivel de conexión/redirects. Así ninguna fuente puede
    # frenar la corrida entera.
    from concurrent.futures import ThreadPoolExecutor, TimeoutError as FTimeout
    LIMITE = 45
    LIMITE_LARGO = 150      # para fuentes pesadas como el Boletín (baja ~20 PDFs)

    def trabajo(f):
        fn = PARSERS.get(f['parser'], p_generico)
        try:
            return fn(red, f['url'], f['nombre'])
        finally:
            # si este hilo abrió un navegador, lo cierra él mismo
            try:
                import navegador
                navegador.cerrar()
            except Exception:
                pass

    with ThreadPoolExecutor(max_workers=6) as pool:
        futuros = {f['nombre']: (pool.submit(trabajo, f), time.time(), f) for f in fuentes}
        for nombre, (fut, t0, f) in futuros.items():
            try:
                tope = LIMITE_LARGO if f['parser'] == 'boletin' else LIMITE
                got = fut.result(timeout=tope)
            except FTimeout:
                got = []
                red.log.setdefault(nombre, {})['error'] = f'No respondió a tiempo (se abandonó)'
                red.log[nombre].setdefault('http', 'TIMEOUT')
                fut.cancel()
            except Exception as e:
                got = []
                red.log.setdefault(nombre, {})['error'] = f'{type(e).__name__}: {e}'[:120]
            d = red.log.setdefault(nombre, {})
            d['crudos'] = len(got)
            items += got
            print(f'  {nombre[:38]:<40} {str(d.get("http","")):>7}  {len(got):>4} items  '
                  f'{time.time()-t0:5.1f}s  {d.get("error","")[:32]}')

    # filtrar por rubro
    relev = []
    for it in items:
        texto = f'{it.get("objeto","")} {it.get("tipo","")}'
        if es_ruido(texto): continue
        rubro = it.get('rubro') or clasificar(texto)
        if not rubro: continue
        it['rubro'] = rubro
        red.log.setdefault(it['fuente'], {}).setdefault('relev', 0)
        red.log[it['fuente']]['relev'] = red.log[it['fuente']].get('relev', 0) + 1
        relev.append(it)

    # armar filas + dedup por ID
    filas, nuevas = [], []
    for it in relev:
        _id = uid_hash(it.get('uid') or f'{it["fuente"]}|{it.get("numero","")}|{it.get("objeto","")}')
        est = estado_de(it.get('fecha'))
        if est == 'Vencida' and not args.todo: continue
        fila = {
            'Detectada': vistos.get(_id, HOY.strftime('%d/%m/%Y')),
            'Estado': est, 'Rubro': it['rubro'], 'Organismo': it.get('organismo', ''),
            'Tipo': it.get('tipo', ''), 'N°': it.get('numero', ''), 'Objeto': it.get('objeto', ''),
            # fecha limpia para ordenar y filtrar; el texto largo va aparte
            'Apertura': it['fecha'].strftime('%d/%m/%Y') if it.get('fecha') else '',
            'Detalle apertura': limpiar_apertura(it.get('apertura_txt', '')),
            'Venta hasta': it.get('venta') or 'Ver pliego',
            'Valor pliego': it.get('valor_pliego') or '',
            'Link': it.get('link', ''), 'Pliego PDF': it.get('pdf', ''),
            'Fuente': it['fuente'], 'ID': _id,
        }
        # dedup dentro de la misma corrida
        if any(x['ID'] == _id for x in filas): continue
        filas.append(fila)
        if _id not in vistos:
            nuevas.append(fila)
            vistos[_id] = HOY.strftime('%d/%m/%Y')

    # Si una fuente falló, NO tiramos lo que había traído la última vez.
    # Es lo que hace posible el esquema mixto: cuando corre en la nube Paraná
    # falla (Cloudflare), pero sus licitaciones —que subió la Mac— se mantienen.
    fallaron = {n for n, d in red.log.items() if d.get('error')}
    if fallaron:
        previas = leer_csv_previo()
        rescatadas = [f for f in previas
                      if f.get('Fuente') in fallaron
                      and f.get('ID') not in {x['ID'] for x in filas}
                      and f.get('Estado') != 'Vencida']
        if rescatadas:
            filas += rescatadas
            print(f'\n  Se conservaron {len(rescatadas)} licitación(es) de '
                  f'{len(fallaron)} fuente(s) que hoy fallaron: {", ".join(sorted(fallaron))[:70]}')

    xlsx = BASE / 'Licitaciones-Entre-Rios.xlsx'
    escribir_excel(filas, xlsx)
    # CSV para que la Google Sheet lo lea con =IMPORTDATA (sin credenciales)
    try:
        import notificar
        notificar.escribir_csv(filas, COLS)
    except Exception as e:
        print(f'  (no se pudo escribir el CSV: {e})')
    estado['vistos'] = vistos
    estado['ultima_corrida'] = datetime.now().isoformat(timespec='seconds')

    vig = [f for f in filas if f['Estado'] == 'Vigente']
    rev = [f for f in filas if f['Estado'] == 'Revisar fecha']
    print(f'\n{"-"*78}')
    print(f'  {len(items)} items bajados  ->  {len(relev)} de tus rubros  ->  {len(filas)} en el Excel')
    print(f'  {len(vig)} VIGENTES · {len(rev)} a confirmar fecha · {len(nuevas)} NUEVAS desde la última corrida')
    print(f'{"-"*78}')
    if vig:
        print('\n  VIGENTES (con fecha de apertura confirmada a futuro):')
        for f in vig:
            nuevo = ' *NUEVA*' if f in nuevas else ''
            print(f'   · [{f["Rubro"]:<15}] {f["Organismo"][:30]:<32} {f["N°"]:<12} {f["Objeto"][:52]}')
            print(f'       apertura: {f["Apertura"][:46]:<48} venta: {f["Venta hasta"]}{nuevo}')
    # estado de cada fuente -> diagnostico.csv + alertas para el mail
    alertas = escribir_diagnostico(red, fuentes, estado)
    # recién ahora se guarda: escribir_diagnostico actualiza el contador de
    # corridas sin resultados, que es lo que dispara el aviso a los 3 días
    est_path.write_text(json.dumps(estado, ensure_ascii=False, indent=1), encoding='utf-8')
    if alertas:
        print('\n  FUENTES CON PROBLEMAS:')
        for a in alertas: print(f'   ! {a[:100]}')

    try:
        import navegador; navegador.cerrar()
    except Exception:
        pass

    print(f'\n  Excel: {xlsx}')
    try:
        import notificar
        notificar.avisar_mac(nuevas, len(vig))          # no necesita contraseñas
        if not args.sin_mail:
            notificar.enviar(nuevas, COLS, len(vig), alertas, adjunto=xlsx)
    except Exception as e:
        print(f'  (aviso: {type(e).__name__}: {e})')
    print()
    return filas, nuevas

if __name__ == '__main__':
    main()
