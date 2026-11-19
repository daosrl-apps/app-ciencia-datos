#!/usr/bin/env python3
"""
Pipeline DAO SRL — de los Excel mensuales de Luciano a los CSV del tablero.

Qué hace
  1. Recorre una carpeta con los archivos que Luciano guarda cada mes:
       ER_MM-AAAA[_SUPERFICIE].xls   (estado de resultados teórico, superficie y receta por artículo)
       WEG MM_AAAA.xls               (compras de pintura a Emprendimientos / WEG)
       COMPRAS MM_AAAA.xls           (facturas y pagos del mes)
       Resumen*AAAA*.xls             (Resumen anual o interanual: sueldos, cargas, MOD real, retiros, kg)
  2. Identifica cada archivo por su nombre (tolerante a espacios, guiones, mayúsculas)
     y valida que el contenido corresponda al mes del nombre.
  3. Extrae los datos ubicándolos por rótulo, nunca por celda fija.
  4. Guarda cada extracción en la base histórica (base/) con UPSERT por mes:
     un mes nuevo se agrega, un mes existente se reemplaza solo si el archivo es más nuevo
     (fecha de emisión que trae el propio reporte). Nunca se borran meses.
  5. Arma data/detalle_clientes_articulos.csv, data/resumen_mensual.csv y data/calidad.json.

Uso
  python dao_pipeline.py --entrada <carpeta_excel> --base <carpeta_base> --salida <carpeta_data>

Requisitos: Python 3.9+, openpyxl (para .xlsx) y xlrd (para .xls).
"""
from __future__ import annotations
import argparse, csv, datetime as dt, hashlib, json, os, re, shutil, subprocess, sys, tempfile
from collections import defaultdict

VERSION = '1.3'
MESES = {'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6, 'jul': 7, 'ago': 8, 'sep': 9, 'set': 9, 'oct': 10, 'nov': 11, 'dic': 12}

# ════════════════════════════ Errores para el usuario ════════════════════════════
class ErrorCarga(Exception):
    """Error que se le muestra a Luciano en el tablero: qué está mal y qué hacer.
    El archivo se antepone al mostrarlo; el detalle técnico queda en registro_cargas.csv."""
    def __init__(self, problema, que_hacer='Revisá que sea el archivo que exporta el sistema, sin editar, y volvé a guardarlo en la carpeta.'):
        super().__init__(f'{problema}. {que_hacer}')

def _fila(rows, cond, que, hoja, tipo):
    """Índice de la primera fila que cumple cond; si no existe, un error claro en vez de StopIteration."""
    i = next((i for i, r in enumerate(rows) if r and cond(r)), None)
    if i is None:
        raise ErrorCarga(f'no encuentro {que} en la hoja {hoja} del archivo {tipo}')
    return i

def _col(hdr, nombre, hoja, tipo):
    """Posición de una columna por su encabezado; si no está, un error claro en vez de "'X' is not in list"."""
    if nombre not in hdr:
        raise ErrorCarga(f'falta la columna "{nombre}" en la hoja {hoja} del archivo {tipo}')
    return hdr.index(nombre)

# ════════════════════════════ Lectura de planillas ════════════════════════════
_TEXTOS_COMO_CERO = []   # textos encontrados donde se esperaba un número (se reinicia por archivo)

def _num(v):
    if v is None or v == '':
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace('$', '').replace(' ', '')
    if s in ('', '-', '—'):
        return 0.0
    if re.match(r'^-?[\d.]+,\d+$', s):            # formato 1.234,56
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        _TEXTOS_COMO_CERO.append(str(v).strip()[:40])   # se toma 0, pero queda aviso
        return 0.0

def leer_libro(path):
    """Devuelve {hoja: [[valores...], ...]} con fechas como datetime. Soporta .xls (xlrd) y .xlsx (openpyxl)."""
    ext = os.path.splitext(path)[1].lower()
    if ext == '.xls':
        try:
            import xlrd
            wb = xlrd.open_workbook(path)
            out = {}
            for sh in wb.sheets():
                rows = []
                for r in range(sh.nrows):
                    row = []
                    for c in range(sh.ncols):
                        cell = sh.cell(r, c)
                        if cell.ctype == xlrd.XL_CELL_DATE:
                            row.append(xlrd.xldate.xldate_as_datetime(cell.value, wb.datemode))
                        elif cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                            row.append(None)
                        else:
                            row.append(cell.value)
                    rows.append(row)
                out[sh.name] = rows
            return out
        except ImportError:
            # Sin xlrd: convierte con LibreOffice a .xlsx
            tmp = tempfile.mkdtemp()
            subprocess.run(['soffice', '--headless', '--convert-to', 'xlsx', '--outdir', tmp, path],
                           check=True, capture_output=True, timeout=300)
            path = os.path.join(tmp, os.path.splitext(os.path.basename(path))[0] + '.xlsx')
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    return {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}

def txt(v):
    return '' if v is None else str(v).strip()

def fecha_emision(rows):
    """Los reportes del sistema traen en A1 'Fecha: dd/mm/aa Hora: hh:mm:ss'."""
    for r in rows[:3]:
        for v in r[:3]:
            m = re.search(r'Fecha:\s*(\d{2})/(\d{2})/(\d{2,4})\s*Hora:\s*(\d{2}):(\d{2}):(\d{2})', txt(v))
            if m:
                d, mo, y, h, mi, s = m.groups()
                y = int(y) + (2000 if len(y) == 2 else 0)
                return dt.datetime(y, int(mo), int(d), int(h), int(mi), int(s)).isoformat()
    return None

# ════════════════════════════ Identificación de archivos ════════════════════════════
PAT_MENSUAL = re.compile(r'^(ER|WEG|COMPRAS)[\s_\-]*(\d{1,2})[\s_\-]+(\d{4})', re.I)
PAT_RESUMEN = re.compile(r'resumen', re.I)

def identificar(nombre):
    base = os.path.basename(nombre)
    if base.startswith('~$') or not re.search(r'\.xlsx?$', base, re.I):
        return None
    m = PAT_MENSUAL.match(base)
    if m:
        tipo, mes, anio = m.group(1).upper(), int(m.group(2)), int(m.group(3))
        if 1 <= mes <= 12:
            return {'tipo': tipo, 'periodo': f'{anio}-{mes:02d}'}
    if PAT_RESUMEN.search(base):
        return {'tipo': 'RESUMEN', 'periodo': None}
    return None

# ════════════════════════════ Extractor ER ════════════════════════════
def extraer_er(libro, periodo):
    hojas = {k.upper(): v for k, v in libro.items()}
    for req in ('ESTADO_DE_RESULTADO', 'DETALLE_DE_CALCULOS', 'SUP_PINTADA'):
        if req not in hojas:
            raise ErrorCarga(f'falta la hoja {req}. Puede ser otro tipo de archivo guardado con nombre de ER, o una hoja renombrada')
    er = hojas['ESTADO_DE_RESULTADO']
    emision = fecha_emision(er)
    # Período según el propio reporte (fila DESDE)
    desde = next((r[1] for r in er[:6] if r and len(r) > 1 and txt(r[0]).upper() == 'DESDE'), None)
    if isinstance(desde, dt.datetime) and f'{desde.year}-{desde.month:02d}' != periodo:
        raise ErrorCarga(f'el ER dice DESDE {desde:%m/%Y} pero el nombre indica {periodo[5:]}/{periodo[:4]}',
                         'Corregí el nombre del archivo para que coincida con el mes del reporte.')

    # Ingresos por cliente: entre el encabezado CLIENTE y 'TOTAL CV'
    clientes, nombres = {}, {}
    i_hdr = _fila(er, lambda r: txt(r[0]).upper() == 'CLIENTE', 'la tabla de clientes (fila "CLIENTE")', 'ESTADO_DE_RESULTADO', 'ER')
    hdr = [txt(x).upper() for x in er[i_hdr]]
    c_total = _col(hdr, 'TOTAL', 'ESTADO_DE_RESULTADO', 'ER'); c_malbec = _col(hdr, 'MALBEC', 'ESTADO_DE_RESULTADO', 'ER')
    c_neto = _col(hdr, 'NETO', 'ESTADO_DE_RESULTADO', 'ER')
    totales = {}
    for r in er[i_hdr + 1:]:
        a = txt(r[0])
        if a.upper().startswith('TOTAL CV'):
            break
        if a and re.match(r'^\d+$', a):
            clientes[a] = {'total': _num(r[c_total]), 'malbec': _num(r[c_malbec]), 'neto': _num(r[c_neto])}
            nombres[a] = txt(r[1])
        elif not a and r[c_total] not in (None, ''):
            totales['ingresos'] = _num(r[c_total])
    # Bloque de costos y resultado (por rótulo)
    def valor_rotulo(rot, col=None):
        for r in er:
            if txt(r[0]).lower().startswith(rot.lower()):
                vals = [x for x in r[1:] if isinstance(x, (int, float))]
                return vals[-1] if vals else 0.0
        return None
    totales['cv_pintura'] = valor_rotulo('Pintura')
    totales['cv_quimicos'] = valor_rotulo('Químicos') or valor_rotulo('Quimicos')
    totales['cv_mod'] = valor_rotulo('MOD')
    totales['cmg'] = valor_rotulo('Contribucion')
    totales['cf'] = valor_rotulo('Costos fijos')
    totales['rb'] = valor_rotulo('Resultado bruto')

    # Superficie: buscar la fila de encabezado (a veces hay una fila de totales arriba)
    sup = hojas['SUP_PINTADA']
    i_s = _fila(sup, lambda r: txt(r[0]).upper() == 'ARTICULO', 'el encabezado de artículos (fila "ARTICULO")', 'SUP_PINTADA', 'ER')
    hs = [txt(x).upper() for x in sup[i_s]]
    for k in ('ARTICULO', 'DESCRIPCIÓN', 'TOTAL SUP.', 'CLIENTE', 'NOMBRE'):
        _col(hs, k, 'SUP_PINTADA', 'ER')
    ix = {k: hs.index(k) for k in ('ARTICULO', 'DESCRIPCIÓN', 'TOTAL SUP.', 'CLIENTE', 'NOMBRE', 'CANT. MALBEC', 'CANT. DAO') if k in hs}
    superficie = []
    for r in sup[i_s + 1:]:
        art = txt(r[ix['ARTICULO']])
        if not re.match(r'^\d+$', art):
            continue
        cli = txt(r[ix['CLIENTE']])
        superficie.append({'articulo': art, 'descripcion': txt(r[ix['DESCRIPCIÓN']]), 'cod_cliente': cli,
                           'm2': _num(r[ix['TOTAL SUP.']]),
                           'piezas': _num(r[ix['CANT. MALBEC']]) + _num(r[ix['CANT. DAO']]) if 'CANT. DAO' in ix else None})
        nombres.setdefault(cli, txt(r[ix['NOMBRE']]))

    # Receta por artículo (DETALLE_DE_CALCULOS)
    det = hojas['DETALLE_DE_CALCULOS']
    if not det:
        raise ErrorCarga('la hoja DETALLE_DE_CALCULOS del ER está vacía')
    hd = [txt(x) for x in det[0]]
    def cols(prefix):
        return [i for i, h in enumerate(hd) if h.lower().startswith(prefix.lower())]
    c_pin, c_qui, c_mod = cols('Pintura monto'), cols('Químicos monto') or cols('Quimicos monto'), cols('MOD monto')
    c_cm, c_cd = _col(hd, 'CANT. MALBEC', 'DETALLE_DE_CALCULOS', 'ER'), _col(hd, 'CANT. DAO', 'DETALLE_DE_CALCULOS', 'ER')
    receta, cur = {}, None
    for r in det[1:]:
        a = txt(r[0])
        if re.match(r'^\d+$', a):
            cur = a
            receta[cur] = {'pin': 0.0, 'qui': 0.0, 'mod': 0.0, 'cf_pieza': 0.0, 'piezas': 0.0}
            continue
        if cur is None or not txt(r[1]):
            continue
        lab = txt(r[1]).upper()
        if lab.startswith('COSTO FIJO X PIEZA'):
            receta[cur]['cf_pieza'] = _num(r[2]); continue
        if lab.startswith(('PIEZAS X', 'VEL.')):
            continue
        receta[cur]['piezas'] = max(receta[cur]['piezas'], _num(r[c_cm]) + _num(r[c_cd]))
        receta[cur]['pin'] += sum(_num(r[c]) for c in c_pin if c < len(r))
        receta[cur]['qui'] += sum(_num(r[c]) for c in c_qui if c < len(r))
        receta[cur]['mod'] += sum(_num(r[c]) for c in c_mod if c < len(r))

    # Filas de detalle: ingreso del cliente repartido por m²; costo por receta
    m2_cli = defaultdict(float)
    for s in superficie:
        m2_cli[s['cod_cliente']] += s['m2']
    filas, sin_receta = [], []
    for s in superficie:
        rc = receta.get(s['articulo'])
        if rc is None:
            sin_receta.append(s['articulo'])
            rc = {'pin': 0, 'qui': 0, 'mod': 0, 'cf_pieza': 0, 'piezas': 0}
        tot_cli = clientes.get(s['cod_cliente'], {}).get('total', 0.0)
        share = s['m2'] / m2_cli[s['cod_cliente']] if m2_cli[s['cod_cliente']] else 0
        ing = tot_cli * share
        filas.append({'periodo': periodo, 'cod_cliente': s['cod_cliente'], 'cliente': nombres.get(s['cod_cliente'], ''),
                      'articulo': s['articulo'], 'descripcion': s['descripcion'], 'm2': round(s['m2'], 4),
                      'ingresos': round(ing, 2), 'cv_pintura': round(rc['pin'], 2), 'cv_quimicos': round(rc['qui'], 2),
                      'cv_mod': round(rc['mod'], 2), 'cf_fx': round(rc['cf_pieza'] * (s['piezas'] if s.get('piezas') else rc['piezas']), 2),
                      'sin_receta': 1 if s['articulo'] not in receta else 0})
    for cod, v in clientes.items():          # facturó pero no pintó en el mes
        if m2_cli.get(cod, 0) == 0 and v['total']:
            filas.append({'periodo': periodo, 'cod_cliente': cod, 'cliente': nombres.get(cod, ''), 'articulo': '',
                          'descripcion': 'SIN SUPERFICIE EN EL MES', 'm2': 0, 'ingresos': round(v['total'], 2),
                          'cv_pintura': 0, 'cv_quimicos': 0, 'cv_mod': 0, 'cf_fx': 0, 'sin_receta': 0})

    # Controles internos del ER
    chk = []
    def ctrl(nombre, a, b, tol=2.0):
        ok = a is not None and b is not None and abs(a - b) <= tol
        chk.append({'control': nombre, 'calculado': round(a or 0, 2), 'reporte': round(b or 0, 2), 'ok': ok})
    ctrl('Ingresos = suma clientes', sum(v['total'] for v in clientes.values()), totales.get('ingresos'))
    ctrl('Pintura Fx = suma receta', sum(f['cv_pintura'] for f in filas), totales['cv_pintura'], 5)
    ctrl('Químicos Fx = suma receta', sum(f['cv_quimicos'] for f in filas), totales['cv_quimicos'], 5)
    ctrl('MOD Fx = suma receta', sum(f['cv_mod'] for f in filas), totales['cv_mod'], 5)
    ctrl('CF Fx = costo fijo por pieza × piezas', sum(f['cf_fx'] for f in filas), totales['cf'], 5)
    ctrl('Ingresos repartidos = ingresos', sum(f['ingresos'] for f in filas), totales.get('ingresos'), 5)
    return {'emision': emision, 'empresa': {'m2': round(sum(s['m2'] for s in superficie), 4), **totales,
                                             'ingresos_malbec': sum(v['malbec'] for v in clientes.values())},
            'filas': filas, 'sin_receta': sin_receta, 'controles': chk, 'clientes': nombres}

# ════════════════════════════ Extractor WEG (pintura real) ════════════════════════════
# Artículos de WEG que no son pintura en polvo (se excluyen de los kg; el $ se informa aparte)
# Criterio acordado: lo que se compra una vez (equipos, pistolas) es INVERSIÓN y queda fuera del resultado;
# lo que se repite (repuestos) es COSTO FIJO; los químicos comprados a WEG siguen dentro de pintura, como hace Luciano.
INVERSION = re.compile(r'\bEQUIPO\b|PISTOLA|CABINA|HORNO|COMPRESOR', re.I)
REPUESTO = re.compile(r'ROD HOLDER|BOQUILLA|REPUESTO|FILTRO|MANGUERA|CABEZAL|ELECTRODO|REPARACI|SERVICE|MANTENIM', re.I)
QUIM_WEG = re.compile(r'CHEMISA|FOSFO|DESENGRAS|ACTIFOS|ACELERANTE|ULTRAL', re.I)

def control_fechas(rows, hi, c_fecha, periodo, fuente):
    """Compara el mes del nombre con la fecha de cada comprobante (FC/CC/NC/ND). Si más de la mitad son de otro mes,
    el archivo se rechaza: casi seguro está mal nombrado. Si son menos, queda un aviso. Sin fechas no hay control."""
    en_mes, otros = 0, defaultdict(int)
    for r in rows[hi + 1:]:
        f = r[c_fecha] if c_fecha < len(r) else None
        if isinstance(f, dt.datetime) and re.match(r'^(FC|CC|NC|ND)-', txt(r[0]).upper()):
            p = f'{f.year}-{f.month:02d}'
            if p == periodo:
                en_mes += 1
            else:
                otros[p] += 1
    fuera = sum(otros.values())
    if fuera > en_mes:
        comun = max(otros, key=otros.get)
        raise ErrorCarga(f'el nombre del {fuente} indica {periodo[5:]}/{periodo[:4]}, pero {fuera} de {fuera + en_mes} comprobantes '
                         f'tienen fecha de otro mes (la mayoría de {comun[5:]}/{comun[:4]})',
                         'Corregí el nombre del archivo o subí el del mes correcto.')
    return [f'{fuera} comprobantes con fecha de otro mes'] if fuera else []

def extraer_weg(libro, periodo):
    rows = next(iter(libro.values()), [])
    emision = fecha_emision(rows)
    items = []           # (codigo, descripcion, cantidad, importe_neto)
    formato = None
    hdr_i = next((i for i, r in enumerate(rows[:6]) if r and txt(r[0]).upper() == 'CPTE.'), None)
    if hdr_i is None:
        raise ErrorCarga('no encuentro el encabezado de comprobantes (celda "CPTE.") en las primeras filas del WEG. '
                         'Puede estar vacío o ser otro tipo de archivo')
    hdr = [txt(x).upper() for x in rows[hdr_i]]
    avisos = control_fechas(rows, hdr_i, hdr.index('FECHA'), periodo, 'WEG') if 'FECHA' in hdr else []
    if 'DESCRIPCIÓN' in hdr and 'CANT.' in hdr:
        # Formato A (2025-01 a 2025-06): listado de comprobantes con renglones
        formato = 'A_comprobantes'
        c_cod, c_desc, c_cant = _col(hdr, 'CÓD./MED. PAGO', 'principal', 'WEG'), hdr.index('DESCRIPCIÓN'), hdr.index('CANT.')
        _col(hdr, 'TOTAL', 'principal', 'WEG')
        c_tot = len(hdr) - 1 - hdr[::-1].index('TOTAL')          # TOTAL del renglón (la última columna TOTAL)
        for r in rows[hdr_i + 1:]:
            cp = txt(r[0]).upper()
            if not re.match(r'^(FC|CC|NC|ND)-', cp):
                continue                                          # PA = pagos, no son compras
            signo = -1 if cp.startswith(('CC', 'NC')) else 1
            items.append((txt(r[c_cod]), txt(r[c_desc]), signo * _num(r[c_cant]), signo * _num(r[c_tot])))
    else:
        # Formato B (2025-07) y C (2025-08 en adelante): "COMPRAS POR ARTICULO" con subtotal por artículo
        formato = 'C_por_articulo' if any('COMPRAS POR ARTICULO' in txt(r[0]).upper() for r in rows[:3]) else 'B_por_articulo'
        for r in rows[hdr_i + 1:]:
            a, d = txt(r[0]), txt(r[1])
            if re.match(r'^\d{4,}$', a) and d.startswith('('):
                desc = re.sub(r'^\(\d+\)\s*', '', d)
                importe = _num(r[8]) if len(r) > 8 and r[8] not in (None, '') else _num(r[7])
                items.append((a, desc, _num(r[3]), importe))
    if not items:
        raise ErrorCarga('el WEG no tiene renglones de compra reconocibles')
    rep = [i for i in items if REPUESTO.search(i[1])]                       # reparaciones y repuestos: costo fijo
    inv = [i for i in items if INVERSION.search(i[1]) and i not in rep]
    quim = [i for i in items if QUIM_WEG.search(i[1])]
    pintura = [i for i in items if i not in inv and i not in rep and i not in quim]
    return {'emision': emision, 'formato': formato, 'avisos': avisos,
            'empresa': {'weg_total': round(sum(i[3] for i in items), 2),
                        'weg_pintura': round(sum(i[3] for i in pintura), 2),
                        'weg_quimicos': round(sum(i[3] for i in quim), 2),
                        'weg_inversion': round(sum(i[3] for i in inv), 2),
                        'weg_repuestos': round(sum(i[3] for i in rep), 2),
                        'detalle_no_pintura': [f'{i[1]} ({i[2]:g} u.): {i[3]:,.0f}' for i in inv + rep],
                        'kg_pintura_weg': round(sum(i[2] for i in pintura), 2),
                        'unidades_weg_total': round(sum(i[2] for i in items), 2)},
            'items': [{'codigo': c, 'descripcion': d, 'cantidad': q, 'importe': imp} for c, d, q, imp in items]}

# ════════════════════════════ Extractor COMPRAS ════════════════════════════
PROV_PINTURA = re.compile(r'EMPRENDIMIENTOS', re.I)
PROV_QUIMICOS = re.compile(r'WHOR', re.I)
# Reglas por proveedor y ajustes por comprobante. Se pueden editar en <base>/reglas_compras.csv y
# <base>/ajustes_comprobantes.csv sin tocar el código. Categorías: pintura, quimicos, gasto_fijo, excluir.
REGLAS_DEFECTO = [('INDELPLAST', 'a_confirmar', 'Proveedor 30170 con facturas de $13 a $40 M que Luciano no cuenta como gasto. Pendiente de confirmar qué es.')]
AJUSTES_DEFECTO = [('FC-A-3-11677', 'inversion', 'Jun-25: 2 EQUIPOS DE PINTURA EN POLVO KCI ($13,2 M). Compra única: inversión.')]
REGLAS, AJUSTES = list(REGLAS_DEFECTO), {c: (cat, n) for c, cat, n in AJUSTES_DEFECTO}

def cargar_reglas(base_dir):
    global REGLAS, AJUSTES
    pr = os.path.join(base_dir, 'reglas_compras.csv'); pa = os.path.join(base_dir, 'ajustes_comprobantes.csv')
    if not os.path.exists(pr):
        with open(pr, 'w', newline='', encoding='utf-8') as f:
            w = csv.writer(f); w.writerow(['patron_proveedor', 'categoria', 'nota']); w.writerows(REGLAS_DEFECTO)
    if not os.path.exists(pa):
        with open(pa, 'w', newline='', encoding='utf-8') as f:
            w = csv.writer(f); w.writerow(['comprobante', 'categoria', 'nota']); w.writerows(AJUSTES_DEFECTO)
    with open(pr, encoding='utf-8') as f:
        REGLAS = [(r['patron_proveedor'], r['categoria'], r.get('nota', '')) for r in csv.DictReader(f) if r.get('patron_proveedor')]
    with open(pa, encoding='utf-8') as f:
        AJUSTES = {r['comprobante'].strip().upper(): (r['categoria'], r.get('nota', '')) for r in csv.DictReader(f) if r.get('comprobante')}

def categoria_compra(cpte, prov):
    if cpte.upper() in AJUSTES:
        return AJUSTES[cpte.upper()][0]
    for pat, cat, _ in REGLAS:
        if re.search(pat, prov, re.I):
            return cat
    if PROV_PINTURA.search(prov): return 'pintura'
    if PROV_QUIMICOS.search(prov): return 'quimicos'
    return 'gasto_fijo'

def extraer_compras(libro, periodo):
    rows = next(iter(libro.values()), [])
    emision = fecha_emision(rows)
    hi = _fila(rows[:6], lambda r: txt(r[0]).upper() == 'CPTE.', 'el encabezado de comprobantes (celda "CPTE.")',
               'principal', 'COMPRAS')
    h = [txt(x).upper() for x in rows[hi]]
    c = {k: _col(h, k, 'principal', 'COMPRAS') for k in ('CPTE.', 'FECHA', 'NETO', 'CONTACTO', 'CONCEPTO', 'TOTAL')}
    avisos = control_fechas(rows, hi, c['FECHA'], periodo, 'COMPRAS')
    tot = defaultdict(float); por_prov = defaultdict(float)
    for r in rows[hi + 1:]:
        cp = txt(r[c['CPTE.']]).upper()
        if not re.match(r'^(FC|CC|NC|ND)-', cp):
            continue                                              # PA (pagos) y totales no son gasto
        neto = -_num(r[c['NETO']])                                # el sistema exporta las facturas en negativo
        prov = txt(r[c['CONTACTO']])
        cat = categoria_compra(cp, prov)
        if cat == 'pintura':
            tot['compras_pintura'] += neto
        elif cat == 'quimicos':
            tot['quimicos_real'] += neto
        elif cat == 'excluir':
            tot['excluido'] += neto
        elif cat == 'a_confirmar':
            tot['a_confirmar'] += neto
        elif cat == 'inversion':
            if not PROV_PINTURA.search(prov):       # las de Emprendimientos ya las detecta WEG
                tot['inversion'] += neto
        else:
            tot['cf_gastos_real'] += neto
            por_prov[re.sub(r'^\(\d+\)\s*', '', prov)] += neto

    return {'emision': emision, 'empresa': {k: round(v, 2) for k, v in tot.items()},
            'gastos_por_proveedor': {k: round(v, 2) for k, v in sorted(por_prov.items(), key=lambda x: -x[1])},
            'avisos': avisos}

# ════════════════════════════ Extractor RESUMEN (anual o interanual) ════════════════════════════
# Conceptos que solo existen en el Resumen y se toman de ahí: (hoja, rótulo) -> campo
RESUMEN_CAMPOS = {
    ('Resumen', 'M2'): 'res_m2', ('Resumen', 'Ingresos'): 'res_ingresos',
    ('Resumen', 'CV (Fx)'): 'res_cv_fx', ('Resumen', 'CV ®'): 'res_cv_real',
    ('Resumen', 'CF (Fx)'): 'res_cf_fx', ('Resumen', 'CF ®'): 'res_cf_real',
    ('Resumen', 'Retiros Socios'): 'retiros_socios',
    ('CV', 'Pintura ®'): 'res_pintura_real', ('CV', 'Quimicos ®'): 'res_quimicos_real_cv', ('CV', 'MOD ®'): 'cv_mod_real',
    ('Quimicos', 'Quimicos ®'): 'res_quimicos_real', ('Quimicos', 'Quimicos (Fx)'): 'res_quimicos_fx',
    ('Pintura', 'Pintura (Fx)'): 'res_pintura_fx', ('Pintura', 'Kg DE PINTURA ®'): 'kg_pintura_real',
    ('Pintura', 'Kg DE PINTURA (Fx)'): 'kg_pintura_fx', ('Pintura', 'Lts DE QUIMICOS'): 'lts_quimicos',
    ('Pintura', '$ Pagados ®'): 'pintura_pagada_real',
    ('MOD', 'MOD (Fx)'): 'res_mod_fx',
    ('CF (Calculo)', 'CF ®'): 'res_cf_gastos_real', ('CF (Calculo)', 'Cargas Sociales'): 'cargas_sociales',
}

def _columnas_mes(rows, anio_archivo):
    """Ubica la fila de meses y devuelve (fila, {col: 'AAAA-MM'}). Acepta fechas o 'Ene'..'Dic' con el año arriba."""
    for i, r in enumerate(rows[:8]):
        fechas = {j: f'{v.year}-{v.month:02d}' for j, v in enumerate(r) if isinstance(v, dt.datetime)}
        if len(fechas) >= 3:
            return i, fechas
        abrev = {j: MESES[txt(v).lower()[:3]] for j, v in enumerate(r) if txt(v).lower()[:3] in MESES and len(txt(v)) <= 4}
        if len(abrev) >= 3:
            out = {}
            for j, m in abrev.items():
                y = None
                for k in range(i - 1, -1, -1):          # año en alguna fila de arriba
                    v = rows[k][j] if j < len(rows[k]) else None
                    if isinstance(v, (int, float)) and 2000 < v < 2100:
                        y = int(v); break
                out[j] = f'{y or anio_archivo}-{m:02d}'
            return i, out
    return None, {}

# Hojas del Resumen y qué se pierde si faltan. Sin la hoja "Resumen" no hay m² ni ingresos oficiales: se rechaza.
HOJAS_RESUMEN = {
    'CV': 'pintura, químicos y mano de obra reales; el mes queda incompleto y el tablero usa el costo teórico',
    'CF (Calculo)': 'gastos fijos reales, sueldos y cargas sociales; el mes queda incompleto y el tablero usa el costo teórico',
    'Pintura': 'pintura teórica y kg de pintura; se usa el ER del mes',
    'Quimicos': 'el control de químicos reales',
    'MOD': 'la mano de obra teórica; se usa el ER del mes',
    'Cliente $': 'el control de ingresos por cliente',
    'Cliente m2': 'el control de m² por cliente',
}

def extraer_resumen(libro, nombre):
    anio = int(re.search(r'(20\d{2})', os.path.basename(nombre)).group(1)) if re.search(r'(20\d{2})', os.path.basename(nombre)) else None
    datos = defaultdict(dict)       # periodo -> campo -> valor
    clientes_m = defaultdict(dict)  # (periodo, cliente) -> {'ingresos','m2'}
    avisos = []
    if 'Resumen' not in libro:
        raise ErrorCarga('al Resumen le falta la hoja "Resumen" (m², ingresos y retiros)',
                         'Revisá que la hoja no se haya renombrado o borrado y volvé a guardar el archivo.')
    for hoja, pierde in HOJAS_RESUMEN.items():
        if hoja not in libro:
            avisos.append(f'El Resumen ({os.path.basename(nombre)}) no trae la hoja "{hoja}": faltan {pierde}.')
    for hoja, rows in libro.items():
        hi, cols = _columnas_mes(rows, anio)
        if hi is None:
            continue
        # sueldos: cualquier fila 'Sueldo ...' de CF (Calculo)
        for r in rows[hi + 1:]:
            rot = txt(r[1]) if len(r) > 1 else ''
            if not rot:
                continue
            campo = RESUMEN_CAMPOS.get((hoja, rot))
            if hoja == 'CF (Calculo)' and rot.lower().startswith('sueldo'):
                for j, p in cols.items():
                    datos[p]['sueldos'] = datos[p].get('sueldos', 0.0) + _num(r[j] if j < len(r) else 0)
                continue
            if hoja in ('Cliente $', 'Cliente m2') and rot.lower() not in ('nombre',):
                k = 'ingresos' if hoja == 'Cliente $' else 'm2'
                for j, p in cols.items():
                    v = _num(r[j] if j < len(r) else 0)
                    if v:
                        clientes_m[(p, rot)][k] = v
                continue
            if campo:
                for j, p in cols.items():
                    datos[p][campo] = _num(r[j] if j < len(r) else 0)
    if not datos or not any('res_ingresos' in d for d in datos.values()):
        raise ErrorCarga('no encuentro la fila de meses ni los ingresos en la hoja "Resumen"',
                         'Revisá que la hoja tenga los meses arriba (fechas o Ene…Dic) y la fila "Ingresos".')
    return {'empresa_por_mes': {p: {k: round(v, 2) for k, v in d.items()} for p, d in datos.items()},
            'clientes_por_mes': [{'periodo': p, 'cliente': c, **v} for (p, c), v in clientes_m.items()],
            'avisos': avisos}

# ════════════════════════════ Escritura segura y bloqueo ════════════════════════════
def escribir_atomico(path, escribir, newline=None):
    """Escribe en un temporal y lo renombra al final: si la corrida se corta, queda el archivo anterior entero."""
    tmp = f'{path}.tmp{os.getpid()}'
    try:
        with open(tmp, 'w', encoding='utf-8', newline=newline) as f:
            escribir(f)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)

class CargaEnCurso(Exception):
    pass

class Bloqueo:
    """Evita dos corridas a la vez sobre la misma base (p. ej. la programada y la que dispara un archivo nuevo).
    Si quedó un bloqueo de una corrida que se cortó hace más de VENCE_MIN minutos, se descarta."""
    VENCE_MIN = 60

    def __init__(self, base_dir):
        os.makedirs(base_dir, exist_ok=True)
        self.path = os.path.join(base_dir, '.carga_en_curso')
        self.viejo = None

    def __enter__(self):
        for _ in range(2):
            try:
                fd = os.open(self.path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                with os.fdopen(fd, 'w') as f:
                    f.write(f'{os.getpid()} {dt.datetime.now().isoformat(timespec="seconds")}')
                return self
            except FileExistsError:
                edad = (dt.datetime.now().timestamp() - os.path.getmtime(self.path)) / 60
                if edad < self.VENCE_MIN:
                    raise CargaEnCurso(f'Ya hay una carga en curso (empezó hace {edad:.0f} min). No se hizo nada.')
                self.viejo = edad
                os.remove(self.path)   # bloqueo de una corrida que se cortó
        raise CargaEnCurso('No se pudo tomar el bloqueo de la carga.')

    def __exit__(self, *exc):
        try:
            os.remove(self.path)
        except FileNotFoundError:
            pass

def mensaje_usuario(e, tipo):
    """Traduce cualquier excepción a un mensaje que Luciano pueda entender y accionar."""
    if isinstance(e, ErrorCarga):
        return str(e)
    t = f'{type(e).__name__}: {e}'
    if 'xlrd' in type(e).__module__ or 'BOF' in t or 'zip' in t.lower() or 'InvalidFileException' in t or 'corrupt' in t.lower():
        return 'el archivo está dañado o no es un Excel. Volvé a exportarlo desde el sistema y guardalo de nuevo en la carpeta.'
    if isinstance(e, PermissionError):
        return 'no se pudo abrir el archivo (puede estar abierto en Excel o sin permisos). Cerralo y se carga en la próxima corrida.'
    return (f'error inesperado al leer el {tipo}. Avisale al equipo: el detalle técnico quedó en registro_cargas.csv.')

# ════════════════════════════ Base histórica (upsert por mes) ════════════════════════════
class Base:
    """Una carpeta con un JSON por fuente y mes: base/ER/2026-08.json, base/RESUMEN/2026-08.json…
    Regla: nunca se borra un mes. Si llega el mismo mes otra vez, gana el archivo con fecha de emisión más nueva
    (o, si no tiene, el de fecha de modificación más nueva)."""
    def __init__(self, path):
        self.path = path
        os.makedirs(path, exist_ok=True)
        self.log_path = os.path.join(path, 'registro_cargas.csv')
        self.problemas = []   # registros de la base que se encontraron dañados
        self.pendientes = []  # meses incompletos de la última corrida (regla de los 4 archivos)

    # Errores ya informados (por huella del archivo): sirven para avisar a n8n solo cuando aparece uno nuevo
    def errores_vistos(self):
        p = os.path.join(self.path, 'errores_vistos.json')
        try:
            with open(p, encoding='utf-8') as f:
                return json.load(f)
        except (OSError, ValueError):
            return {}

    def guardar_errores_vistos(self, vistos):
        escribir_atomico(os.path.join(self.path, 'errores_vistos.json'), lambda f: json.dump(vistos, f, ensure_ascii=False, indent=1))

    def _p(self, fuente, periodo):
        d = os.path.join(self.path, fuente); os.makedirs(d, exist_ok=True)
        return os.path.join(d, f'{periodo}.json')

    def leer(self, fuente, periodo):
        p = self._p(fuente, periodo)
        if os.path.exists(p):
            try:
                with open(p, encoding='utf-8') as f:
                    return json.load(f)
            except (json.JSONDecodeError, UnicodeDecodeError):
                # Un JSON roto (p. ej. de una versión vieja que se cortó al escribir) no debe frenar todo:
                # se aparta y el mes se vuelve a cargar desde el Excel en la próxima corrida.
                destino = p + '.corrupto'
                os.replace(p, destino)
                self.problemas.append({'archivo': f'base/{fuente}/{periodo}.json', 'tipo': fuente, 'periodo': periodo,
                                       'error': f'el registro guardado de {fuente} {periodo} estaba dañado y se apartó como '
                                                f'{os.path.basename(destino)}. Se vuelve a cargar desde el Excel si sigue en la carpeta.'})
        return None

    def upsert(self, fuente, periodo, payload, meta):
        actual = self.leer(fuente, periodo)
        clave_n = meta.get('emision') or meta.get('mtime')
        if actual:
            clave_a = actual['_meta'].get('emision') or actual['_meta'].get('mtime')
            if actual['_meta'].get('sha1') == meta['sha1']:
                return 'sin cambios'
            if clave_a and clave_n and clave_n < clave_a:
                return 'ignorado (hay uno más nuevo)'
        escribir_atomico(self._p(fuente, periodo), lambda f: json.dump({'_meta': meta, **payload}, f, ensure_ascii=False, default=str))
        return 'reemplazado' if actual else 'nuevo'

    def periodos(self, fuente):
        d = os.path.join(self.path, fuente)
        return sorted(f[:-5] for f in os.listdir(d) if f.endswith('.json')) if os.path.isdir(d) else []

    def registrar(self, filas):
        nuevo = not os.path.exists(self.log_path)
        with open(self.log_path, 'a', newline='', encoding='utf-8') as f:
            w = csv.writer(f)
            if nuevo:
                w.writerow(['procesado', 'archivo', 'tipo', 'periodo', 'resultado', 'detalle'])
            w.writerows(filas)

def sha1(path):
    h = hashlib.sha1()
    with open(path, 'rb') as f:
        for b in iter(lambda: f.read(1 << 16), b''):
            h.update(b)
    return h.hexdigest()

FUENTES_MES = ('ER', 'WEG', 'COMPRAS', 'RESUMEN')
NOMBRE_FUENTE = {'ER': 'ER', 'WEG': 'WEG', 'COMPRAS': 'COMPRAS', 'RESUMEN': 'Resumen'}

def listar_excel(entrada):
    """Excel de la carpeta de entrada (con subcarpetas). Se ignoran las carpetas que empiezan con '_' o '.'
    (por ejemplo _reemplazados/, donde quedan las versiones anteriores que reemplazó una carga)."""
    out = []
    for dp, dirs, fs in os.walk(entrada):
        dirs[:] = [d for d in dirs if not d.startswith(('_', '.'))]
        out += [os.path.join(dp, f) for f in fs]
    return sorted(out)

def leer_archivo(path, idf, ahora):
    """Lee y valida un Excel. Devuelve (meta, resultado) o lanza la excepción del archivo."""
    meta = {'archivo': os.path.basename(path), 'sha1': sha1(path),
            'mtime': dt.datetime.fromtimestamp(os.path.getmtime(path)).isoformat(timespec='seconds'), 'procesado': ahora}
    del _TEXTOS_COMO_CERO[:]
    libro = leer_libro(path)
    if idf['tipo'] == 'ER':
        r = extraer_er(libro, idf['periodo'])
    elif idf['tipo'] == 'WEG':
        r = extraer_weg(libro, idf['periodo'])
    elif idf['tipo'] == 'COMPRAS':
        r = extraer_compras(libro, idf['periodo'])
    else:
        r = extraer_resumen(libro, path)
    if _TEXTOS_COMO_CERO:
        ej = ', '.join(f'"{t}"' for t in dict.fromkeys(_TEXTOS_COMO_CERO))
        r.setdefault('avisos', []).append(f'{meta["archivo"]}: {len(_TEXTOS_COMO_CERO)} celda(s) con texto donde se esperaba '
                                          f'un número ({ej[:120]}). Se tomaron como 0: revisá esas celdas.')
    return meta, r

def procesar_carpeta(entrada, base):
    """Paso 1: lee y valida todos los Excel. Paso 2: aplica la REGLA DEL MES COMPLETO (v1.2):
    un mes se carga o se actualiza solo si en la carpeta están sus cuatro archivos (ER, WEG, COMPRAS y un Resumen
    que traiga la columna de ese mes), todos válidos. Mientras falte alguno, el mes queda pendiente y la base no se toca."""
    ahora = dt.datetime.now().isoformat(timespec='seconds')
    log, errores, leidos = [], [], []
    for path in listar_excel(entrada):
        idf = identificar(path)
        if not idf:
            continue
        try:
            meta, r = leer_archivo(path, idf, ahora)
            leidos.append((idf, meta, r))
        except Exception as e:  # un archivo roto no frena al resto
            tecnico = f'{type(e).__name__}: {e}'
            huella = ''
            try:
                huella = sha1(path)
            except OSError:
                pass
            errores.append({'archivo': os.path.basename(path), 'tipo': idf['tipo'], 'periodo': idf.get('periodo') or '',
                            'sha1': huella, 'error': mensaje_usuario(e, idf['tipo']), 'tecnico': tecnico})
            log.append([ahora, os.path.basename(path), idf['tipo'], idf.get('periodo') or '', 'ERROR', tecnico])

    # Qué fuentes válidas hay para cada mes
    presentes = defaultdict(set)
    for idf, meta, r in leidos:
        if idf['tipo'] == 'RESUMEN':
            for p in r['empresa_por_mes']:
                presentes[p].add('RESUMEN')
        else:
            presentes[idf['periodo']].add(idf['tipo'])
    completos = {p for p, f in presentes.items() if set(FUENTES_MES) <= f}

    def igual_a_base(fuente, p, sha):
        actual = base.leer(fuente, p)
        return bool(actual) and actual['_meta'].get('sha1') == sha

    pendientes = {}
    for idf, meta, r in leidos:
        if idf['tipo'] == 'RESUMEN':
            ultimo = max(r['empresa_por_mes'])
            for p, d in r['empresa_por_mes'].items():
                if p not in completos:
                    # Columnas del Resumen de meses sin sus archivos mensuales: no se cargan (no se informa cada una)
                    log.append([ahora, meta['archivo'], 'RESUMEN', p, 'pendiente (mes incompleto)', ''])
                    continue
                cl = [c for c in r['clientes_por_mes'] if c['periodo'] == p]
                # Entre dos Resumen gana el que llega a un mes más reciente (el interanual 2026 le gana al anual 2025)
                res = base.upsert('RESUMEN', p, {'empresa': d, 'clientes': cl, 'avisos': r.get('avisos', [])},
                                  {**meta, 'emision': f'{ultimo}-28T00:00:00', 'cubre_hasta': ultimo})
                log.append([ahora, meta['archivo'], 'RESUMEN', p, res, '; '.join(r.get('avisos', []))])
            continue
        p = idf['periodo']
        if p not in completos:
            if igual_a_base(idf['tipo'], p, meta['sha1']):
                log.append([ahora, meta['archivo'], idf['tipo'], p, 'sin cambios', ''])
                continue
            pendientes.setdefault(p, set()).add(meta['archivo'])
            log.append([ahora, meta['archivo'], idf['tipo'], p, 'pendiente (mes incompleto)',
                        'faltan: ' + ', '.join(NOMBRE_FUENTE[f] for f in FUENTES_MES if f not in presentes[p])])
            continue
        meta['emision'] = r.get('emision')
        res = base.upsert(idf['tipo'], p, r, meta)
        bad = [c['control'] for c in r.get('controles', []) if not c['ok']]
        log.append([ahora, meta['archivo'], idf['tipo'], p, res, '; '.join(bad + r.get('avisos', []))])

    # Meses nuevos que solo trae el Resumen (todavía sin ningún archivo mensual)
    cargados = set(base.periodos('ER')) | set(base.periodos('RESUMEN'))
    ultimo_cargado = max(cargados) if cargados else ''
    for p, f in presentes.items():
        if p not in completos and p not in pendientes and p not in cargados and f == {'RESUMEN'} and p > ultimo_cargado:
            pendientes[p] = set()
    base.pendientes = [{'periodo': p, 'llegaron': [NOMBRE_FUENTE[f] for f in FUENTES_MES if f in presentes[p]],
                        'faltan': [NOMBRE_FUENTE[f] for f in FUENTES_MES if f not in presentes[p]],
                        'ya_cargado': p in cargados}
                       for p in sorted(pendientes)]

    for pr in base.problemas:   # registros dañados de la base encontrados al leer
        errores.append({**pr, 'sha1': '', 'tecnico': pr['error']})
        log.append([ahora, pr['archivo'], pr['tipo'], pr['periodo'], 'ERROR', pr['error']])
    base.problemas = []
    base.registrar(log)
    return log, errores

def clasificar_errores(base, errores):
    """Marca cada error como nuevo o ya informado (por huella del archivo + mensaje) y guarda la lista vigente.
    Un Excel roto que sigue en la carpeta se sigue mostrando en el tablero, pero no vuelve a disparar la alerta de n8n."""
    vistos = base.errores_vistos()
    ahora = dt.datetime.now().isoformat(timespec='seconds')
    vigentes = {}
    for e in errores:
        clave = f"{e.get('sha1') or e['archivo']}|{e['error']}"
        e['nuevo'] = clave not in vistos
        e['desde'] = vistos.get(clave, ahora)
        vigentes[clave] = e['desde']
    base.guardar_errores_vistos(vigentes)
    return errores

# ════════════════════════════ Armado de salidas para el tablero ════════════════════════════
DET_COLS = ['anio', 'mes', 'cod_cliente', 'cliente', 'articulo', 'descripcion', 'm2', 'ingresos',
            'cv_pintura', 'cv_quimicos', 'cv_mod', 'cf_fx', 'cmg', 'sin_receta']
RES_COLS = ['periodo', 'm2', 'ingresos', 'cv_pintura_fx', 'cv_quimicos_fx', 'cv_mod_fx', 'cf_fx',
            'cv_pintura_real', 'cv_quimicos_real', 'cv_mod_real', 'cf_gastos_real', 'sueldos', 'cargas_sociales',
            'retiros_socios', 'kg_pintura_real', 'kg_pintura_fx', 'lts_quimicos', 'pintura_pagada_real',
            'inversiones', 'egresos_a_confirmar', 'fuente_fx', 'fuente_real', 'completo']

NOMBRE_CAMPO = {'retiros_socios': 'Retiros de socios', 'cv_quimicos_real': 'Químicos reales', 'cv_mod_real': 'Mano de obra real',
                'cv_pintura_real': 'Pintura real', 'cf_gastos_real': 'Gastos fijos reales'}

def ars(v):
    return ('-' if v < 0 else '') + '$' + f'{abs(v):,.0f}'.replace(',', '.')

def armar_salidas(base, salida, tolerancia=0.01, errores=None):
    """Empresa: Resumen de Luciano primero (es el número oficial y ya trae su criterio de imputación);
    los archivos mensuales completan los meses que el Resumen no tiene y sirven de control.
    Detalle por artículo: ER del mes; si el ER fue re-emitido a otros precios, se ajusta para sumar lo mismo que el Resumen."""
    os.makedirs(salida, exist_ok=True)
    periodos = sorted(set(base.periodos('ER')) | set(base.periodos('RESUMEN')))
    filas, det = [], []
    calidad = {'generado': dt.datetime.now().isoformat(timespec='seconds'), 'version': VERSION, 'meses': {}}
    for p in periodos:
        er, weg, com, res = (base.leer(x, p) for x in ('ER', 'WEG', 'COMPRAS', 'RESUMEN'))
        rs = (res or {}).get('empresa', {})
        q = {'fuentes': [x for x, v in (('ER', er), ('WEG', weg), ('COMPRAS', com), ('RESUMEN', res)) if v],
             'faltantes': [], 'diferencias': [], 'avisos': [], 'notas': []}
        e = er['empresa'] if er else {}
        # Valores derivados de los archivos mensuales (control o respaldo)
        men = {'m2': e.get('m2'), 'ingresos': e.get('ingresos'), 'cv_pintura_fx': e.get('cv_pintura'),
               'cv_quimicos_fx': e.get('cv_quimicos'), 'cv_mod_fx': e.get('cv_mod'), 'cf_fx': e.get('cf'),
               'cv_pintura_real': weg['empresa']['weg_total'] if weg else None,
               'cv_quimicos_real': com['empresa'].get('quimicos_real', 0.0) if com else None,
               'cf_gastos_real': com['empresa'].get('cf_gastos_real', 0.0) if com else None,
               'kg_pintura_real': weg['empresa']['kg_pintura_weg'] if weg else None}
        # Químicos reales: ante dudas se sigue al Resumen (hoja CV, la que suma en sus totales) y se deja el aviso
        qr = rs.get('res_quimicos_real_cv')
        if rs and rs.get('res_quimicos_real') is not None and qr is not None and abs(qr - rs['res_quimicos_real']) > 1:
            extra = f" COMPRAS del mes registra {ars(men['cv_quimicos_real'])}." if men['cv_quimicos_real'] is not None else ''
            q['avisos'].append(f"Químicos reales: la hoja CV del Resumen dice {ars(qr)} y la hoja Quimicos {ars(rs['res_quimicos_real'])}. Se usa el del Resumen (hoja CV).{extra} Posible celda copiada.")
        ofi = {'m2': rs.get('res_m2'), 'ingresos': rs.get('res_ingresos'), 'cv_pintura_fx': rs.get('res_pintura_fx'),
               'cv_quimicos_fx': rs.get('res_quimicos_fx'), 'cv_mod_fx': rs.get('res_mod_fx'), 'cf_fx': rs.get('res_cf_fx'),
               'cv_pintura_real': rs.get('res_pintura_real'), 'cv_quimicos_real': qr, 'cf_gastos_real': rs.get('res_cf_gastos_real'),
               'kg_pintura_real': rs.get('kg_pintura_real')}
        row = {'periodo': p}
        for k in ofi:
            row[k] = ofi[k] if ofi[k] is not None else men.get(k)
        for k in ('cv_mod_real', 'sueldos', 'cargas_sociales', 'retiros_socios', 'kg_pintura_fx', 'lts_quimicos', 'pintura_pagada_real'):
            row[k] = rs.get(k)
        if row['m2'] is None or row['ingresos'] is None:
            continue
        row['inversiones'] = 0.0
        row['egresos_a_confirmar'] = round(com['empresa'].get('a_confirmar', 0.0), 2) if com else None
        explic = set()
        if weg and (weg['empresa'].get('weg_inversion') or weg['empresa'].get('weg_repuestos')) and row.get('cv_pintura_real') is not None:
            inv, rp, wt, pb = weg['empresa']['weg_inversion'], weg['empresa']['weg_repuestos'], weg['empresa']['weg_total'], row['cv_pintura_real']
            tol = max(1.0, 0.01 * wt)
            if abs(pb - wt) <= tol:
                incl_inv, incl_rep = inv > 0, rp > 0
            elif abs(pb - (wt - inv)) <= tol:
                incl_inv, incl_rep = False, rp > 0
                explic.add('cv_pintura_real')
            else:
                incl_inv = incl_rep = False
            row['cv_pintura_real'] = pb - (inv if incl_inv else 0) - (rp if incl_rep else 0)
            if incl_rep and row.get('cf_gastos_real') is not None:
                row['cf_gastos_real'] += rp
            if inv and not incl_inv and com and rs.get('res_cf_gastos_real') is not None:
                # ¿Luciano la cargó en costos fijos? (como en jun-25)
                if abs((rs['res_cf_gastos_real'] - com['empresa'].get('cf_gastos_real', 0)) - inv) <= max(1.0, 0.03 * inv):
                    row['cf_gastos_real'] -= inv
                    explic.add('cf_gastos_real')
                    q['notas'].append(f'Inversión de {ars(inv)} que el Resumen tenía en costos fijos: se sacó del resultado.')
            row['inversiones'] = inv
            if inv:
                q['notas'].append(f"Inversión en equipos: {ars(inv)} ({'; '.join(d for d in weg['empresa']['detalle_no_pintura'] if INVERSION.search(d) and not REPUESTO.search(d))}). Queda fuera de pintura y del resultado.")
            if rp and incl_rep:
                q['notas'].append(f'Repuestos y reparaciones de WEG por {ars(rp)}: pasan de pintura a costo fijo (se repiten en el año).')
        if com and com['empresa'].get('inversion'):
            row['inversiones'] += com['empresa']['inversion']
        if row['egresos_a_confirmar']:
            q['avisos'].append(f"Egresos a confirmar (INDELPLAST): {ars(row['egresos_a_confirmar'])}. No se restan del resultado hasta saber qué son.")
        row['fuente_fx'] = 'Resumen' if rs.get('res_cf_fx') is not None else 'ER'
        row['fuente_real'] = 'Resumen' if rs.get('res_pintura_real') is not None else '+'.join(x for x, v in (('WEG', weg), ('COMPRAS', com)) if v)
        if not er: q['faltantes'].append('ER (sin detalle por cliente y artículo)')
        if not weg: q['faltantes'].append('WEG')
        if not com: q['faltantes'].append('COMPRAS')
        if not res: q['faltantes'].append('Resumen (MOD real, sueldos, cargas sociales y retiros)')
        # Sueldos y cargas pueden venir en una sola fila total ("Sueldos y cargas"): entonces cargas = 0
        if row.get('sueldos') is not None and row.get('cargas_sociales') is None:
            row['cargas_sociales'] = 0.0
        req = ['cv_pintura_real', 'cv_quimicos_real', 'cv_mod_real', 'cf_gastos_real', 'sueldos', 'cargas_sociales', 'retiros_socios']
        row['completo'] = 1 if all(row.get(k) is not None for k in req) else 0
        # Conciliación mensual vs oficial
        nombres = {'m2': 'm² pintados', 'ingresos': 'Ingresos', 'cv_pintura_fx': 'Pintura teórica', 'cv_quimicos_fx': 'Químicos teóricos',
                   'cv_mod_fx': 'MOD teórica', 'cf_fx': 'CF teórico', 'cv_pintura_real': 'Pintura real (WEG)',
                   'cv_quimicos_real': 'Químicos reales (COMPRAS)', 'cf_gastos_real': 'Gastos fijos (COMPRAS)', 'kg_pintura_real': 'Kg de pintura (WEG)'}
        if rs:
            for k, n in nombres.items():
                a, b = men.get(k), ofi.get(k)
                if a is not None and b is not None and abs(a - b) > max(1.0, tolerancia * abs(b)):
                    q['diferencias'].append({'concepto': n, 'campo': k, 'archivo_mensual': round(a, 2), 'resumen': round(b, 2), 'dif': round(a - b, 2), 'explicada': k in explic})
        # Detalle por artículo con ajuste a los totales oficiales
        if er:
            factores = {}
            for kf, kr in (('cv_pintura', 'cv_pintura_fx'), ('cv_quimicos', 'cv_quimicos_fx'), ('cv_mod', 'cv_mod_fx'), ('cf_fx', 'cf_fx')):
                s_art = sum(f[kf] for f in er['filas'])
                factores[kf] = (row[kr] / s_art) if (s_art and row.get(kr) is not None and abs(row[kr] - s_art) > max(5.0, 0.005 * s_art)) else 1.0
            if any(abs(v - 1) > 1e-9 for v in factores.values()):
                emi = (er['_meta'].get('emision') or '')[:10]
                q['notas'].append(f'El ER de este mes fue emitido el {emi[8:10]}/{emi[5:7]}/{emi[:4]} con precios posteriores: su costo teórico no coincide con el Resumen. '
                                  'El detalle por artículo se ajustó para sumar lo mismo que el Resumen.')
                for d in q['diferencias']:
                    if d['campo'] in ('cv_pintura_fx', 'cv_quimicos_fx', 'cv_mod_fx', 'cf_fx'):
                        d['explicada'] = True
            for f in er['filas']:
                r2 = {'anio': p[:4], 'mes': int(p[5:]), **{k: f[k] for k in ('cod_cliente', 'cliente', 'articulo', 'descripcion', 'm2', 'ingresos', 'sin_receta')}}
                for kf in ('cv_pintura', 'cv_quimicos', 'cv_mod', 'cf_fx'):
                    r2[kf] = round(f[kf] * factores[kf], 2)
                r2['cmg'] = round(r2['ingresos'] - r2['cv_pintura'] - r2['cv_quimicos'] - r2['cv_mod'], 2)
                det.append(r2)
            if er['sin_receta']:
                q['sin_receta'] = er['sin_receta']
            q['controles_er'] = [c for c in er['controles'] if not c['ok']]
        if com and com['empresa'].get('excluido'):
            q['notas'].append(f"COMPRAS: {ars(com['empresa']['excluido'])} no se cuentan como gasto fijo por regla (ver reglas_compras.csv).")
        # Avisos que dejó la lectura de cada archivo del mes (hojas faltantes del Resumen, textos tomados como 0, fechas de otro mes)
        for nombre_f, fuente in (('ER', er), ('WEG', weg), ('COMPRAS', com), ('Resumen', res)):
            for a in (fuente or {}).get('avisos', []):
                if 'comprobantes con fecha de otro mes' in a:
                    q['notas'].append(f'{nombre_f}: {a}.')
                elif a not in q['avisos']:
                    q['avisos'].append(a)
        calidad['meses'][p] = q
        filas.append({k: (round(v, 2) if isinstance(v, float) else v) for k, v in row.items()})
    # Valores idénticos al centavo en meses distintos: suele ser una celda copiada
    vistos = defaultdict(list)
    for r in filas:
        for k in ('retiros_socios', 'cv_quimicos_real', 'cv_mod_real', 'cv_pintura_real', 'cf_gastos_real'):
            v = r.get(k)
            if v:
                vistos[(k, round(v, 2))].append(r['periodo'])
    for (k, v), ps in vistos.items():
        if len(ps) > 1:
            for p in ps:
                calidad['meses'][p]['avisos'].append(f'{NOMBRE_CAMPO[k]}: {ars(v)} es idéntico al centavo en {" y ".join(ps)}. Puede ser un valor aproximado o copiado; se usa el del Resumen.')
    for p, q in calidad['meses'].items():
        pend = [d for d in q['diferencias'] if not d['explicada']]
        q['estado'] = 'faltan_archivos' if q['faltantes'] else ('revisar' if (pend or q['avisos']) else 'ok')
    # Errores de carga vigentes: el tablero los muestra en la franja roja
    errores = list(errores or [])
    for pr in base.problemas:
        errores.append({**pr, 'sha1': '', 'tecnico': pr['error'], 'nuevo': True, 'desde': calidad['generado']})
    base.problemas = []
    calidad['errores_carga'] = [{k: e.get(k) for k in ('archivo', 'tipo', 'periodo', 'error', 'nuevo', 'desde')} for e in errores]
    calidad['pendientes'] = list(base.pendientes)

    def _csv(filas_, cols, fn=None):
        def escribir(f):
            w = csv.DictWriter(f, fieldnames=cols, extrasaction='ignore'); w.writeheader()
            w.writerows(fn(r) for r in filas_) if fn else w.writerows(filas_)
        return escribir
    escribir_atomico(os.path.join(salida, 'detalle_clientes_articulos.csv'), _csv(det, DET_COLS), newline='')
    escribir_atomico(os.path.join(salida, 'resumen_mensual.csv'),
                     _csv(filas, RES_COLS, lambda r: {k: ('' if r.get(k) is None else r.get(k)) for k in RES_COLS}), newline='')
    # calidad.json se escribe último: el tablero lo usa para saber cuándo fue la última carga completa
    escribir_atomico(os.path.join(salida, 'calidad.json'), lambda f: json.dump(calidad, f, ensure_ascii=False, indent=1))
    return det, filas, calidad


def main(argv=None):
    ap = argparse.ArgumentParser(description='Pipeline DAO SRL: Excel mensuales -> CSV del tablero')
    ap.add_argument('--entrada', required=True, help='Carpeta con los Excel de Luciano (puede tener subcarpetas)')
    ap.add_argument('--base', required=True, help='Carpeta de la base histórica (se crea si no existe)')
    ap.add_argument('--salida', required=True, help='Carpeta data/ que lee el tablero')
    a = ap.parse_args(argv)
    if not os.path.isdir(a.entrada):
        print(f'ERROR: no existe la carpeta de entrada {a.entrada}')
        return 1
    try:
        with Bloqueo(a.base) as bloqueo:
            if bloqueo.viejo:
                print(f'Aviso: había un bloqueo de una corrida que se cortó hace {bloqueo.viejo:.0f} min; se descartó.')
            base = Base(a.base)
            cargar_reglas(a.base)
            log, errores = procesar_carpeta(a.entrada, base)
            det, res, cal = armar_salidas(base, a.salida, errores=clasificar_errores(base, errores))
    except CargaEnCurso as e:
        print(e)
        return 0          # no es una falla: la otra corrida ya está trabajando
    errores = cal['errores_carga']
    nuevos = [e for e in errores if e['nuevo']]
    cambios = [l for l in log if l[4] in ('nuevo', 'reemplazado')]
    print(f'Archivos leídos: {len(log)} · nuevos/reemplazados: {len(cambios)} · errores: {len(errores)} ({len(nuevos)} nuevos)')
    for l in cambios:
        print(f'  {l[4]:<12} {l[2]:<8} {l[3]}  {l[1]}')
    for e in errores:
        print(f'  ERROR{" (nuevo)" if e["nuevo"] else " (ya informado)"} {e["archivo"]}: {e["error"]}')
    for pe in cal.get('pendientes', []):
        print(f'  PENDIENTE {pe["periodo"]}: llegaron {", ".join(pe["llegaron"]) or "ninguno"} · faltan {", ".join(pe["faltan"])} '
              f'(un mes se carga solo con los 4 archivos juntos)')
    print(f'Salida: {len(det)} filas de detalle · {len(res)} meses de empresa ({res[0]["periodo"] if res else "-"} a {res[-1]["periodo"] if res else "-"})')
    # 1 solo cuando aparece un error nuevo: un archivo roto que sigue en la carpeta no vuelve a disparar la alerta
    return 1 if nuevos else 0

if __name__ == '__main__':
    sys.exit(main())
