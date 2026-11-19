#!/usr/bin/env python3
"""
Carga de un mes completo — DAO SRL (v1.3)

Recibe los cuatro Excel de un mes (ER, WEG, COMPRAS y Resumen), los valida ANTES de tocar nada,
los deja en la carpeta de entrada y corre el pipeline. Es lo que usa el botón "Subir Excel del mes" del tablero.

Reglas:
  - Exactamente 4 archivos, .xls o .xlsx, hasta 20 MB en total.
  - Uno de cada tipo: ER, WEG, COMPRAS y Resumen (reconocidos por el nombre).
  - ER, WEG y COMPRAS del mismo mes; el Resumen tiene que traer la columna de ese mes.
  - Cada archivo se lee y se valida con las mismas reglas del pipeline. Si uno falla, no se guarda ninguno.
  - Las versiones anteriores de esos archivos se mueven a <entrada>/_reemplazados/<fecha>/ (no se borra nada).
  - No se aceptan archivos más viejos que los ya cargados: quedarían en la carpeta sin cambiar los datos.
  - Si el mes ya está cargado con otros archivos, se reemplaza solo con confirmación (--reemplazar o confirmar=True).
  - El resultado lista los otros meses que cambiaron con la carga (por ejemplo, porque el Resumen nuevo trae otro valor).

Uso:
  python cargar_mes.py --entrada <carpeta_excel> --base <carpeta_base> --salida <carpeta_data> [--reemplazar] ARCH1 ARCH2 ARCH3 ARCH4
Devuelve un JSON por la salida estándar y sale con 0 si el mes se cargó, 1 si se rechazó o falta confirmar.
"""
from __future__ import annotations
import argparse, contextlib, csv, datetime as dt, io, json, os, re, shutil, sys, time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import dao_pipeline as dp  # noqa: E402

MAX_MB = 20
ESPERA_BLOQUEO_SEG = 90
NOMBRE = dp.NOMBRE_FUENTE
MESES_L = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']

def mes_txt(p):
    return f'{MESES_L[int(p[5:]) - 1]} {p[:4]}' if p else ''

def nombre_base(n):
    """Nombre sin extensión, sin sufijos de copia ' (2)' / ' - copia' y en minúsculas: sirve para reconocer versiones del mismo archivo."""
    n = os.path.splitext(os.path.basename(n))[0].lower()
    n = re.sub(r'\s*(\(\d+\)|-\s*copia(\s*\(\d+\))?|copia)$', '', n)
    return re.sub(r'[\s_\-]+', '_', n).strip('_')

def rechazo(problemas, periodo=None):
    return {'ok': False, 'periodo': periodo, 'mes': mes_txt(periodo),
            'mensaje': f'No se cargó {("el mes de " + mes_txt(periodo)) if periodo else "ningún archivo"}. Corregí lo siguiente y volvé a subir los 4 archivos juntos.',
            'problemas': problemas, 'avisos': []}

def validar(archivos):
    """archivos: lista de (nombre_original, ruta_temporal). Devuelve (periodo, {tipo: (nombre, ruta, meta, r)}) o (periodo, [problemas])."""
    problemas = []
    if len(archivos) != 4:
        problemas.append(f'Se eligieron {len(archivos)} archivo(s). Hay que subir los 4 del mes juntos: ER, WEG, COMPRAS y Resumen.')
    total = sum(os.path.getsize(r) for _, r in archivos)
    if total > MAX_MB * 1024 * 1024:
        problemas.append(f'Los archivos pesan {total / 1048576:.1f} MB en total; el máximo es {MAX_MB} MB.')
    tipos = {}
    for nombre, ruta in archivos:
        if not re.search(r'\.xlsx?$', nombre, re.I):
            problemas.append(f'{nombre}: no es un Excel (.xls o .xlsx).')
            continue
        idf = dp.identificar(nombre)
        if not idf:
            problemas.append(f'{nombre}: no reconozco el tipo por el nombre. Tiene que empezar con ER, WEG o COMPRAS y el mes y año '
                             f'(por ejemplo "ER_09_2026.xlsx"), o contener "Resumen".')
            continue
        if idf['tipo'] in tipos:
            problemas.append(f'Hay dos archivos de tipo {NOMBRE[idf["tipo"]]}: {tipos[idf["tipo"]][0]} y {nombre}.')
            continue
        tipos[idf['tipo']] = (nombre, ruta, idf)
    faltan = [NOMBRE[t] for t in dp.FUENTES_MES if t not in tipos]
    if faltan and len(archivos) <= 4:
        problemas.append(f'Falta{"n" if len(faltan) > 1 else ""}: {", ".join(faltan)}.')
    meses = {tipos[t][2]['periodo'] for t in ('ER', 'WEG', 'COMPRAS') if t in tipos}
    if len(meses) > 1:
        problemas.append('Los archivos son de meses distintos: ' + ', '.join(f'{NOMBRE[t]} de {mes_txt(tipos[t][2]["periodo"])}'
                                                                         for t in ('ER', 'WEG', 'COMPRAS') if t in tipos) + '.')
    periodo = next(iter(meses)) if len(meses) == 1 else None
    if problemas:
        return periodo, problemas

    # Lectura y validación del contenido con las mismas reglas del pipeline
    ahora = dt.datetime.now().isoformat(timespec='seconds')
    leidos = {}
    for t, (nombre, ruta, idf) in tipos.items():
        try:
            # se valida con el nombre original (identificar y el año del Resumen salen del nombre)
            tmp_nombre = os.path.join(os.path.dirname(ruta), '_v_' + os.path.basename(nombre))
            shutil.copy2(ruta, tmp_nombre)
            try:
                meta, r = dp.leer_archivo(tmp_nombre, idf, ahora)
            finally:
                os.remove(tmp_nombre)
            if t == 'RESUMEN' and periodo not in r['empresa_por_mes']:
                cols = sorted(r['empresa_por_mes'])
                problemas.append(f'{nombre}: el Resumen no trae la columna de {mes_txt(periodo)} (llega hasta {mes_txt(cols[-1]) if cols else "—"}). '
                                 'Subí el Resumen actualizado con ese mes.')
                continue
            leidos[t] = (nombre, ruta, meta, r)
        except Exception as e:
            problemas.append(f'{nombre}: {dp.mensaje_usuario(e, NOMBRE[t])}')
    return periodo, (problemas if problemas else leidos)

def fecha_txt(iso):
    return f'{iso[8:10]}/{iso[5:7]}/{iso[:4]}' if iso else 'sin fecha'

def revisar_contra_base(base, leidos, periodo):
    """Compara con lo ya cargado. Devuelve (problemas, nombres de los archivos que cambiarían el mes)."""
    b = dp.Base(base)
    problemas, distintos = [], []
    for t, (nombre, _, meta, r) in leidos.items():
        actual = b.leer(t, periodo)
        if not actual or actual['_meta'].get('sha1') == meta['sha1']:
            continue
        distintos.append(nombre)
        am = actual['_meta']
        if t == 'RESUMEN':
            llega, cargado = max(r['empresa_por_mes']), am.get('cubre_hasta')
            if cargado and llega < cargado:
                problemas.append(f'{nombre}: llega hasta {mes_txt(llega)}, pero el Resumen ya cargado llega hasta {mes_txt(cargado)}. Subí el Resumen más nuevo.')
        else:
            nueva, cargada = r.get('emision') or meta.get('mtime'), am.get('emision') or am.get('mtime')
            if nueva and cargada and nueva < cargada:
                problemas.append(f'{nombre}: es del {fecha_txt(nueva)} y el ya cargado es más nuevo (del {fecha_txt(cargada)}). Subí la versión más nueva.')
    return problemas, distintos

ETIQUETAS = {'m2': 'm² pintados', 'ingresos': 'Ingresos', 'cv_pintura_fx': 'Pintura teórica', 'cv_quimicos_fx': 'Químicos teóricos',
             'cv_mod_fx': 'Mano de obra teórica', 'cf_fx': 'Costo fijo teórico', 'cv_pintura_real': 'Pintura real',
             'cv_quimicos_real': 'Químicos reales', 'cv_mod_real': 'Mano de obra real', 'cf_gastos_real': 'Gastos fijos reales',
             'sueldos': 'Sueldos', 'cargas_sociales': 'Cargas sociales', 'retiros_socios': 'Retiros de socios',
             'kg_pintura_real': 'Kg de pintura reales', 'kg_pintura_fx': 'Kg de pintura teóricos', 'lts_quimicos': 'Litros de químicos',
             'pintura_pagada_real': 'Pintura pagada', 'inversiones': 'Inversiones', 'egresos_a_confirmar': 'Egresos a confirmar'}
SIN_PESOS = {'m2', 'kg_pintura_real', 'kg_pintura_fx', 'lts_quimicos'}

def foto_salida(salida):
    """Los valores por mes que muestra hoy el tablero: sirven para comparar antes y después de una carga."""
    try:
        with open(os.path.join(salida, 'resumen_mensual.csv'), encoding='utf-8', newline='') as f:
            return {r['periodo']: r for r in csv.DictReader(f)}
    except (OSError, ValueError):
        return {}

def otros_meses_cambiados(antes, despues, periodo):
    """Meses distintos del cargado cuyos valores cambiaron, con el detalle en palabras."""
    def valor(campo, v):
        if v in ('', None):
            return 'sin dato'
        return dp.ars(float(v)) if campo not in SIN_PESOS else f'{float(v):,.0f}'.replace(',', '.')
    salida = []
    for p in sorted(set(antes) & set(despues)):
        if p == periodo:
            continue
        cambios = []
        for campo, etiqueta in ETIQUETAS.items():
            # Se compara lo que se muestra: el mismo Excel guardado como .xls o .xlsx puede diferir en un centavo por redondeo
            a, d = valor(campo, antes[p].get(campo)), valor(campo, despues[p].get(campo))
            if a != d:
                cambios.append(f'{etiqueta}: de {a} a {d}')
        if cambios:
            salida.append({'periodo': p, 'mes': mes_txt(p), 'cambios': cambios})
    return salida

def ubicar(entrada, leidos, periodo):
    """Mueve las versiones anteriores a _reemplazados/<fecha>/ y copia los nuevos a la entrada. Devuelve (nuevos, movidos)."""
    sello = dt.datetime.now().strftime('%Y-%m-%d_%H%M%S')
    destino_viejos = os.path.join(entrada, '_reemplazados', sello)
    movidos, nuevos = [], []
    existentes = [p for p in dp.listar_excel(entrada) if dp.identificar(p)]
    for t, (nombre, ruta, meta, r) in leidos.items():
        for p in existentes:
            idf = dp.identificar(p)
            mismo = (idf['tipo'] == t and t != 'RESUMEN' and idf['periodo'] == periodo) or \
                    (t == 'RESUMEN' and idf['tipo'] == 'RESUMEN' and nombre_base(p) == nombre_base(nombre)) or \
                    os.path.basename(p).lower() == os.path.basename(nombre).lower()
            if mismo and os.path.exists(p):
                os.makedirs(destino_viejos, exist_ok=True)
                dest = os.path.join(destino_viejos, os.path.relpath(p, entrada).replace(os.sep, '__'))
                shutil.move(p, dest)
                movidos.append((p, dest))
        final = os.path.join(entrada, os.path.basename(nombre))
        shutil.copy2(ruta, final)
        nuevos.append(final)
    return nuevos, movidos

def deshacer(nuevos, movidos):
    for p in nuevos:
        if os.path.exists(p):
            os.remove(p)
    for orig, dest in movidos:
        if os.path.exists(dest):
            shutil.move(dest, orig)

def correr_pipeline(entrada, base, salida):
    """Corre el pipeline esperando hasta ESPERA_BLOQUEO_SEG si hay otra carga en curso."""
    limite = time.time() + ESPERA_BLOQUEO_SEG
    while True:
        try:
            with dp.Bloqueo(base):
                b = dp.Base(base)
                dp.cargar_reglas(base)
                log, errores = dp.procesar_carpeta(entrada, b)
                dp.armar_salidas(b, salida, errores=dp.clasificar_errores(b, errores))
                with open(os.path.join(salida, 'calidad.json'), encoding='utf-8') as f:
                    return log, json.load(f)
        except dp.CargaEnCurso:
            if time.time() > limite:
                raise
            time.sleep(3)

def cargar_mes(archivos, entrada, base, salida, confirmar=False):
    """archivos: lista de (nombre_original, ruta). Devuelve el resultado como dict.
    confirmar: autoriza a reemplazar un mes ya cargado con archivos distintos."""
    os.makedirs(entrada, exist_ok=True)
    periodo, res = validar(archivos)
    if isinstance(res, list):
        return rechazo(res, periodo)
    problemas, distintos = revisar_contra_base(base, res, periodo)
    if problemas:
        return rechazo(problemas, periodo)
    if distintos and not confirmar:
        return {'ok': False, 'requiere_confirmacion': True, 'periodo': periodo, 'mes': mes_txt(periodo),
                'mensaje': f'{mes_txt(periodo).capitalize()} ya está cargado. ¿Querés reemplazarlo? Los archivos anteriores quedan guardados.',
                'distintos': sorted(distintos), 'problemas': [], 'avisos': []}
    antes = foto_salida(salida)
    nuevos, movidos = ubicar(entrada, res, periodo)
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            log, cal = correr_pipeline(entrada, base, salida)
    except dp.CargaEnCurso:
        deshacer(nuevos, movidos)
        return rechazo(['Hay otra carga en curso desde hace más de un minuto. Probá de nuevo en unos minutos.'], periodo)
    nombres = {os.path.basename(n) for n in nuevos}
    errs = [e for e in cal.get('errores_carga', []) if e.get('archivo') in nombres]
    q = cal.get('meses', {}).get(periodo)
    pendiente = next((x for x in cal.get('pendientes', []) if x['periodo'] == periodo), None)
    if errs or not q or pendiente:
        deshacer(nuevos, movidos)
        with contextlib.redirect_stdout(io.StringIO()):
            correr_pipeline(entrada, base, salida)     # deja la salida como estaba
        probs = [f'{e["archivo"]}: {e["error"]}' for e in errs] or ['El mes no quedó completo después de la carga. Avisale al equipo.']
        return rechazo(probs, periodo)
    acciones = sorted({l[4] for l in log if l[1] in nombres and l[3] == periodo})
    reemplazo = 'reemplazado' in acciones or (bool(movidos) and 'nuevo' not in acciones)
    fecha_otro_mes = [n for n in q.get('notas', []) if 'comprobantes con fecha de otro mes' in n]
    return {'ok': True, 'periodo': periodo, 'mes': mes_txt(periodo),
            'mensaje': (f'{mes_txt(periodo).capitalize()}: los 4 archivos son iguales a los que ya estaban cargados. No hubo cambios.'
                        if acciones == ['sin cambios'] else
                        f'{mes_txt(periodo).capitalize()} {"actualizado" if reemplazo else "cargado"} con los 4 archivos.'),
            'sin_cambios': acciones == ['sin cambios'],
            'estado': q.get('estado'), 'avisos': q.get('avisos', []) + fecha_otro_mes + [f'Falta: {f}' for f in q.get('faltantes', [])],
            'otros_meses': otros_meses_cambiados(antes, foto_salida(salida), periodo),
            'archivos': sorted(nuevos), 'reemplazados': [d for _, d in movidos], 'problemas': []}

def main(argv=None):
    ap = argparse.ArgumentParser(description='Carga los 4 Excel de un mes (ER, WEG, COMPRAS y Resumen) y corre el pipeline')
    ap.add_argument('--entrada', required=True); ap.add_argument('--base', required=True); ap.add_argument('--salida', required=True)
    ap.add_argument('--reemplazar', action='store_true', help='Autoriza a reemplazar un mes ya cargado')
    ap.add_argument('archivos', nargs='*')
    a = ap.parse_args(argv)
    res = cargar_mes([(os.path.basename(p), p) for p in a.archivos], a.entrada, a.base, a.salida, confirmar=a.reemplazar)
    print(json.dumps(res, ensure_ascii=False, indent=1))
    return 0 if res['ok'] else 1

if __name__ == '__main__':
    sys.exit(main())
