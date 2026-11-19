#!/usr/bin/env python3
"""
Arma tablero_dao.html con los datos adentro (un solo archivo que se abre con doble clic o se manda por mail).

Uso: python tablero/build_tablero.py --data tablero/data --salida tablero/tablero_dao.html
Lee plantilla.html (junto a este script) y reemplaza los marcadores por los CSV y el JSON de calidad.
Sin este paso, plantilla.html igual funciona si se sirve desde un servidor web con la carpeta data/ al lado.
"""
import argparse, os

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--data', required=True, help='Carpeta con detalle_clientes_articulos.csv, resumen_mensual.csv y calidad.json')
    ap.add_argument('--salida', required=True, help='Ruta del HTML a generar')
    ap.add_argument('--plantilla', default='plantilla.html', help='Plantilla a usar, junto a este script (por defecto plantilla.html)')
    a = ap.parse_args()
    aqui = os.path.dirname(os.path.abspath(__file__))
    html = open(os.path.join(aqui, a.plantilla), encoding='utf-8').read()
    leer = lambda n: open(os.path.join(a.data, n), encoding='utf-8').read().strip() if os.path.exists(os.path.join(a.data, n)) else ''
    det, res, cal = leer('detalle_clientes_articulos.csv'), leer('resumen_mensual.csv'), leer('calidad.json')
    if '</script' in det + res + cal:
        raise SystemExit('Los datos contienen "</script": no se pueden incrustar.')
    for marca, valor in (('/*__DETALLE__*/', det), ('/*__RESUMEN__*/', res), ('/*__CALIDAD__*/', cal)):
        if marca not in html:
            raise SystemExit(f'No encuentro el marcador {marca} en plantilla.html')
        html = html.replace(marca, valor)
    with open(a.salida, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'Generado {a.salida} ({len(html) / 1024:.0f} KB)')

if __name__ == '__main__':
    main()
