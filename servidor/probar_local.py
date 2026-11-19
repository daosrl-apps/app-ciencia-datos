#!/usr/bin/env python3
"""
Prueba local del tablero publicado, sin Docker (hace de Nginx + servicio de carga en una sola compu).

  python servidor/probar_local.py --excel <carpeta_excel> --base <carpeta_base> --tablero tablero [--puerto 8080]

Abrí http://localhost:8080/tablero_dao.html: el botón "Subir Excel del mes" funciona igual que en el servidor.
Los datos se escriben en <tablero>/data/. No tiene usuario y contraseña: es solo para probar en la propia compu.
"""
import argparse, functools, os, sys
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--excel', required=True); ap.add_argument('--base', required=True)
    ap.add_argument('--tablero', required=True); ap.add_argument('--puerto', type=int, default=8080)
    a = ap.parse_args()
    os.environ['DAO_ENTRADA'] = os.path.abspath(a.excel)
    os.environ['DAO_BASE'] = os.path.abspath(a.base)
    os.environ['DAO_SALIDA'] = os.path.join(os.path.abspath(a.tablero), 'data')
    os.environ.setdefault('N8N_COPIA_URL', '')
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import cargador

    class Local(cargador.Manejador, SimpleHTTPRequestHandler):
        def do_GET(self):
            if self.path.split('?')[0].rstrip('/') == '/salud':
                return cargador.Manejador.do_GET(self)
            return SimpleHTTPRequestHandler.do_GET(self)
        def end_headers(self):
            self.send_header('Cache-Control', 'no-store')
            super().end_headers()

    for d in (os.environ['DAO_ENTRADA'], os.environ['DAO_BASE'], os.environ['DAO_SALIDA']):
        os.makedirs(d, exist_ok=True)
    srv = ThreadingHTTPServer(('127.0.0.1', a.puerto), functools.partial(Local, directory=os.path.abspath(a.tablero)))
    print(f'Tablero en http://localhost:{a.puerto}/tablero_dao.html  (Ctrl+C para salir)', flush=True)
    srv.serve_forever()

if __name__ == '__main__':
    main()
