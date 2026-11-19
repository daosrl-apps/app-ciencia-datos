#!/usr/bin/env python3
"""
Servicio de carga — DAO SRL (v1.3)

Pequeño servidor HTTP (solo librería estándar) que corre junto al tablero en Contabo:
  POST /login     ← pantalla de inicio de sesión del tablero: usuario y contraseña (del htpasswd) → clave de sesión.
  POST /logout    ← botón "Cerrar sesión".
  GET  /sesion    ← el tablero pregunta si su sesión sigue valiendo (y la extiende).
  GET  /verificar ← Nginx pregunta lo mismo antes de entregar data/ (auth_request).
  POST /subir     ← el botón "Subir Excel del mes" del tablero (vía Nginx; pide sesión iniciada).
                    Recibe los 4 Excel del mes, corre cargar_mes.py y responde el resultado en JSON.
                    Si el mes ya estaba cargado con otros archivos responde 409 y pide confirmación
                    (el tablero reenvía con el campo confirmar_reemplazo=si).
                    Si el mes se cargó, le avisa a n8n para que guarde una copia en el Drive.
  POST /procesar  ← n8n, cuando baja un archivo nuevo del Drive o una vez por día. Corre el pipeline completo.
  GET  /salud     ← control de que el servicio está vivo.

No queda expuesto a internet: Nginx reenvía /subir y n8n lo llama por la red interna de Docker.

Variables de entorno:
  DAO_ENTRADA   carpeta de los Excel (la misma que sincroniza n8n con el Drive)   [/datos/excel]
  DAO_BASE      base histórica                                                    [/datos/base]
  DAO_SALIDA    carpeta data/ que lee el tablero                                 [/datos/tablero/data]
  N8N_COPIA_URL webhook de n8n que copia los archivos al Drive (opcional)         [vacío = no copia]
  DAO_HTPASSWD  archivo de usuarios y contraseñas (htpasswd, $apr1$)         [vacío = sin inicio de sesión, solo para pruebas]
  PUERTO        [8000]
"""
import email.parser, email.policy, json, os, re, shutil, sys, tempfile, threading, urllib.request, datetime as dt
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

AQUI = os.path.dirname(os.path.abspath(__file__))
for cand in (os.path.join(AQUI, '..', 'pipeline'), os.path.join(AQUI, 'pipeline'), '/app/pipeline'):
    if os.path.exists(os.path.join(cand, 'cargar_mes.py')):
        sys.path.insert(0, os.path.abspath(cand)); break
import cargar_mes   # noqa: E402
sys.path.insert(0, AQUI)
import sesiones   # noqa: E402
import dao_pipeline as dp   # noqa: E402

ENTRADA = os.environ.get('DAO_ENTRADA', '/datos/excel')
BASE = os.environ.get('DAO_BASE', '/datos/base')
SALIDA = os.environ.get('DAO_SALIDA', '/datos/tablero/data')
N8N_COPIA_URL = os.environ.get('N8N_COPIA_URL', '')
HTPASSWD = os.environ.get('DAO_HTPASSWD', '')
MAX_CUERPO = (cargar_mes.MAX_MB + 5) * 1024 * 1024
UNA_A_LA_VEZ = threading.Lock()

def log(msg):
    print(f'{dt.datetime.now().isoformat(timespec="seconds")} {msg}', flush=True)

def nombre_seguro(n):
    n = os.path.basename((n or '').replace('\\', '/')).strip()
    n = re.sub(r'[^\w\s().\-áéíóúÁÉÍÓÚñÑ]', '_', n)
    return n[:120] or 'archivo'

def avisar_n8n(resultado):
    """Le pide a n8n que copie al Drive los archivos recién cargados. Si falla, no afecta la carga."""
    if not N8N_COPIA_URL:
        return
    cuerpo = json.dumps({'periodo': resultado['periodo'],
                         'archivos': [{'nombre': os.path.basename(p), 'ruta': p} for p in resultado.get('archivos', [])]}).encode()
    try:
        req = urllib.request.Request(N8N_COPIA_URL, data=cuerpo, headers={'Content-Type': 'application/json'}, method='POST')
        urllib.request.urlopen(req, timeout=30).read()
        log(f'copia al Drive pedida a n8n ({resultado["periodo"]})')
    except Exception as e:
        log(f'AVISO: no se pudo pedir la copia al Drive a n8n: {e}')

class Manejador(BaseHTTPRequestHandler):
    server_version = 'DAOCargador/1.2'

    def _json(self, codigo, datos):
        cuerpo = json.dumps(datos, ensure_ascii=False).encode('utf-8')
        self.send_response(codigo)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Cache-Control', 'no-store')
        self.send_header('Content-Length', str(len(cuerpo)))
        self.end_headers()
        self.wfile.write(cuerpo)

    def log_message(self, fmt, *args):
        log(f'{self.address_string()} {fmt % args}')

    # ── Sesión ──
    def _ip(self):
        return (self.headers.get('X-Real-IP') or self.client_address[0]).strip()

    def _token(self):
        return sesiones.token_de(self.headers.get('Authorization'))

    def _usuario(self):
        """Usuario con sesión válida; si el inicio de sesión está apagado (pruebas), '-'."""
        if not HTPASSWD:
            return '-'
        return sesiones.verificar(HTPASSWD, self._token())

    def _sin_sesion(self):
        return self._json(401, cargar_mes.rechazo(['Tu sesión venció. Iniciá sesión de nuevo.']))

    def do_GET(self):
        ruta = self.path.split('?')[0].rstrip('/')
        if ruta in ('/salud', ''):
            return self._json(200, {'ok': True, 'version': dp.VERSION})
        if ruta.endswith('/verificar'):        # para Nginx: solo el código importa
            ok = self._usuario()
            self.send_response(204 if ok else 401)
            self.send_header('Content-Length', '0')
            self.end_headers()
            return
        if ruta.endswith('/sesion'):
            if not HTPASSWD:
                return self._json(200, {'ok': True, 'auth': False})
            u = self._usuario()
            return self._json(200, {'ok': True, 'auth': True, 'usuario': u, 'horas': sesiones.SESION_HORAS}) if u else self._sin_sesion()
        self._json(404, {'ok': False, 'mensaje': 'No existe.'})

    def _login(self, cuerpo):
        if not HTPASSWD:
            return self._json(200, {'ok': True, 'auth': False})
        try:
            datos = json.loads(cuerpo.decode('utf-8') or '{}')
        except Exception:
            datos = {}
        if not isinstance(datos, dict):
            datos = {}
        codigo, res = sesiones.iniciar(HTPASSWD, str(datos.get('usuario') or ''), str(datos.get('clave') or ''), self._ip())
        log(f'inicio de sesión: {str(datos.get("usuario") or "")[:40]!r} desde {self._ip()} → '
            f'{"OK" if codigo == 200 else "BLOQUEADO" if codigo == 429 else "rechazado"}')
        self._json(codigo, res)

    def do_POST(self):
        ruta = self.path.split('?')[0].rstrip('/')
        largo = int(self.headers.get('Content-Length') or 0)
        if largo > MAX_CUERPO:
            return self._json(413, cargar_mes.rechazo([f'Los archivos superan los {cargar_mes.MAX_MB} MB permitidos.']))
        cuerpo = self.rfile.read(largo) if largo else b''
        if ruta.endswith('/login'):
            return self._login(cuerpo)
        if ruta.endswith('/logout'):
            sesiones.cerrar(self._token())
            return self._json(200, {'ok': True})
        if ruta.endswith('/subir'):
            usuario = self._usuario()
            if not usuario:
                return self._sin_sesion()
            log(f'subida de {usuario}')
            return self._subir(cuerpo)
        if ruta.endswith('/procesar'):
            return self._procesar()
        self._json(404, {'ok': False, 'mensaje': 'No existe.'})

    def _subir(self, cuerpo):
        ctype = self.headers.get('Content-Type', '')
        if 'multipart/form-data' not in ctype:
            return self._json(400, cargar_mes.rechazo(['El envío no trae archivos (se esperaba multipart/form-data).']))
        msg = email.parser.BytesParser(policy=email.policy.default).parsebytes(
            b'Content-Type: ' + ctype.encode('latin-1') + b'\r\nMIME-Version: 1.0\r\n\r\n' + cuerpo)
        tmp = tempfile.mkdtemp(prefix='dao_subida_')
        try:
            archivos, confirmar = [], False
            for parte in msg.iter_parts():
                nombre = parte.get_filename()
                if not nombre:
                    confirmar = confirmar or (parte.get_param('name', header='content-disposition') == 'confirmar_reemplazo'
                                              and parte.get_content().strip() == 'si')
                    continue
                nombre = nombre_seguro(nombre)
                destino = os.path.join(tmp, f'{len(archivos)}_{nombre}')
                with open(destino, 'wb') as f:
                    f.write(parte.get_payload(decode=True) or b'')
                archivos.append((nombre, destino))
            log(f'subida: {[n for n, _ in archivos]}')
            with UNA_A_LA_VEZ:
                res = cargar_mes.cargar_mes(archivos, ENTRADA, BASE, SALIDA, confirmar=confirmar)
            estado = 'OK' if res['ok'] else 'PIDE CONFIRMACIÓN' if res.get('requiere_confirmacion') else 'RECHAZADO'
            log(f'resultado: {estado} {res.get("periodo")} {res.get("problemas")}')
            if res['ok'] and res.get('sin_cambios'):
                res['copia_drive'] = 'no hacía falta (sin cambios)'
            elif res['ok']:
                res['copia_drive'] = 'pedida' if N8N_COPIA_URL else 'no configurada'
                threading.Thread(target=avisar_n8n, args=(res,), daemon=True).start()
            self._json(200 if res['ok'] else 409 if res.get('requiere_confirmacion') else 422, res)
        except Exception as e:
            log(f'ERROR en /subir: {type(e).__name__}: {e}')
            self._json(500, cargar_mes.rechazo(['Error inesperado en el servidor. Avisale al equipo: el detalle quedó en el registro del servicio.']))
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def _procesar(self):
        try:
            with UNA_A_LA_VEZ:
                log_, cal = cargar_mes.correr_pipeline(ENTRADA, BASE, SALIDA)
            nuevos = [e for e in cal.get('errores_carga', []) if e.get('nuevo')]
            cambios = [l for l in log_ if l[4] in ('nuevo', 'reemplazado')]
            self._json(200, {'ok': not nuevos, 'cambios': len(cambios), 'errores_nuevos': nuevos,
                             'errores': cal.get('errores_carga', []), 'pendientes': cal.get('pendientes', [])})
        except dp.CargaEnCurso as e:
            self._json(200, {'ok': True, 'cambios': 0, 'mensaje': str(e)})
        except Exception as e:
            log(f'ERROR en /procesar: {type(e).__name__}: {e}')
            self._json(500, {'ok': False, 'mensaje': f'{type(e).__name__}: {e}'})

def main():
    for d in (ENTRADA, BASE, SALIDA):
        os.makedirs(d, exist_ok=True)
    puerto = int(os.environ.get('PUERTO', '8000'))
    log(f'cargador DAO {dp.VERSION} en el puerto {puerto} · entrada={ENTRADA} base={BASE} salida={SALIDA} · '
        f'inicio de sesión: {HTPASSWD or "APAGADO (solo para pruebas)"}')
    ThreadingHTTPServer(('0.0.0.0', puerto), Manejador).serve_forever()

if __name__ == '__main__':
    main()
