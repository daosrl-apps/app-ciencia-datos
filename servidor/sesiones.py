"""
Inicio de sesión del tablero DAO SRL (solo librería estándar).

- Usuarios y contraseñas: el mismo archivo htpasswd de siempre (contraseñas con `openssl passwd -apr1`).
  Se lee en cada intento, así que agregar, quitar o cambiar un usuario vale al instante, sin reiniciar.
- Sesión: una clave al azar que el tablero guarda solo en esa pestaña (sessionStorage). Vence a las
  SESION_HORAS sin uso. Si al usuario lo borran o le cambian la contraseña, sus sesiones dejan de valer.
  Las sesiones viven en memoria: si se reinicia el servicio, hay que volver a entrar.
- Bloqueo: INTENTOS_MAX fallos seguidos del mismo usuario desde la misma conexión lo bloquean BLOQUEO_MIN
  minutos; INTENTOS_MAX_IP fallos desde una conexión (con cualquier usuario) bloquean esa conexión.
"""
import hashlib, hmac, os, secrets, threading, time

SESION_HORAS = 8
INTENTOS_MAX = 5
INTENTOS_MAX_IP = 20
BLOQUEO_MIN = 15

_ITOA64 = './0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'
_lock = threading.Lock()
_sesiones = {}   # clave -> {'usuario', 'hash', 'uso'}
_fallos = {}     # ('u', usuario, ip) o ('ip', ip) -> {'n', 'hasta'}


def apr1(clave, sal):
    """Hash $apr1$ de Apache (el que genera `openssl passwd -apr1`)."""
    pw, sal = clave.encode('utf-8'), sal.encode('utf-8')[:8]
    ctx = pw + b'$apr1$' + sal
    alt = hashlib.md5(pw + sal + pw).digest()
    for i in range(len(pw), 0, -16):
        ctx += alt[:min(16, i)]
    i = len(pw)
    while i:
        ctx += b'\x00' if i & 1 else pw[:1]
        i >>= 1
    f = hashlib.md5(ctx).digest()
    for i in range(1000):
        c = (pw if i & 1 else f) + (sal if i % 3 else b'') + (pw if i % 7 else b'') + (f if i & 1 else pw)
        f = hashlib.md5(c).digest()

    def to64(v, n):
        s = ''
        for _ in range(n):
            s += _ITOA64[v & 0x3f]
            v >>= 6
        return s
    out = ''.join(to64((f[a] << 16) | (f[b] << 8) | f[c], 4)
                  for a, b, c in ((0, 6, 12), (1, 7, 13), (2, 8, 14), (3, 9, 15), (4, 10, 5)))
    return f'$apr1${sal.decode()}${out}{to64(f[11], 2)}'


def leer_usuarios(ruta):
    usuarios = {}
    try:
        with open(ruta, encoding='utf-8') as fh:
            for linea in fh:
                linea = linea.strip()
                if linea and not linea.startswith('#') and ':' in linea:
                    u, h = linea.split(':', 1)
                    usuarios[u.strip()] = h.strip()
    except FileNotFoundError:
        pass
    return usuarios


def _clave_ok(clave, h):
    if not h.startswith('$apr1$'):
        return False
    sal = h.split('$')[2]
    return hmac.compare_digest(apr1(clave, sal), h)


def _bloqueado(clave_f, ahora):
    f = _fallos.get(clave_f)
    return f['hasta'] - ahora if f and f['hasta'] > ahora else 0


def _fallo(clave_f, maximo, ahora):
    f = _fallos.setdefault(clave_f, {'n': 0, 'hasta': 0})
    if f['hasta'] and f['hasta'] <= ahora:   # ya cumplió un bloqueo: arranca de cero
        f.update(n=0, hasta=0)
    f['n'] += 1
    if f['n'] >= maximo:
        f['hasta'] = ahora + BLOQUEO_MIN * 60
    return f['hasta'] > ahora


def _limpiar(ahora):
    for k in [k for k, s in _sesiones.items() if ahora - s['uso'] > SESION_HORAS * 3600]:
        del _sesiones[k]
    for k in [k for k, f in _fallos.items() if f['hasta'] and f['hasta'] <= ahora - 3600]:
        del _fallos[k]


def _minutos(seg):
    m = max(1, int((seg + 59) // 60))
    return f'{m} minuto{"s" if m != 1 else ""}'


def iniciar(ruta_htpasswd, usuario, clave, ip):
    """Devuelve (codigo_http, respuesta)."""
    usuario, clave = (usuario or '').strip()[:100], (clave or '')[:200]
    ahora = time.time()
    ku, kip = ('u', usuario.lower(), ip), ('ip', ip)
    with _lock:
        _limpiar(ahora)
        espera = max(_bloqueado(ku, ahora), _bloqueado(kip, ahora))
        if espera:
            return 429, {'ok': False, 'mensaje': f'Demasiados intentos. Probá de nuevo en {_minutos(espera)}.'}
    h = leer_usuarios(ruta_htpasswd).get(usuario)
    if h and usuario and clave:
        ok = _clave_ok(clave, h)
    else:
        apr1(clave or 'x', 'xxxxxxxx')   # mismo tiempo de respuesta: no deja adivinar qué usuarios existen
        ok = False
    with _lock:
        if not ok:
            b1 = _fallo(ku, INTENTOS_MAX, ahora)
            b2 = _fallo(kip, INTENTOS_MAX_IP, ahora)
            if b1 or b2:
                return 429, {'ok': False, 'mensaje': f'Demasiados intentos. Probá de nuevo en {BLOQUEO_MIN} minutos.'}
            return 401, {'ok': False, 'mensaje': 'Usuario o contraseña incorrectos.'}
        _fallos.pop(ku, None)
        token = secrets.token_urlsafe(32)
        _sesiones[token] = {'usuario': usuario, 'hash': h, 'uso': ahora}
    return 200, {'ok': True, 'token': token, 'usuario': usuario, 'horas': SESION_HORAS}


def token_de(cabecera):
    cabecera = cabecera or ''
    return cabecera[7:].strip() if cabecera[:7].lower() == 'bearer ' else ''


def verificar(ruta_htpasswd, token):
    """Usuario de la sesión, o None si no vale. Cada uso la extiende SESION_HORAS más."""
    if not token:
        return None
    ahora = time.time()
    with _lock:
        s = _sesiones.get(token)
        if not s:
            return None
        if ahora - s['uso'] > SESION_HORAS * 3600:
            del _sesiones[token]
            return None
    if leer_usuarios(ruta_htpasswd).get(s['usuario']) != s['hash']:   # lo borraron o le cambiaron la contraseña
        with _lock:
            _sesiones.pop(token, None)
        return None
    with _lock:
        s['uso'] = ahora
    return s['usuario']


def cerrar(token):
    with _lock:
        _sesiones.pop(token, None)
