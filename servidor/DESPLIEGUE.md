# Despliegue del tablero DAO SRL en Contabo

El tablero (versión 2 del diseño) queda publicado en **https://cienciadedatos.daosrl.com.ar**, con la carga de meses desde el propio tablero:

- Luciano entra al tablero y usa **"Subir Excel del mes"** con los 4 archivos. El servidor los valida, corre el pipeline y el tablero se actualiza en el momento.
- **Regla:** un mes se carga solo cuando están sus **4 archivos** (ER, WEG, COMPRAS y un Resumen con la columna de ese mes).
- **Mes ya cargado:** si los archivos son distintos, el tablero pide confirmación antes de reemplazar y después lista los otros meses que cambiaron.
- Por pedido de Luciano, los meses se cargan **solo con el botón**, subiendo los 4 Excel a mano. No hay carga automática desde Google Drive (la sección 11 queda como referencia, apagada).

## Cómo está armado en el servidor de Luciano

El servidor ya tiene otras aplicaciones de Luciano. Su aplicación **app-dao-ops** tiene un Nginx que atiende los puertos 80 y 443 de todos sus sitios (planta., 3dceci., metricas., comidas.daosrl.com.ar), con sus certificados de Let's Encrypt. El tablero se suma como **un sitio más** de ese esquema, sin tocar lo suyo:

```
 Navegador ──HTTPS──▶ Nginx de app-dao-ops (Luciano: puertos 80/443, certificado)
                         └─ cienciadedatos.daosrl.com.ar ──▶ 172.20.0.1:8088
                                                              Nginx del tablero
                                                                ├─▶ /tablero_dao.html   (pantalla de inicio de sesión y tablero, sin datos adentro)
                                                                ├─▶ /login, /logout, /sesion ──▶ cargador (valida contra servidor/htpasswd)
                                                                ├─▶ /data/  (solo con sesión: Nginx se lo pregunta al cargador)
                                                                └─▶ /subir  (solo con sesión) ──▶ cargador (pipeline) ──▶ /opt/dao/datos
```

**Inicio de sesión (desde el 12/11/2026):** el tablero tiene su propia pantalla de inicio (ya no la ventanita gris del navegador). La sesión vive solo en la pestaña: al cerrarla, la próxima vez pide la contraseña; recargar (F5) no. Vence a las 8 horas sin uso y se puede cerrar con el botón **Salir**. Cinco contraseñas equivocadas seguidas bloquean ese usuario 15 minutos desde esa conexión. Las sesiones viven en la memoria del servicio de carga: si se reinicia, todos vuelven a entrar.

| Qué | Dónde | De quién |
|---|---|---|
| Código y datos del tablero | `/opt/dao` (código de root, `datos/` del usuario 1000 = `ubuntu`) | Nuestro |
| Contenedores `dao-tablero-web-1` y `dao-tablero-cargador-1` | `/opt/dao/servidor/docker-compose.yml` (proyecto `dao-tablero`) | Nuestro |
| Puerto interno `172.20.0.1:8088` | Solo dentro del servidor: desde internet no se ve | Nuestro |
| Certificado `cienciadedatos.daosrl.com.ar` | `/etc/letsencrypt` (certbot *webroot*, igual que los de Luciano; se renueva solo y su script `renewal-hooks/deploy/reload-nginx.sh` recarga su Nginx) | Compartido |
| Bloque `server` para el dominio | `nginx/app-dao-ops.conf` del **repositorio de app-dao-ops** | **De Luciano** (sección 5.4) |
| Backup diario | `crontab` de root → `/backups/dao-AAAA-MM-DD.tgz` | Nuestro |

> **Instalado el 11/11/2026** (Ubuntu 24.04, Docker 29.1, Compose 2.40, certbot ya instalado por Luciano). Antes se probó en Docker en una Mac, incluso con un Nginx igual al de Luciano adelante.

---

## 1. Requerimientos

| Qué | Estado en el servidor de Luciano |
|---|---|
| Acceso SSH como `root` con clave | Hecho (clave `~/.ssh/dao_contabo` en la Mac de Bauti) |
| Docker + Docker Compose | Ya instalados |
| certbot | Ya instalado (lo usa Luciano para sus sitios) |
| Dominio `cienciadedatos.daosrl.com.ar` → `169.58.110.177` | Hecho (verificado el 09/11/2026) |
| Puerto interno libre (8088) | Libre el 11/11/2026. Si alguna vez lo usa otra aplicación, cambiarlo en `docker-compose.yml` (`PUERTO_TABLERO`) y en el bloque de Luciano |

> Con HTTPS, la contraseña y los Excel que se suben (sueldos, ingresos por cliente, retiros de socios) viajan cifrados. Igual, usá para el tablero contraseñas que no se usen en ningún otro lado.

---

## 2. Copiar el proyecto al servidor

Desde la Mac, en la carpeta `DAO/` (no copia `tablero/data` ni los tableros con datos adentro: el servidor arma los suyos):

```bash
SSH="ssh -i ~/.ssh/dao_contabo root@169.58.110.177"
COPYFILE_DISABLE=1 tar czf - --exclude=.DS_Store --exclude=__pycache__ pipeline servidor \
  tablero/plantilla_v2.html tablero/plantilla.html tablero/build_tablero.py | $SSH 'mkdir -p /opt/dao && tar xzf - -C /opt/dao'
```

Tiene que quedar así:

```
/opt/dao/
├── pipeline/   (dao_pipeline.py, cargar_mes.py, config/)
├── tablero/    (plantilla_v2.html ← la que se publica, plantilla.html, build_tablero.py)
├── servidor/   (docker-compose.yml, nginx.conf, Dockerfile.cargador, cargador.py, sesiones.py, htpasswd, esta guía)
└── datos/      (paso 3; NO va al repositorio)
```

## 3. Carpetas de datos y Excel históricos

Copiá **la carpeta `excel/` completa**: las subcarpetas `2025/` y `2026/` **y también los archivos sueltos** (hoy, los 4 de septiembre 2026 y el `Resumen_InterAnual_2026`). Sin el Resumen, todos los meses de 2026 quedan "pendientes".

```bash
# desde la Mac, en la carpeta DAO/
COPYFILE_DISABLE=1 tar czf - --exclude=.DS_Store -C excel . | $SSH 'mkdir -p /opt/dao/datos/excel /opt/dao/datos/base /opt/dao/datos/tablero/data && tar xzf - -C /opt/dao/datos/excel'
# en el servidor
cd /opt/dao
cp pipeline/config/reglas_compras.csv pipeline/config/ajustes_comprobantes.csv datos/base/
chmod -R a+rX pipeline servidor tablero     # los contenedores tienen que poder leer el código
chown -R 1000:1000 datos                    # el servicio de carga corre con el usuario 1000
: > servidor/htpasswd && chmod 644 servidor/htpasswd   # sin usuarios todavía: nadie puede entrar
```

(Los avisos de `tar` sobre `LIBARCHIVE.xattr.com.apple…` son metadatos de la Mac y no afectan.)

## 4. Usuarios del tablero

Un usuario por persona (Luciano y cada integrante del equipo). Todos pueden ver el tablero y subir archivos. Cada uno se agrega **desde la Terminal de la Mac**: el comando pide la contraseña dos veces y no queda escrita en ningún lado (en el servidor se guarda solo cifrada). Reemplazá `NOMBRE`:

```bash
ssh -t -i ~/.ssh/dao_contabo root@169.58.110.177 'printf "NOMBRE:%s\n" "$(openssl passwd -apr1)" >> /opt/dao/servidor/htpasswd'
```

- El servicio de carga lee el archivo en cada inicio de sesión: no hace falta reiniciar nada. Si a alguien lo sacan o le cambian la contraseña, su sesión abierta deja de valer al instante.
- Las contraseñas tienen que estar cifradas con `openssl passwd -apr1` (las que empiezan con `$apr1$`), como hace el comando de arriba.
- **Ver los usuarios:** `$SSH 'cut -d: -f1 /opt/dao/servidor/htpasswd'`.
- **Sacar un usuario:** `$SSH "sed -i '/^NOMBRE:/d' /opt/dao/servidor/htpasswd"`.
- **Cambiar una contraseña:** sacarlo y volver a agregarlo.

## 5. Levantar, carga inicial y publicar

### 5.1 Levantar el tablero

```bash
cd /opt/dao/servidor
docker compose up -d --build
docker compose ps    # dao-tablero-web-1 en 172.20.0.1:8088 y dao-tablero-cargador-1 "healthy"
```

### 5.2 Carga inicial

```bash
docker compose exec cargador python -c "import urllib.request as u; print(u.urlopen(u.Request('http://localhost:8000/procesar', method='POST'), timeout=600).read().decode())"
# → {"ok": true, "cambios": 88, "errores_nuevos": [], "pendientes": [] ...}   (11/11/2026: 88 cambios, 0 pendientes)
```

### 5.3 Certificado (sin cortar los sitios de Luciano)

Se saca como los de Luciano: su Nginx ya responde en el puerto 80 la validación de Let's Encrypt desde `/var/www/certbot`.

```bash
certbot certonly --webroot -w /var/www/certbot -d cienciadedatos.daosrl.com.ar --non-interactive --agree-tos --dry-run   # ensayo
certbot certonly --webroot -w /var/www/certbot -d cienciadedatos.daosrl.com.ar --non-interactive --agree-tos
certbot certificates    # cienciadedatos.daosrl.com.ar, vence a los 90 días y se renueva solo
```

### 5.4 Bloque para el Nginx de Luciano (lo agrega él, en su repositorio de app-dao-ops)

En `nginx/app-dao-ops.conf`, junto a los bloques de sus otros sitios. Igual al de `comidas`, con dos diferencias necesarias para el botón de carga: acepta envíos de hasta 25 MB (los 4 Excel) y espera hasta 3 minutos (la carga corre el pipeline).

```nginx
# Tablero DAO SRL (equipo de Laboratorio II): https://cienciadedatos.daosrl.com.ar → tablero en 172.20.0.1:8088
server {
    listen 443 ssl;
    listen [::]:443 ssl;
    server_name cienciadedatos.daosrl.com.ar;

    ssl_certificate     /etc/letsencrypt/live/cienciadedatos.daosrl.com.ar/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/cienciadedatos.daosrl.com.ar/privkey.pem;
    ssl_protocols TLSv1.2 TLSv1.3;
    ssl_prefer_server_ciphers on;

    add_header Strict-Transport-Security "max-age=31536000; includeSubDomains" always;
    add_header X-Content-Type-Options "nosniff" always;
    add_header X-Frame-Options "SAMEORIGIN" always;

    client_max_body_size 25m;           # "Subir Excel del mes": los 4 Excel juntos

    location / {
        proxy_pass http://172.20.0.1:8088;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_read_timeout 180s;        # la carga corre el pipeline: puede tardar más de 60 s
        proxy_send_timeout 180s;
    }
}
```

Además, agregar `cienciadedatos.daosrl.com.ar` a la lista `server_name` del bloque del puerto 80, para que también redirija a HTTPS. Después de su despliegue, `docker exec app-dao-ops-nginx-1 nginx -t` tiene que decir "syntax is ok".

### 5.5 Ver el tablero antes de que esté el bloque de Luciano

Con un túnel SSH desde la Mac (deja la Terminal abierta mientras lo usás):

```bash
ssh -N -L 8088:172.20.0.1:8088 -i ~/.ssh/dao_contabo root@169.58.110.177
# y en el navegador: http://localhost:8088  (usuario y contraseña del paso 4, en la pantalla de inicio)
```

## 6. Backup diario

Sin Drive, los Excel y la historia de los meses quedan **solo en Contabo**. Sin backup, si se pierde el servidor se pierde todo.

```bash
sudo mkdir -p /backups
sudo crontab -e
```

Agregá esta línea. Todos los días a las 3:30 guarda una copia y borra las de más de 30 días:

```
30 3 * * * tar czf /backups/dao-$(date +\%F).tgz -C /opt/dao datos servidor/htpasswd && find /backups -name 'dao-*.tgz' -mtime +30 -delete
```

Esa copia queda en el mismo disco: si falla el disco, se pierde con todo lo demás. **Una vez por mes, después de cargar el mes, bajá la última copia** a un lugar seguro:

```bash
scp usuario@169.58.110.177:/backups/dao-AAAA-MM-DD.tgz .
```

Si `/backups` solo lo puede leer `root`, copiá antes el archivo a tu usuario con `sudo cp`.

## 7. Prueba de punta a punta (el día de la instalación)

| # | Prueba | Resultado esperado |
|---|---|---|
| 1 | Abrir `http://cienciadedatos.daosrl.com.ar` (sin la "s") | Pasa solo a `https://`, con candado, y muestra la pantalla de inicio de sesión |
| 2 | Iniciar sesión | Se ven todos los meses y, abajo de todo, "Datos del servidor". F5 no pide la contraseña; cerrar la pestaña y abrir otra, sí; "Salir" vuelve a la pantalla de inicio |
| 3 | "Subir Excel del mes" con solo 3 archivos | No deja cargar: "Falta: …" |
| 4 | Subir los 4 del último mes, iguales a los ya cargados | "No hubo cambios", sin pedir confirmación |
| 5 | Subir el último mes con algún archivo más nuevo o corregido, y en la confirmación elegir **Cerrar** | Avisa "… ya está cargado. ¿Querés reemplazarlo?" y, al cerrar, no cambia nada |
| 6 | Repetir la 5 y elegir **Sí, reemplazar** | "… actualizado con los 4 archivos" y, si cambiaron, la lista de otros meses. Los anteriores quedan en `datos/excel/_reemplazados/` |
| 7 | Subir copias del ER, el WEG y el COMPRAS renombradas a otro mes, junto con el Resumen | Se rechaza y explica que las fechas son de otro mes. No se carga nada |
| 8 | `docker compose restart`, y después correr a mano el comando de backup de la sección 6 | El tablero sigue igual (los datos están en `datos/`) y aparece el archivo en `/backups` |

## 8. HTTPS: cómo está armado

- El HTTPS lo pone el **Nginx de Luciano** (app-dao-ops) con el certificado de `cienciadedatos.daosrl.com.ar`, y reenvía al tablero en `172.20.0.1:8088`. El Nginx del tablero solo atiende por dentro del servidor y no entrega datos sin sesión iniciada.
- El puerto 80 lo atiende también el Nginx de Luciano: valida los certificados (`/var/www/certbot`) y redirige a HTTPS.
- Renovación: `systemctl list-timers | grep certbot` muestra cuándo corre; al renovar, su script `reload-nginx.sh` recarga su Nginx. Se prueba con `certbot renew --dry-run --cert-name cienciadedatos.daosrl.com.ar`.
- Si Luciano cambia su Nginx o su repositorio, el bloque de la sección 5.4 tiene que seguir estando.

## 9. Operación y mantenimiento

- **Registros:**
  - `docker compose logs -f cargador`: cada subida y su resultado;
  - `datos/base/registro_cargas.csv`: cada archivo leído.
- **Versiones anteriores de los Excel:** quedan en `datos/excel/_reemplazados/<fecha>/` y no se borran solas.
- **Actualizar el código:** copiar los archivos nuevos y correr `docker compose up -d --build`. Los datos no se tocan.
- **Agregar o cambiar contraseñas:** paso 4. No hace falta reiniciar nada.
- **Ver los inicios de sesión:** `docker compose logs cargador | grep "inicio de sesión"` (usuario, IP y si entró, fue rechazado o quedó bloqueado).
- **Copia para abrir con doble clic:** el botón no la actualiza, porque lleva los datos adentro. Cuando haga falta una al día, se regenera con `python tablero/build_tablero.py --data <carpeta data> --plantilla plantilla_v2.html --salida tablero_dao_v2.html`.
- **Cambios en el diseño del tablero:** se editan en `tablero/plantilla_v2.html`; se copia ese archivo al servidor y alcanza con recargar la página (Nginx lo lee directo, no hace falta reiniciar).

## 10. Problemas comunes

| Síntoma | Causa probable | Qué hacer |
|---|---|---|
| "Subir Excel" falla al guardar por permisos | Permisos de `datos/` | `sudo chown -R 1000:1000 /opt/dao/datos` |
| Vuelve a la pantalla de inicio con "Tu sesión venció" | Pasaron 8 horas sin uso, se reinició el servicio de carga, o se sacó el usuario o se le cambió la contraseña | Iniciar sesión de nuevo |
| "Demasiados intentos. Probá de nuevo en 15 minutos" | 5 contraseñas equivocadas seguidas | Esperar 15 minutos; si urge, `docker compose restart cargador` lo libera (y cierra todas las sesiones) |
| "Usuario o contraseña incorrectos" con la contraseña correcta | El usuario se escribe distinto (mayúsculas cuentan) o la contraseña no se guardó con `-apr1` | Ver los usuarios (paso 4) y volver a cargarlo |
| El tablero dice "Sin conexión a los datos" | El contenedor `web` no ve `datos/tablero/data` o la carpeta está vacía | `docker compose ps` y correr la carga inicial (paso 5) |
| Rechaza un archivo por "fecha de otro mes" | El nombre no coincide con el mes de sus comprobantes | Corregir el nombre o subir el archivo del mes correcto |
| Rechaza un archivo porque "el ya cargado es más nuevo" | Se subió una versión vieja | Subir la versión más nueva |
| El navegador avisa "conexión no segura" | El certificado no se emitió o venció | `certbot certificates`; si falta o venció, sección 5.3 |
| El dominio muestra otro sitio de Luciano o un error de su Nginx | Falta el bloque de la sección 5.4 en el Nginx de app-dao-ops | Pedirle a Luciano que lo agregue (o que revise su último despliegue) |
| "Subir Excel" falla con "413" o se corta al minuto | El bloque de Luciano no tiene `client_max_body_size 25m` o `proxy_read_timeout 180s` | Revisar el bloque (sección 5.4) |
| Imágenes de Docker Hub con "429 Too Many Requests" | Límite de descargas de Docker Hub | Esperar, o hacer `docker login` con una cuenta gratuita |

## 11. Sumar el Drive más adelante (opcional, hoy descartado)

Luciano pidió cargar solo con el botón, así que esta sección no se usa. Queda como referencia. Los flujos ya están en `servidor/n8n/`, pero nunca se probaron con el Drive real.

1. En `docker-compose.yml`, descomentar el bloque de n8n y poner en `N8N_COPIA_URL` el valor `http://n8n:5678/webhook/dao-copiar-drive`.
2. Crear la clave con la que n8n cifra las credenciales de Google y guardarla: sin ella se pierden las credenciales.
   ```bash
   cd /opt/dao/servidor
   echo "N8N_ENCRYPTION_KEY=$(openssl rand -hex 24)" > .env
   chmod 600 .env
   docker compose up -d --build
   ```
3. Entrar al editor de n8n, que no queda publicado en internet, con un túnel SSH desde tu compu:
   ```bash
   ssh -L 5678:localhost:5678 usuario@IP-DE-CONTABO
   # y en el navegador: http://localhost:5678  → crear la cuenta de administrador de n8n
   ```
4. **Credencial de Google Drive** (*Credentials → New → Google Drive*), con una de estas opciones:
   - **OAuth2**, con la cuenta de Luciano. Al crear el cliente OAuth en Google Cloud, el *redirect URL* es `http://localhost:5678/rest/oauth2-credential/callback`, que funciona por el túnel.
   - **Cuenta de servicio** (más estable en un servidor): crearla en Google Cloud, bajar el JSON y compartir la carpeta del Drive con el mail de la cuenta de servicio.
5. **Importar los flujos** (*Workflows → Import from file*): `servidor/n8n/flujo_drive_a_tablero.json` y `servidor/n8n/flujo_copia_al_drive.json`. En cada uno:
   - elegir la credencial en los nodos de Google Drive;
   - reemplazar `PEGAR_ID_DE_LA_CARPETA_DE_DRIVE` por el ID de la carpeta (la parte final de la URL, después de `/folders/`);
   - **opcional:** en "Avisar al equipo", configurar SMTP, poner los mails y activar el nodo.
6. **Publicar** los dos flujos (botón *Publish*).
7. Agregar `servidor/.env` al comando de backup de la sección 6 y hacer estas pruebas:
   - guardar un Excel suelto en el Drive: en 15 minutos o menos, el tablero muestra "Mes pendiente";
   - agregar los otros 3: en 15 minutos o menos, el mes queda cargado;
   - subir un mes con el botón: los 4 archivos aparecen en el Drive.
8. Volver a agregar en el tablero (`tablero/plantilla_v2.html`) la mención a la carpeta de Drive en los textos de ayuda.

**Ojo:** en el camino del Drive no hay pantalla de confirmación. Un archivo nuevo que entra por ahí reemplaza al mes sin preguntar.
