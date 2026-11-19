# Pipeline DAO SRL: de los Excel de Luciano al tablero

## Cómo se usa

```bash
pip install openpyxl xlrd
python dao_pipeline.py --entrada <carpeta_de_Luciano> --base <carpeta_base> --salida <carpeta_del_tablero>/data
```

- **entrada:** la carpeta sincronizada con el Drive de Luciano. Puede tener subcarpetas y archivos de otros meses o años.
- **base:** la base histórica. Hay un JSON por fuente y por mes (`base/ER/2026-08.json`, etc.), más `registro_cargas.csv` con cada corrida.
- **salida:** `detalle_clientes_articulos.csv`, `resumen_mensual.csv` y `calidad.json`. El tablero los lee de la carpeta `data/` que está junto al HTML.

**Código de salida (v1.1):**
- **0:** todo bien, o solo errores ya informados en corridas anteriores, o ya había otra carga en curso y esta no hizo nada.
- **1:** apareció **un error nuevo**, es decir, un archivo que falla por primera vez. Es el momento de avisar.

Un Excel roto que sigue en la carpeta se muestra siempre en la franja roja del tablero, pero no vuelve a disparar la alerta en cada corrida. El detalle técnico de cada error queda en `registro_cargas.csv`.

## Control de errores (v1.1)

- **Los errores llegan al tablero.** `calidad.json` trae `errores_carga`: archivo, mes, un mensaje en castellano que dice qué hacer, si el error es nuevo y desde cuándo está.
- **Mensajes claros.** Cada rotura típica explica qué falta y qué hacer. Por ejemplo: falta una hoja o una columna, no está la fila "CLIENTE", el Excel está dañado, el nombre no coincide con el mes.
- **Resumen incompleto.** Si le falta una hoja, se carga igual y el mes queda en "revisar", con un aviso de qué datos faltan. Si le falta la hoja "Resumen", se rechaza.
- **Texto en una celda numérica.** Se toma como 0 y queda un aviso con el texto encontrado.
- **Escritura segura.** La base y los CSV se escriben en un temporal que se renombra al final. Si la corrida se corta, queda la versión anterior entera. Un registro de la base que aparezca dañado se aparta como `.corrupto` y se vuelve a armar desde el Excel.
- **Una sola corrida a la vez.** Mientras corre, existe el archivo `base/.carga_en_curso`. Si se lanza otra corrida, termina enseguida sin hacer nada. Si quedó un bloqueo de hace más de 60 minutos, por una corrida que se cortó, se descarta.
- **Pruebas:** `python tests/test_pipeline.py --excel <carpeta>`. Son 63 pruebas: la carga incremental, el control de errores, la carga por mes completo y la validación de la carga desde el tablero.

## Carga por mes completo (v1.2)

- **Regla:** un mes se carga solo cuando están sus **4 archivos**: ER, WEG, COMPRAS y un Resumen que traiga la columna de ese mes. Vale igual para el botón "Subir Excel del mes" y para los archivos que se dejan en la carpeta de entrada.
- **Mes incompleto = pendiente.** Si falta alguno, el mes no se toca y `calidad.json` lo lista en `pendientes` (qué llegó y qué falta). El tablero lo muestra en la franja amarilla "Mes pendiente". Si ese mes ya estaba cargado, sigue la versión anterior.
- **Botón "Subir Excel del mes":** `cargar_mes.py` valida los 4 juntos antes de tocar nada: que sean 4, `.xls` o `.xlsx`, de hasta 20 MB, uno de cada tipo, del mismo mes, que se puedan leer y que el Resumen tenga ese mes. Si algo falla, no se carga nada y se explica qué corregir. Si el pipeline falla después, se deshace la carga.
- **Corregir un mes ya cargado:** se vuelven a subir los 4 y se confirma el reemplazo. Los Excel anteriores de ese mes se mueven a `excel/_reemplazados/<fecha>/` y no se borran. Si los 4 son iguales a los ya cargados, el mensaje es "No hubo cambios".
- Las carpetas que empiezan con `_` o `.` dentro de la carpeta de Excel no se leen.
- Uso por línea de comandos: `python pipeline/cargar_mes.py --entrada <excel> --base <base> --salida <data> [--reemplazar] ER.xlsx WEG.xlsx COMPRAS.xlsx Resumen.xlsx` (sale con 0 si se cargó, 1 si se rechazó o falta confirmar).

## Validación de la carga desde el tablero (v1.3)

- **Mes del nombre contra las fechas:** en WEG y COMPRAS se compara el mes del nombre con la fecha de los comprobantes FC/CC/NC/ND. Si **más de la mitad** son de otro mes, el archivo se rechaza. Si son menos, se carga y queda un aviso. Medido con los 42 archivos reales de ene-2025 a sep-2026: todos sus comprobantes son del mes del nombre. Vale para cualquier archivo que lea el pipeline.
- **Mes ya cargado con archivos distintos:** el botón pide confirmación antes de reemplazar (`requiere_confirmacion`, respuesta 409 del servicio de carga). Se confirma reenviando con `confirmar_reemplazo=si`.
- **Otros meses que cambiaron:** al terminar, el resultado lista en `otros_meses` los demás meses cuyos valores de `resumen_mensual.csv` cambiaron, con el valor anterior y el nuevo. Se comparan los valores como se muestran, para no informar diferencias de redondeo.
- **Archivos más viejos:** se rechaza, aun confirmando, un ER, WEG o COMPRAS con fecha de emisión anterior al ya cargado, o un Resumen que llega a menos meses que el cargado para ese mes.

## Reglas de la carga (por qué nunca se pierde información)

1. **Un mes = una unidad.** Cada archivo se guarda en la base bajo su mes. El Resumen se parte columna por columna y cada columna es un mes.
2. **Upsert.** Un mes nuevo se agrega y un mes que ya existe solo se reemplaza si el archivo que llega es más nuevo. Nunca se borra un mes.
   - Para ER, WEG y COMPRAS se usa la fecha de emisión que el sistema escribe en A1 ("Fecha: 03/09/26 Hora: 11:25:13").
   - Entre dos Resumen gana el que cubre meses más recientes. Por ejemplo, el interanual 2026, que llega a agosto 2026, le gana al anual 2025 para ago–dic 2025. Así, cuando el interanual de 2027 ya no traiga 2025, esos meses siguen en la base.
3. **Idempotente.** Correr dos veces lo mismo no cambia nada: el registro dice "sin cambios".
4. **Validación al entrar.** Si el ER dice "DESDE 08/2026" pero el archivo se llama `ER_09_2026`, se rechaza con un error claro y no se toca la base.
5. **Un archivo roto no frena al resto:** se registra el error y sigue con los demás.

## Nombres de archivo aceptados

Los nombres reales mezclan espacios, guiones y mayúsculas. El patrón acepta todas estas variantes:

| Tipo | Ejemplos reales | Regla |
|---|---|---|
| ER | `ER_01-2025_SUPERFICIE.XLS`, `ER_01-2026.XLS`, `ER_03_2026.XLS` | `ER` + mes (1-2 dígitos) + año (4), separados por espacio, `_` o `-`; lo que sigue se ignora |
| WEG | `WEG 01_2025.xls`, `WEG 08_2026.XLS` | igual |
| COMPRAS | `COMPRAS 01_2025.xls` | igual |
| Resumen | `Resumen 2025.xls`, `Resumen_InterAnual_2026.XLS` | contiene "resumen" |

Se ignoran los archivos temporales de Excel (`~$...`).

## Qué se saca de cada archivo

El pipeline **no usa celdas fijas**. Cada dato se ubica por su rótulo, porque la estructura cambia de un mes a otro (ver la auditoría más abajo).

### ER_MM_AAAA (3 hojas)

| Hoja | Cómo se ubica | Qué se toma |
|---|---|---|
| ESTADO_DE_RESULTADO | fila "DESDE" | mes del reporte (validación) |
| | filas entre el encabezado "CLIENTE" y "TOTAL CV" | por cliente: código, nombre, MALBEC, NETO, **TOTAL** = ingreso |
| | fila sin código justo debajo del encabezado | ingreso total del mes |
| | filas "Pintura", "Químicos", "MOD", "Contribucion Maginal", "Costos fijos", "Resultado bruto" (último número de la fila) | totales teóricos (Fx) |
| SUP_PINTADA | fila cuyo A dice "ARTICULO" (a veces es la fila 1, a veces la 2) | por artículo y cliente: **TOTAL SUP.** (m²), CANT. MALBEC + CANT. DAO (piezas) |
| DETALLE_DE_CALCULOS | cada fila con código en A abre un artículo | por artículo: suma de "Pintura monto", "Químicos monto", "MOD monto" (Malbec + DAO) y "COSTO FIJO X PIEZA" |

Controles automáticos del ER, todos en verde para los 20 meses:
- suma de clientes = ingreso total;
- suma de recetas = Pintura, Químicos y MOD del ER;
- costo fijo por pieza × piezas = Costos fijos del ER;
- ingreso repartido = ingreso total.

### WEG MM_AAAA (pintura real)

Hay **tres formatos** según la época, y el pipeline reconoce cada uno solo:

| Período | Formato | Qué se suma |
|---|---|---|
| ene–jun 2025 | listado de comprobantes con renglones (como COMPRAS) | TOTAL de cada renglón de comprobantes FC/CC/NC/ND (los PA son pagos y se descartan; las notas de crédito restan) |
| jul 2025 | "por artículo" sin título, totales en la fila 2 | subtotal de cada artículo (filas "(código) descripción") |
| ago 2025 en adelante | "COMPRAS POR ARTICULO" | subtotal de cada artículo |

También clasifica cada renglón:
- **Inversión:** compras únicas (EQUIPO, PISTOLA). Quedan fuera del resultado y se muestran aparte.
- **Repuestos y reparaciones** (ROD HOLDER, MANGUERA, REPARACIÓN): pasan de pintura a costo fijo, porque se repiten.
- **Químicos comprados a WEG** (CHEMISA, FOSFODESENGRASE): siguen dentro de pintura, como hace Luciano.

### COMPRAS MM_AAAA

- Toma comprobantes FC/CC/NC/ND (los PA son pagos y se descartan) y usa **NETO** (el sistema lo exporta en negativo).
- Proveedor EMPRENDIMIENTOS → pintura; WHOR QUIMICA → **químicos reales**; el resto → **gastos fijos reales**.
- Las excepciones se configuran sin tocar código:
  - `base/reglas_compras.csv`: por proveedor. Por ejemplo, INDELPLAST → a_confirmar (se muestra aparte, fuera del resultado).
  - `base/ajustes_comprobantes.csv`: por comprobante. Por ejemplo, FC-A-3-11677 → inversion.
  - Categorías: pintura, quimicos, gasto_fijo, inversion, a_confirmar, excluir.

### Resumen (anual o interanual)

- La fila de meses se busca en las primeras 8 filas. Acepta fechas (interanual) o "Ene…Dic" con el año en la fila de arriba (anual 2025).
- La columna "Media" se ignora, porque el tablero recalcula todo.
- El amarillo del año anterior no hace falta: cada columna trae su fecha.

| Hoja | Rótulo | Campo |
|---|---|---|
| Resumen | M2, Ingresos, CV (Fx), CV ®, CF (Fx), CF ®, Retiros Socios | control y retiros |
| CV | Pintura ®, Quimicos ®, MOD ® | pintura, químicos y **mano de obra real** |
| Quimicos | Quimicos ®, Quimicos (Fx) | químicos |
| Pintura | Pintura (Fx), Kg DE PINTURA ® y (Fx), Lts DE QUIMICOS, $ Pagados ® | pintura teórica y kg |
| MOD | MOD (Fx) | MOD teórica |
| CF (Calculo) | CF ®, Sueldo … (todas las filas), Cargas Sociales | gastos, **sueldos** y **cargas sociales** |
| Cliente $ / Cliente m2 | una fila por cliente | ingresos y m² por cliente (control) |

## De dónde sale cada número del tablero

| Dato | Fuente principal | Respaldo si falta |
|---|---|---|
| m², ingresos, costo teórico (pintura, químicos, MOD, CF) | Resumen | ER del mes |
| Pintura real | Resumen | WEG |
| Químicos reales | Resumen (hoja CV, la que suma en sus totales; si no coincide con la hoja Quimicos, deja un aviso) | COMPRAS |
| Gastos fijos reales | Resumen | COMPRAS + reglas |
| MOD real, sueldos, cargas sociales, retiros, kg | **solo Resumen** | — |
| Detalle por cliente y artículo | ER (ajustado para sumar igual que el Resumen) | — |

**¿Por qué el Resumen manda a nivel empresa?** Es el número que Luciano mira y ya incluye su criterio: qué es inversión, dónde imputa un equipo, qué notas de crédito descuenta. Los archivos mensuales quedan como control y como respaldo si un mes no está en el Resumen.

## Auditoría de estructura (ene-2025 a ago-2026: 60 archivos mensuales y 2 Resumen)

| Archivo | ¿Igual todos los meses? | Qué cambia |
|---|---|---|
| ER: hojas y encabezados | Sí | — |
| ER: ESTADO_DE_RESULTADO | No | el bloque de costos se corre entre las filas 17 y 26 según la cantidad de clientes |
| ER: SUP_PINTADA | No | en may, jun y jul 2026 hay una fila de totales arriba del encabezado |
| ER: DETALLE_DE_CALCULOS | Sí en columnas | algunas líneas de receta vienen sin cantidad (por eso las piezas se toman de SUP_PINTADA) |
| WEG | No | 3 formatos distintos (ver arriba) |
| COMPRAS | Sí | — |
| Resumen | No | anual con "Ene…Dic" y año arriba; interanual con fechas; la fila de meses y la columna "Media" cambian de lugar según la hoja |
| Nombres de archivo | No | espacios, guiones, `_SUPERFICIE` solo en 2025, .xls/.XLS |

**Conclusión:** con celdas fijas el pipeline se habría roto en la mayoría de los meses (el ER solo tiene 8 disposiciones distintas en 20 meses). Por eso todo se lee por rótulo.

## Decisiones acordadas

Las decisiones de criterio (químicos, retiros, equipos, INDELPLAST, ESSEN, sueldos) y el diccionario de cada gráfico están en el documento "Tablero DAO SRL: definiciones y guía de lectura". Ante una duda en los datos se sigue al Resumen de Luciano y queda un aviso en "Estado de los datos".

## Datos que no están en ningún archivo mensual

El Resumen es la **única** fuente de estos cinco datos. Si Luciano deja de actualizarlo, el tablero pierde el costo real completo:
- mano de obra directa real (MOD ®);
- sueldos (hoy Cecilia, Rodrigo y Mauri);
- cargas sociales;
- retiros de socios;
- kg de pintura (se pueden estimar desde WEG, con ±5% de diferencia).

Si Luciano no quiere mostrar los sueldos por persona, alcanza con una sola fila "Sueldos y cargas" con el total del mes en la hoja "CF (Calculo)": el pipeline la reconoce. Si deja las celdas vacías, el mes queda marcado como incompleto y el tablero usa el costo teórico.
