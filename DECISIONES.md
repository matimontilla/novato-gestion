# Decisiones y criterios — Novato Gestión

Memoria del proyecto: **por qué** las cosas están como están. El código dice *qué*
hace; esto dice *por qué*. Si algo acá parece raro, probablemente hay una razón
dolorosa detrás.

Última actualización: noviembre 2026.

---

## 1. Criterios contables

### CI vs CO — la distinción más importante

Hay dos tipos de costo que NO son directos de un producto, y se tratan distinto:

| | **CI** (Costo Indirecto) | **CO** (Costo Operativo) |
|---|---|---|
| Qué es | Costo de **producción** compartido entre varios productos (tapones, flete, elaboración) | Costo de **la empresa como entidad** (AFIP, contadores, constitución de la SAS) |
| Producto asignado | Sí, prorrateado por producción real (%ANUAL) | **NO** — va sin producto |
| Afecta costo/botella | Sí | **No** |

**Por qué:** un vino que se embotelló en mayor cantidad absorbe más del costo
compartido de producirlo — eso es CI. Pero mantener la SAS al día no es un costo de
elaborar vino; si lo prorrateás, inflás artificialmente el costo por botella con algo
que no tiene relación causal con producir. Por eso CO va sin producto.

**Descartado:** prorratear CO por *stock disponible al momento del gasto*. Suena
razonable pero falla: el mismo gasto se repartiría distinto según qué día lo cargues
(el stock cambia por ventas, no por nada relacionado al gasto), y un producto agotado
se llevaría $0 aunque haya estado activo todo el año que existió la empresa.

### Referencias de costos históricos

El código `CCI22-001` se usó originalmente para nuclear los costos indirectos de los
varietales 2022, pero después quedó usándose por inercia durante casi 3 años para
cosas que no correspondían. Se separó en tres referencias limpias:

| Referencia | Qué contiene | Filas CAJA | Monto |
|---|---|---|---|
| `CCI22-001` | Costo indirecto real 2022 + envíos de esa partida | 12 | $156.909 |
| `CCI23-001` | CO: AFIP/Contadores (constitución SAS), sin producto | 23 | $1.056.577 |
| `SEM23-001` | Préstamo Semilla → financió producción de Blend 2023 | 35 | $721.118 |

El criterio de separación **no fue por año** sino por a qué correspondía cada pago
(AFIP/Contadores de cualquier año → CO; los ya etiquetados con el producto → Semilla).

### Fechas de operaciones agregadas

Las filas de BALANCE que resumen muchos pagos a lo largo de años (CO, Semilla) llevan
la **fecha mediana ponderada por monto** de sus pagos reales — la fecha donde se
acumuló la mitad de la plata gastada:

- CO (`CCI23-001`) → **25/11/2024**
- Semilla (`SEM23-001`) → **22/08/2024**

**Por qué:** la fecha determina a qué dólar se valúa el monto. Si les ponés la fecha
de hoy, convertís pesos de 2022-2025 al dólar actual y el valor en USD queda mal
(CO daba USD 691 con dólar de hoy vs USD 935 con el dólar real de la época).

Recalculable con `corregirFechasCOSemilla()`.

### Transacciones "en papel"

Ventas contables sin entrega física (ej. Yuniku / Chardonnay 2022) van con **botellas
en cero**, para no descontar stock que no se movió. Las fórmulas de costo unitario
contemplan este caso: si `M=0`, CU$ y CU US$ quedan vacíos en vez de dar `#DIV/0!`.

---

## 2. Reglas técnicas — Apps Script

### NUNCA usar setFormulas() sobre un rango completo

**Esta lección costó una pérdida de datos real.**

```javascript
// MAL — destructivo
var formulas = rango.getFormulas();   // devuelve '' para celdas con valor plano
formulas[i][0] = nuevaFormula;         // modificás una
rango.setFormulas(formulas);           // ← pisa con vacío TODAS las que tenían números

// BIEN — sólo la celda puntual
sheet.getRange(fila, col).setFormula(nuevaFormula);
```

`getFormulas()` devuelve string vacío para **cualquier** celda que no sea fórmula,
incluidas las que tienen un número cargado a mano. Al reescribir el array completo,
esos números se borran. Usar `getFormulas()` sólo para **identificar** filas; escribir
siempre celda por celda.

Red de seguridad ante un error así: **Archivo → Historial de versiones** en Sheets.

### Rangos siempre abiertos

```
MAL:  CAJA!$I$3:$I999      BALANCE!$B$3:$B$369     BLUE_API!$A$56:$A$4644
BIEN: CAJA!$I$3:$I         BALANCE!$B$3:$B         BLUE_API!$A$2:$A
```

Los rangos con tope fijo se rompen de dos formas: cuando la hoja crece más allá del
tope (los datos nuevos dejan de contarse, **en silencio**), y cuando las filas se
corren por inserciones (el rango apunta a otro lado).

Casos reales que esto causó: los saldos de BALANCE dejaban de sumar cobros al pasar
la fila 999; las stats de CLIENTES iban a dejar de contar ventas al pasar la fila 369
(estaba en 367 cuando se detectó); la valuación de STOCK tomaba una fecha vieja.

Reparación: `normalizarRangosAbiertos()`.

### Insertar filas en batch, no de a una

Insertar miles de filas con `insertRowAfter()` en un loop cuelga la ejecución y supera
el límite de 6 minutos. Para volumen: `insertRowsAfter(fila, n)` + un solo `setValues()`.

### Zona horaria: todo en America/Argentina/Mendoza

El script (manifiesto `appsscript.json`) y la planilla (Archivo → Configuración) están
en **America/Argentina/Mendoza**. Desde dónde se edite no importa: la zona es del archivo.

**Antecedente (octubre 2026):** el script estaba en Australia/Darwin y la planilla en
America/Los_Angeles. Toda fecha cargada desde la app quedaba **un día antes**, y como
BLUE_API también estaba corrida (con otra hora), esas operaciones se convertían con la
cotización del día hábil anterior. Las horas raras (07:30, 19:30) eran el síntoma.
Se corrigió con `migrarZonaHorariaMendoza()`, identificando las filas por su "huella"
de hora UTC (14:30 o 02:30) sin tocar las cargadas a mano.

**Reglas que quedan:**
- BLUE_API a **medianoche** de Mendoza: la búsqueda "exacto o anterior" encuentra la
  cotización del mismo día para una operación a cualquier hora.
- En la app, la fecha de "hoy" se toma del reloj del dispositivo (`fechaLocalHoy()`),
  nunca de `toISOString()`, que da la fecha en UTC.
- Si alguna vez se vuelven a ver horas raras en las fechas, revisar primero que las dos
  zonas sigan en Mendoza (`diagnosticarZonasHorarias()`).

### Fórmulas con locale es-AR

Las fórmulas escritas desde Apps Script usan **punto y coma** como separador, no coma:
`=IF(A1=0;"";B1)`. Ojo: `getFormulas()` puede devolverlas con comas — no asumir el
separador al hacer búsquedas o regex.

### Blindar divisiones

Toda conversión a dólar va envuelta en `IFERROR(...;"")`. Si el `XLOOKUP` a BLUE_API
no resuelve una fecha puntual, la celda queda vacía en vez de `#DIV/0!`.

**Por qué importa tanto:** un solo `#DIV/0!` en la columna G de CAJA se propaga vía
`SUMIF` a las columnas P/I/Q de BALANCE y de ahí a CLIENTES. Dos celdas rotas
generaban 10 errores en cascada. Blindar la raíz limpia todo.

---

## 3. Arquitectura de datos

### Cuadro "CAJAS" — pesos calculados, dólares declarados

El bloque resumen al pie de la pestaña CAJA tiene **dos criterios distintos** por
columna, y es intencional:

| Columna | Criterio | Por qué |
|---|---|---|
| **C (AR$)** | **Calculado** — `=SUMIF(CAJA!$H$3:$H;$B<fila>;CAJA!$F$3:$F)` | El saldo en pesos es simplemente la suma histórica de todo lo que pasó por esa caja |
| **D (USD)** | **Declarado a mano** | Depende del tipo de cambio al que se compraron esos dólares, si vinieron de cripto, etc. **No es derivable** del histórico de movimientos |

**No convertir la columna USD en fórmula.** Parece una inconsistencia a corregir, pero
es una decisión deliberada: automatizarla daría un número incorrecto, porque el valor
de esos dólares depende de cómo se adquirieron, no de cuántos pesos se movieron.

Antecedente: en algún momento se pegaron valores encima de las fórmulas de AR$ y el
cuadro dejó de actualizarse (síntoma delator: LUDICO quedó en `0.0012` — residuo de
coma flotante de la fórmula original). Reparación: `repararCuadroCajasArs()`, que
sólo toca la columna C.

`getResumenCajas()` **lee** este bloque, no lo calcula — si los valores están mal en
la planilla, la app los muestra mal.

### Columna S (DEPOSITO) — por qué está al final y no al lado de PRODUCTO

BALANCE registra en la columna **S** de qué depósito salieron las botellas en cada
carga hecha desde la app (dato que antes se usaba para descontar stock y se descartaba).

Está al final, después de `R AÑO`, y no insertada al lado de PRODUCTO donde sería más
lógico visualmente. **Motivo:** insertar una columna en el medio corre F..R una
posición. Sheets ajusta las fórmulas existentes, pero el código escribe sus fórmulas
con **letras de columna fijas** (`G`, `H`, `M`, `O`...), así que pasaría a escribir en
las columnas equivocadas — un desastre silencioso.

> **Regla general: nunca insertar columnas en el medio de BALANCE o CAJA.** Agregar al
> final es seguro. *Mover* una columna existente también lo es (Sheets reajusta las
> referencias); lo que rompe es *insertar* y correr todo.

Sólo se llena cuando hay botellas (si no hubo movimiento físico, queda vacío). En
operaciones multi-producto, cada línea registra su propio depósito. No se llenó para
atrás: aplica sólo a las cargas nuevas.

Setup inicial: `prepararColumnaDeposito()`.

### Filtros en encabezados

`agregarFiltrosEncabezados()` pone filtros en BALANCE, CAJA y STOCK (CLIENTES ya los
tenía).

**Cuidado:** filtrar es inofensivo, pero el mismo menú ofrece **ordenar** (A-Z, Z-A), y
eso reordena las filas físicamente **para todos**. En BALANCE eso desarmaría el orden
cronológico y la posición en que se acomodaron a mano las filas de CI/CO. Para explorar
sin afectar la hoja: **Datos → Vistas de filtro**, que son personales.

### Operaciones multi-producto

Una operación (venta o compra) con varios productos = **varias filas de BALANCE con la
misma REFERENCIA**.

El cobro se reparte **proporcionalmente** entre los productos (`repartirProporcional`),
generando una fila de CAJA por producto, cada una con su producto en la columna E. El
ajuste de redondeo va en la última parte para que la suma cierre exacto.

Por eso el saldo de cada fila filtra por **referencia Y producto**:

```
=G - SUMIFS(CAJA!$F$3:$F; CAJA!$I$3:$I; L; CAJA!$E$3:$E; E)
```

Sin el filtro de producto, cada línea restaría el cobro completo → doble conteo. Con
el filtro, cada línea resta sólo lo cobrado de su producto. Funciona igual para
operaciones de un solo producto (el filtro extra no cambia nada).

Esto permite tener saldo en USD y rentabilidad por diferencia de TC **por producto**.

### BLUE_API

- Fuente: `https://api.bluelytics.com.ar/v2/evolution.json`
- Filtrar por `source` = "blue" **case-insensitive** (la API devuelve `"Blue"`)
- Orden: **descendente** (más reciente arriba)
- Columnas: **B = compra** (`value_buy`, más bajo), **C = venta** (`value_sell`, más alto)
- Las fórmulas de la planilla convierten contra la **columna C (venta)**
- Se actualiza sola con un trigger diario (`instalarTriggerBlueApi()`)

`actualizarBlueApi()` **reconstruye la tabla completa** en vez de parchar filas: es
idempotente y arregla cualquier desorden previo de una.

---

## 4. Trampas conocidas

### La descarga xlsx de Google Drive queda CACHEADA

Descargar el Sheet como .xlsx puede devolver un snapshot viejo que **no refleja los
cambios recientes hechos por Apps Script**. Esto causó un rato largo de confusión:
la descarga mostraba rangos cerrados que en vivo ya estaban abiertos.

**La fuente de verdad es Apps Script** (`getFormulas()` / `getValues()`), no la
descarga. Para auditar el estado real: `diagnosticarErrores()`.

### Archivos del backend

- `novatobackend.gs`: lo que usa la app (endpoints, helpers, disparador de BLUE_API).
- `seguridad.gs`: login con PIN en el servidor, tokens, candado, anti-duplicados y
  validación de datos (ver "Seguridad del backend").
- `herramientas.gs`: diagnósticos, reparaciones reutilizables y configuración, para
  correr a mano con ▶ Run. El desplegable del editor muestra sólo las funciones del
  archivo abierto.

Las funciones de una sola vez que ya se corrieron (reclasificación de costos, columnas
DEPOSITO y DIAS CIERRE, migración de zona horaria, etc.) **se borraron en octubre 2026**.
Este documento las sigue nombrando como antecedente; el código está en el historial de
GitHub (anterior al commit que crea `herramientas.gs`).

### Seguridad del backend (desde noviembre 2026)

**Por qué:** el backend es una web app pública (`ANYONE_ANONYMOUS`) y la URL va dentro
del código de la app. Antes, el PIN sólo existía en el celular (por defecto `1234`) y el
servidor no validaba nada: cualquiera con la URL podía leer la plata de la bodega o
escribir en el Sheet, y el nombre de quien cargaba lo decía el propio cliente.

**Cómo funciona ahora (`seguridad.gs`):**
- **Login en el servidor:** la app manda usuario + PIN (`action=login`, por POST). El
  servidor compara contra un hash guardado en Script Properties (`PIN_<usuario>`, salado
  con `AUTH_SECRET`) y devuelve un token firmado (HMAC) que vale 12 h. 5 PIN incorrectos
  bloquean al usuario 15 minutos y avisan por Telegram.
- **El usuario sale del token**, no de lo que mande el cliente.
- **Escrituras protegidas:** un candado de script (una carga a la vez: sin referencias
  repetidas ni stock pisado entre dos socios), y un `requestId` por operación que el
  servidor recuerda 6 h: si la señal se corta y la app reintenta, no se duplica.
- **Validación:** caja, depósito, producto y tipo contra listas conocidas; montos,
  cantidades y fechas con rango; los textos libres pierden símbolos iniciales (`= + - @`)
  para que Sheets no los ejecute como fórmula. **El servidor también valida stock**
  (no deja ventas ni movimientos que dejen un depósito en negativo, y chequea todas las
  líneas ANTES de escribir la primera).
- **Control de stock:** fecha, hora y "teórico" los fija el servidor (hora de Mendoza),
  no el celular, que puede tener datos viejos.
- Un error del backend ahora se muestra al usuario (antes un `{error}` se leía como éxito).

**Rollout seguro (la bandera `AUTH_REQUIRED`):** el código nuevo funciona en modo
"transición" hasta que se active la obligatoriedad, así no hay un momento en que la
app y el backend queden desincronizados:
1. Mergear → se despliegan backend (Action) y app (Vercel).
2. Correr `configurarAutenticacion()` (crea la clave y deja el PIN inicial `1234` a los
   que no tengan). Mientras no se corra, la app usa el PIN local de antes.
3. Cada socio entra a la app y cambia su PIN (⚙ → Cambiar PIN).
4. Con todos adentro y probado, correr `activarAutenticacionObligatoria()`.
5. Marcha atrás de emergencia: `desactivarAutenticacionObligatoria()`.
`estadoAutenticacion()` muestra en qué punto está. Si alguien olvida el PIN:
`restablecerPin()`.

**Límites que conviene recordar:** son 4 dígitos (el bloqueo por intentos es lo que lo
hace razonable, no la longitud); el token vive sólo en memoria, así que al recargar la
app se pide el PIN; la lectura de datos también exige sesión cuando la bandera está activa.
**Sin verificar en el Sheet real al escribir esto:** el POST desde el navegador a Apps
Script (la app cae a GET con el mismo `requestId` si el POST falla por red/CORS).

### Pruebas automáticas y CI

`npm test` corre `tests/backend.test.js`: carga `seguridad.gs` y `novatobackend.gs` con los
servicios de Apps Script simulados y prueba login, tokens, bloqueo, candado, duplicados,
validaciones y reparto proporcional. No toca el Sheet. La Action `ci.yml` corre las
pruebas, revisa la sintaxis de los `.gs` y compila la app en cada pull request y push
(un `.gs` con error de sintaxis desplegaría un backend roto porque `clasp push` reemplaza todo).

### Carga rápida con IA (desde noviembre 2026)

En la pantalla **Cargar** hay un cuadro donde se escribe (o se dicta) una frase, por ejemplo
"vendí 12 Malbec 2022 a Organyca a 9000 cada una", y la app completa el formulario.
Código: `apps-script/asistente.gs` (acción `interpretarCarga`) y `CargaRapida` en `src/App.jsx`.

**Principios que no se tocan:**
- **Sólo propone, nunca escribe.** El servidor devuelve un *borrador*; el usuario lo revisa en
  el formulario de siempre y recién al apretar "Registrar" pasa por los endpoints normales
  (con toda su validación de stock, candado y duplicados). La IA no tiene acceso de escritura.
- **Todo lo que devuelve el modelo se valida contra listas conocidas** (productos, depósitos,
  cajas, categorías, canales). Lo que no existe se vacía y se avisa. Los montos
  (precio × cantidad) los calcula el código, no el modelo.
- **Privacidad:** a la API sólo viajan la frase, la fecha de hoy y los *nombres* del catálogo
  (clientes, productos, depósitos). Nunca saldos, montos de la planilla ni PINs.
- **Requiere sesión del servidor** (haber entrado con el PIN del servidor). Sin sesión no corre.

**Configuración (Script Properties de Apps Script, nunca en el repo):**
- `ANTHROPIC_API_KEY`: clave de la API de Anthropic (cuenta con crédito cargado). Si falta, la
  función se apaga sola y la app muestra "todavía no está activada"; el resto sigue igual.
- `CLAUDE_MODEL`: opcional. Por defecto `claude-haiku-5-5`.

**Unidades:** una caja = 6 botellas (`BOTELLAS_POR_CAJA` en `asistente.gs`). El modelo sólo informa
cuántas cajas dijeron el usuario; la conversión y el total (precio por caja × cajas) los hace el código.

**Topes:** 500 caracteres por frase, 30 frases por usuario por hora, 300 por día entre todos
(constantes al inicio de `asistente.gs`). Es el tope de gasto si alguien abusa.

**Costo:** el modelo por defecto cuesta US$0,10 por millón de tokens de entrada y US$0,50 por
millón de salida (precio verificado en la documentación de Anthropic al diseñarlo; confirmar
en la página de precios antes de apoyarse en el número). Con ~1.500 tokens de entrada y ~200
de salida por frase (estimación, no medida), cada uso cuesta del orden de una fracción de
centavo. El tope diario de 300 frases acota el peor caso.

**Voz:** si el navegador lo permite (Web Speech API, es-AR) aparece un botón de micrófono; si
no, el micrófono del teclado del celular dicta en el mismo cuadro. El audio no pasa por el
servidor de Novato.

**Sin verificar contra el sistema real al escribir esto:** la respuesta de la API de Anthropic
y el comportamiento sobre la planilla real; las pruebas usan servicios simulados. Probar con
frases reales después del primer deploy (ventas, cobros y movimientos de stock).

### App instalable (PWA)

La app se puede "Agregar a pantalla de inicio" y abre como app (sin barra del navegador).
Archivos: `public/manifest.webmanifest`, `public/sw.js`, `public/icons/` y el registro en
`src/main.jsx` (sólo en producción). Es un commit separado: para sacarla basta revertirlo.
- **No guarda datos del negocio.** El service worker sólo guarda la carcasa (la página y los
  archivos con hash). Las llamadas a Apps Script y a otros dominios van siempre a la red.
  Sin conexión no hay datos ni operaciones: es a propósito, para no operar con stock viejo.
- Costo: ninguno (se sirve desde Vercel como el resto).
- En iPhone se instala desde Safari ("Compartir" → "Agregar a inicio"); en Android, desde
  Chrome ("Instalar aplicación"). Los íconos son una "N" dorada generada para esta app.
- Al publicar una versión nueva, las páginas se piden primero a la red, así que se ve la
  actualización sin borrar nada. Si cambia la lógica del service worker, subir `CACHE`.

### Control físico de stock: qué se ve al entrar

La pestaña Stock arranca con los conteos **vacíos** (antes arrancaban en 0 y mostraban todo como
faltante). Arriba queda fijo el **último control de ese depósito** (fecha, hora y lo contado);
el teórico es el stock de hoy (último control menos ventas y movimientos). La diferencia aparece
sólo al cargar el conteo nuevo, y al confirmar el servidor fija STOCK en lo contado.
El backend devuelve `ultimosControles` (uno por depósito) además de `ultimoControlStock`.
En "Para resolver" del Inicio van **todas** las ventas con más de 60 días sin cobrar y, al final,
el aviso de control de stock.

### Deploy: automático desde GitHub (desde octubre 2026)

Cada push a `main` que toca `apps-script/` dispara la GitHub Action
`.github/workflows/deploy-apps-script.yml`, que sube el código con `clasp push` y
actualiza el **mismo** deployment con `clasp update-deployment` (la URL no cambia).
El frontend lo despliega Vercel. **Ya no hace falta pegar código ni hacer
"New version" a mano.** También se puede correr a mano desde la pestaña Actions.

Secretos en GitHub: `CLASPRC_JSON` (credenciales de `clasp login`), `SCRIPT_ID`,
`DEPLOYMENT_ID`. Si faltan, la Action no hace nada.

**Cuidados:**
- `clasp push` **reemplaza todo** el proyecto online con `apps-script/`. El repo tiene
  que tener exactamente los mismos archivos: `novatobackend.gs` (el nombre importa: con
  otro nombre borraría el archivo online) y `appsscript.json`.
- **No editar el código directo en el editor de Apps Script:** el próximo deploy pisa
  esos cambios. Todo cambio va por el repo.
- Lo que **sigue siendo manual**: las funciones de una sola vez (reparaciones,
  `prepararColumna...`) se corren con ▶ Run en el editor.

**Zona horaria:** ver la sección dedicada; script y planilla en Mendoza.

---

## 5. Fuera de alcance

- **Filas vacías de buffer**: ya no hacen falta (las cargas insertan su propia fila).
  Si se limpian, conservar la fila de TOTALES de BALANCE y el bloque "CAJAS" de CAJA.

### Pestaña DINAMICOS — obsoleta

Tenía tablas dinámicas nativas de Sheets (con `#REF!` de arrastre) que se rearmaban a
mano. Quedó reemplazada por la pestaña **Datos** de la app, que arma las mismas
relaciones desde `getAnalytics()` y se actualiza sola.

Antes de borrarla conviene: (1) buscar `DINAMICOS` en todas las hojas *dentro de
fórmulas*, para confirmar que nadie la referencia — si algo la usa, borrarla deja
`#REF!`; (2) sacar una captura por si había algún corte que la app no replica.
Ocultarla en vez de borrarla es la opción reversible.

---

## 5b. Analytics (pestaña Datos)

`getAnalytics()` hace todas las agregaciones en **una sola lectura** de BALANCE.
Gráficos dibujados en SVG a mano, sin librería: para barras alcanza, no infla el
bundle y respeta la estética de la app.

**Dos criterios de lectura que importan:**

- **Flujo anual ≠ rentabilidad.** Los costos de una añada se pagan antes de venderla,
  así que un año de cosecha grande da "pérdida" y uno vendiendo stock viejo da
  "ganancia". El flujo anual sirve para **liquidez**; la rentabilidad real se mide con
  el **margen por añada** (lo vendido de esa añada contra lo que costó producirla).
- **El margen por añada favorece a las añadas viejas**, que ya vendieron su stock. Una
  añada reciente tiene todo el costo cargado y sólo parte de las ventas hechas.

**Evolución de insumos:** sale de agrupar los egresos por DETALLE y añada usando la
columna **CU US$** (costo de la operación ÷ botellas producidas). Depende de que el
mismo insumo se cargue siempre con el mismo DETALLE — si a veces es "Etiquetas" y otras
"Etiquetado", aparecen como dos series distintas.

---

## 6. Trabajando con Claude en este proyecto

### Verificar, no recordar

**Regla:** todo dato numérico sobre la planilla, el código o el negocio se verifica
leyendo la fuente antes de afirmarlo. Nunca citar cifras de memoria.

**Por qué:** un número inventado que suena plausible es peor que un "no sé", porque
dispara análisis equivocados que cuestan tiempo. Caso real: se afirmó de memoria un
saldo pendiente de $1.362.917 en `CCI22-001`; el valor real era $156.909. Se
descubrió recién cuando el usuario mandó una captura de la planilla, después de
varios intercambios de análisis sobre una cifra falsa.

En la práctica:

- Datos de la planilla → leerlos con Apps Script (`diagnosticarErrores()`, etc.)
- Estado del código → abrirlo, no recordarlo
- Cualquier cosa posterior al corte de conocimiento del modelo → buscar en la web
- Al dar un número, aclarar si fue **verificado** o **recordado**

No aplica a razonamiento (discutir un criterio contable, analizar una fórmula): ahí
verificar no aporta y sólo hace lento el intercambio.

### Corte de conocimiento

Cada modelo tiene una fecha de corte fija que **no se actualiza**; sólo cambia si se
cambia de modelo. Claude Opus 5: **mayo 2026**.

Los meses cercanos al corte están peor cubiertos que los anteriores — Anthropic
distingue entre *reliable knowledge cutoff* (hasta donde el conocimiento es confiable)
y *training data cutoff* (rango más amplio, con los últimos meses "delgados").

Todo lo posterior es ciego salvo que se busque en la web. Preguntarle al modelo por su
propia versión no es confiable: en esta conversación reportó ser Opus 4.8 estando en
Opus 5, y negó la existencia de Opus 5 hasta que una búsqueda lo desmintió.

### Elección de modelo

- **Opus** para backend, fórmulas financieras, reconciliación — donde un error tiene
  costo real
- **Sonnet** alcanza de sobra para UI y frontend

### Convenciones de la app

- Estética: fondo casi negro, acentos dorados (etiqueta Novato), tipografía Georgia serif
  para cifras y títulos, system-ui para el resto
- Rediseño (nov 2026): navegación inferior con "Cargar" como botón central; íconos SVG propios
  en lugar de emojis; avatares con la inicial; texto mínimo de 12 px (11 px en gráficos);
  controles de al menos 44 px; contraste de `dim`/`muted` subido para leerse a la luz del día;
  opciones cortas (tipo de operación, cliente, depósito) como chips en vez de selects
- Comunicación del proyecto en español
- Productos con stock 0 se ocultan de los selectores
- Signos: cobro positivo, gasto negativo. Verde `+` / rojo `-`. Sobrepago se muestra
  distinto (azul, "a favor") de deuda real.

## Chardonnay / Yuniku / Organyca (corrección de registros, nov 2026)

Contexto: proyecto de elaboración por encargo para el restaurante Yuniku (nunca abrió). Novato elaboró, etiquetó y cobró (YUN-001, $1.847.400, 10 cobros en CAJA). Después Yuniku ofreció su stock para venderlo por los canales de Novato, pagándole un precio por botella vendida. Más tarde se entregó el Chardonnay restante a Organyca (nuevo proyecto) y el saldo se canceló con Malbec 2022.

- **Stock Chardonnay 2022 cierra en 0.** Las 2208 salidas son 1674 "Muestra"/entregas, 17 ajustes, 499 ventas a terceros y 18 retiros. YUN-001 va sin botellas (venta contable, ya había salido el vino en la entrega).
- **Pago a Yuniku ($320.000, 19-9-23).** Tiene fila en BALANCE (DETALLE "Pago", referencia CCH22-011) y el mismo código en CAJA, así que el saldo da 0. Precio por botella pactado, reconstruido: $320.000 ÷ 209 botellas vendidas a terceros hasta esa fecha ≈ $1.531 (cálculo de nov 2026, no un dato registrado entonces).
- **Categorías nuevas en BALANCE (DETALLE), todas sin monto:**
  - "Entrega / traspaso": el vino sigue siendo de otro y vuelve a su dueño (24 Chardonnay a Yuniku el 17-4-23, 1598 a Organyca el 1-7-25).
  - "Canje / pago en especie": 168 Malbec 2022 a Organyca el 1-7-25, que cancelaron el saldo restante. Se dejó sin valor en plata a propósito: valorarlo inflaría ingresos o abriría un saldo sin caja detrás. El orden de magnitud del saldo era unos $444.000 (290 botellas vendidas después del pago × ≈$1.531), unos $2.643 por botella de Malbec. Es una estimación.
  - "Pago depósito": botellas que se lleva el dueño del depósito de R Peña (Pelado) como pago. Teóricamente una caja por mes; lo confirma Pipi de vez en cuando, por eso se registra en lote cuando se hace un ajuste.
- El stock cuenta botellas por producto, no por DETALLE, así que renombrar categorías no cambia cantidades.
- **Elaboración por encargo.** `CLIENTES_POR_ENCARGO` (novatobackend.gs) lista clientes cuya venta no es comercial. En `getAnalytics` no entran en ventas por mes, por cliente, por producto ni en la concentración; sí cuentan en el flujo y en el margen de su añada (ahí están sus costos). Se muestran aparte en "Elaboración por encargo".
- **Blend 2023 sin etiquetar.** Las 3960 botellas están estibadas en Otero Ramos, sin etiqueta, no se venden todavía. En la app el producto lleva `sinEtiquetar: 'Otero Ramos'`: no aparece en ventas, transferencias ni conteos por depósito, y se muestra aparte en el Dashboard. Limitación: la tabla de depósitos de la hoja STOCK sigue poniendo esas botellas bajo R PEÑA (no hay columna Otero Ramos).
- **Organyca (ORG-001 a 004, $960.000).** Son ventas ordinarias posteriores al saldo del Chardonnay. Organyca quebró. La dolarización del 3-8-26 (US$676,06) es un acuerdo informal, no está registrada en la planilla. Criterio al cobrar: no cambiar el precio de las ventas. Se registra el cobro real en CAJA vinculado a ORG-001..004 y la diferencia con $960.000 va como fila "Ajuste" en BALANCE sin producto, para no tocar precio por botella ni stock.
