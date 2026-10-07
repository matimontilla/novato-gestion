# Decisiones y criterios — Novato Gestión

Memoria del proyecto: **por qué** las cosas están como están. El código dice *qué*
hace; esto dice *por qué*. Si algo acá parece raro, probablemente hay una razón
dolorosa detrás.

Última actualización: julio 2026.

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
- `herramientas.gs`: diagnósticos, reparaciones reutilizables y configuración, para
  correr a mano con ▶ Run. El desplegable del editor muestra sólo las funciones del
  archivo abierto.

Las funciones de una sola vez que ya se corrieron (reclasificación de costos, columnas
DEPOSITO y DIAS CIERRE, migración de zona horaria, etc.) **se borraron en octubre 2026**.
Este documento las sigue nombrando como antecedente; el código está en el historial de
GitHub (anterior al commit que crea `herramientas.gs`).

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
- Comunicación del proyecto en español
- Productos con stock 0 se ocultan de los selectores
- Signos: cobro positivo, gasto negativo. Verde `+` / rojo `-`. Sobrepago se muestra
  distinto (azul, "a favor") de deuda real.
