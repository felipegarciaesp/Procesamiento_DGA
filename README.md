# Procesamiento_DGA
Codigo para procesar los datos de estaciones DGA. Este codigo toma como inputs los outputs del codigo Datos_DGA

## ⚠️ Importante: variables a procesar y carpeta Outputs

Al inicio del script está el vector `variables_a_procesar`:

```r
variables_a_procesar <- c("PP", "TMEDIA", "TMAX", "TMIN")
```

**En este vector deben ir siempre TODAS las variables que se quieren procesar
en cada ejecución.** Por ejemplo, para procesar todo:

```r
variables_a_procesar <- c("PP", "TMEDIA", "TMAX", "TMIN")
```

Cómo funciona la carpeta `Outputs`:

- La carpeta **no se borra** al ejecutar el código; se crea solo si no existe.
- Cada variable tiene su **propia subcarpeta** dentro de `Outputs` con el nombre de
  la variable, y ahí se guardan sus dos archivos: `Outputs\PP\Resumen_PP.xlsx` y
  `Outputs\PP\Criterios_PP.xlsx`; igual para `TMEDIA`, `TMAX` y `TMIN`.
- Cada variable del vector genera o **sobrescribe** su `Resumen_<VARIABLE>.xlsx`.
  (`Criterios_<VARIABLE>.xlsx` nunca se sobrescribe, ver "Planilla de criterios").
- Los archivos de las variables que **no** están en el vector **no se tocan**:
  quedan tal como estaban de una ejecución anterior.

Por lo tanto:

- Para reprocesar solo una variable (por ejemplo, tras corregir algo de
  temperatura), basta con dejar solo esa variable en el vector; las demás
  conservan su archivo.
- **Si se cambia algo que afecta a todas las variables** (por ejemplo, el año
  hidrológico `mes_inicio_hidro` / `mes_fin_hidro`, o el formato de las tablas o
  gráficos), **se deben incluir todas las variables en el vector**. De lo
  contrario, en `Outputs` quedarán archivos generados con configuraciones
  distintas.
- Si un `Resumen_*.xlsx` está abierto en Excel, Windows no permite
  sobrescribirlo y el código se detiene con un mensaje: hay que cerrarlo y
  volver a ejecutar.
- Si una variable del vector no tiene planilla en `Inputs`, el código lo avisa
  al terminar.

## Estructura del código

El script `Procesamiento_DGA.R` está organizado en secciones:

1. **Configuración del usuario**: año hidrológico (`mes_inicio_hidro`, `mes_fin_hidro`)
   y variables a procesar (`variables_a_procesar`, ver sección "Importante").
2. **Configuración por variable** (`config_variables`): para cada variable se
   define cómo reconocer su planilla en `Inputs`, la columna de valores, el tipo
   de agregación (`"suma"` o `"promedio"`), la unidad y los rótulos de tablas y
   gráficos, y el tipo (`grafico_tipo`: barras o líneas) y color (`color_grafico`)
   de los gráficos. Para agregar una variable nueva basta con agregar un elemento
   a esta lista.
3. **Lectura**: cada planilla de `Inputs` se carga una sola vez.
4. **Cálculo**: `calcular_periodo()` calcula todas las tablas de un periodo
   anual. Se usa dos veces por estación: año calendario (mes de inicio ENE) y año
   hidrológico (mes de inicio configurado, con sus dos semestres).
5. **Armado de tablas y disposición**: todas las tablas se arman primero como
   data.frames y luego se calcula automáticamente dónde va cada una (ver
   "Disposición automática de las tablas").
6. **Escritura en Excel** y **gráficos**, en funciones separadas.

Se genera un archivo `Resumen_<VARIABLE>.xlsx` por planilla procesada
(por ejemplo, `Resumen_PP.xlsx`). Ver la sección "Importante" al inicio de
este documento.

### Disposición automática de las tablas

Las tablas y gráficos de cada pestaña se ubican automáticamente de izquierda a
derecha, uno a continuación del otro, según el ancho real de cada tabla. No hay
letras de columna fijas en el código: si se agrega, quita o ensancha una tabla,
las demás se desplazan solas y los gráficos siguen apuntando a las celdas
correctas.

- El orden de los bloques se define en `disponer_hoja()` (información de la
  estación, serie original, bloque año calendario, bloque año hidrológico, serie
  continua y gráfico diario).
- Dentro de cada bloque de periodo, el orden se define en `disponer_periodo()`.
- Las separaciones y tamaños (columnas entre tablas, ancho y alto de los
  gráficos) se configuran en la sección 3 del script (`separacion_tablas`,
  `separacion_tras_serie`, `ancho_grafico`, `alto_grafico`, etc.).

Con la configuración actual, la disposición resultante es la misma de siempre:
el bloque año calendario parte en la columna K, el año hidrológico en BK, la
serie continua en DO y el gráfico diario en DW.

### Gráficos por variable

| Variable | Gráficos | Color |
|---|---|---|
| `PP` | Barras (relleno y borde) | Celeste (Light Blue) |
| `TMAX` | Líneas con marcadores (mensual y anual); línea simple (diario) | Rojo |
| `TMIN` | Líneas con marcadores (mensual y anual); línea simple (diario) | Azul |
| `TMEDIA` | Líneas con marcadores (mensual y anual); línea simple (diario) | Verde |

Los marcadores tienen el mismo color de relleno y borde que la línea. Se aplica a los gráficos de promedio mensual, anuales y de la serie diaria. En los gráficos de líneas, cada `#N/A` se muestra como un salto (sin línea que una los puntos vecinos): el código activa en cada gráfico la opción de Excel "Mostrar #N/A como celda vacía", disponible en Excel 365 y 2019 en adelante. En versiones antiguas de Excel o en LibreOffice se unirían los puntos. Para
cambiarlo, editar `grafico_tipo` y `color_grafico` en `config_variables`.

### Agregación según la variable

| Cálculo | `"suma"` (PP) | `"promedio"` (temperaturas) |
|---|---|---|
| Valor mensual a partir de los datos diarios | Suma | Promedio (no el máximo ni mínimo del mes) |
| Valor anual y semestral a partir de los meses | Suma | Promedio |
| "Promedio anual" a partir de los promedios mensuales | Suma | Promedio |
| Promedio mensual y promedio de las series anuales entre años | Promedio | Promedio |

En las temperaturas, el valor anual es el promedio de los valores mensuales del año
y el "promedio anual" es el promedio de todos los años.

Los valores sin datos se escriben en Excel como `#N/A` (valor de error nativo,
sin fórmulas).

## Pestaña "Resumen"

Es la **primera pestaña** de cada `Resumen_<VARIABLE>.xlsx` y da una vista global de la
disponibilidad de datos de todas las estaciones. Se arma con lo ya calculado (no se
vuelve a leer ninguna pestaña). De izquierda a derecha:

1. **Tabla de estaciones**: ID (`E1`, `E2`… = nombre de su pestaña), nombre, código,
   coordenadas UTM, altitud y periodo de registro (`mes/año - mes/año`).
2. **Días con dato por año**, agrupados por década, con escala de color
   (rojo = 0 días, amarillo = 183, verde = 366). Abajo, la misma tabla en **año hidrológico**.
3. **Meses con dato por año**, también por década, con semáforo y su misma versión
   en año hidrológico justo debajo. La celda queda vacía si no hay datos.
   La leyenda (al final) indica:
   - 🟢 12 meses: año completo
   - 🟡 ≥ 7 meses con información
   - 🔴 < 7 meses con información
   - vacío: sin información

Notas:
- Los años parten en `anio_inicio_resumen` (por defecto 1960, sección 1 del código);
  si alguna estación tiene datos anteriores, la tabla se extiende hacia atrás.
- El último año de las tablas lo define `anio_fin_resumen` (sección 1 del código, inicialmente **2026**).
  Solo fija hasta qué año se arman estas tablas; **no filtra datos**. Debe actualizarse a mano
  (2027, 2028…) cuando se agreguen datos más recientes. Si alguna estación tiene datos posteriores,
  la tabla se extiende sola. En año hidrológico cada columna se ubica bajo el año en que termina
  (1959/60 bajo 1960); con 2026 el último es 2026/27.
- Las tablas se ubican automáticamente según la cantidad de estaciones y años.

## Planilla de criterios

Junto a cada `Resumen_<VARIABLE>.xlsx` se genera `Criterios_<VARIABLE>.xlsx`
(por ejemplo, `Criterios_PP.xlsx`) con una pestaña por estación, igual a las de
Resumen (si una estación no aparece en Resumen, tampoco en Criterios). El
usuario la completa tras revisar los resultados y un código posterior (aún sin
nombre) la usará junto con Resumen para generar la planilla filtrada.

**Importante: el archivo no se sobrescribe.** Si `Criterios_PP.xlsx` ya existe,
el código lo deja tal cual para no perder lo que el usuario ya ingresó. Para
regenerarlo (por ejemplo, si cambian las estaciones o el formato) hay que
borrar el archivo y volver a ejecutar.

| Columna | Encabezado | Contenido |
|---|---|---|
| A, B | Año inicio / Año fin | Periodo de análisis (años calendario, inclusive). Vacío = toda la serie. |
| D | Dias minimos | Un número. Días con dato que debe tener un mes para ser válido; si tiene menos, se descartan todos sus datos. |
| E | Meses a aplicar | Números de mes (1-12), uno por fila, donde aplica "Dias minimos". Vacío = no se aplica el filtro. |
| F | Meses minimos para anual | Un número. Mínimo de meses válidos para calcular el valor anual; si no se alcanza, el anual queda `#N/A`. Vacío = sin mínimo (se ignoran los `#N/A`). |
| G | Meses obligatorios para anual | Números de mes (1-12), uno por fila, que deben tener valor para calcular el anual; si falta alguno, queda `#N/A`. Vacío = da lo mismo cuáles meses estén completos. |
| I | Fechas a eliminar | Fechas puntuales (dd/mm/yyyy), una por fila. |
| J, K | Rango: inicio / Rango: fin | Periodo a eliminar (inclusive). |
| L | Comentario | Motivo de la eliminación, para mantener la trazabilidad. |

Cómo completarla:

- Los datos de las columnas D a G son criterios de la estación: "Dias minimos"
  y "Meses minimos para anual" van en la fila 2; los meses, uno por fila hacia
  abajo.
- En I:L cada fila es una eliminación: una fecha puntual (I) o un rango (J y K),
  con su comentario en L.
- F y G pueden usarse juntas: el valor anual se calcula solo si se cumplen
  ambas condiciones. G puede quedar vacía sin problema.

Reglas para el código que use esta planilla:

1. **Orden del filtro:** primero se eliminan fechas puntuales y rangos, y
   después se cuentan los días del mes (si se eliminan 12 datos de un mes y
   quedan 18 días, el mes se descarta con un mínimo de 20).
2. Un dato eliminado queda como faltante (`#N/A`), no como 0.
3. Los criterios del valor anual (F y G) se aplican igual al año calendario y
   al hidrológico, sobre los meses que queden válidos tras el paso 1.
4. **Periodo de análisis y año hidrológico:** al indicar los años X e Y se
   entiende que se quiere analizar desde enero de X hasta diciembre de Y. Para
   el año hidrológico se toman entonces todos los años hidrológicos que
   contienen esos meses: el primero es X-1/X y el último Y/Y+1. Ejemplo con
   ABR-MAR: el periodo 1980-2026 se analiza como 1979/80 a 2026/27 (enero de
   1980 pertenece al año hidrológico 1979/80 y diciembre de 2026 al 2026/27).
   Si el año hidrológico comenzara en enero, no habría desfase.

## Supuestos de funcionamiento

### Serie diaria continua

En cada pestaña de estación, las columnas DO:DU (con la disposición actual) contienen una serie diaria
continua desde la fecha mínima hasta la máxima de la serie original de D:F,
incluyendo las fechas sin valor. Los encabezados son FECHA, DIA, MES, YEAR,
Hydro YEAR, `<variable> diaria` (por ejemplo, Pp diaria) y Unidad.

- DIA, MES, YEAR y Hydro YEAR se calculan en R a partir de FECHA y se escriben
  como valores. Hydro YEAR sigue el mes inicial configurado en el script y usa el
  formato `1999/00`.
- La columna de la variable contiene una copia del valor original de la columna
  E. Cuando no existe un registro o su valor es `NA`, contiene `#N/A`; si falta
  el registro, la unidad queda vacía.
- Si existen fechas duplicadas en la serie original, se usa el primer registro.
- Al modificar las fechas o valores originales es necesario volver a ejecutar el
  script.

### Rango de años de las tablas de conteo

Las tablas "Cantidad de registros por mes" y "Cantidad de meses con información" (año calendario e hidrológico) se construyen a partir de la serie diaria de cada estación.

**Cómo se define el vector de años:**

- Se consideran solo las fechas con un valor **distinto de `NA`**. Las celdas vacías en la planilla de entrada se leen como `NA`, por lo que se tratan igual.
- El vector de años va desde el **primer** hasta el **último** año con al menos un valor, es decir, `seq(min(anios), max(anios))`.
- El vector es **continuo**: incluye los años intermedios sin registros, que aparecen con `0` en todos los meses de la tabla de conteo y con `0` en la cantidad de meses con información.
- Ambas tablas usan el mismo vector de años, por lo que tienen la misma cantidad de filas.

**Ejemplo:** si la planilla de una estación tiene fechas desde 1980 hasta 2020, pero:

- En 1980 y 1981 la precipitación está vacía (`NA`).
- Los primeros valores reales aparecen en 1982.
- Los últimos valores reales terminan en 2018, y 2019-2020 vienen vacíos.

Las tablas se construyen desde **1982 hasta 2018**.

**Consideraciones:**

- Si la planilla de entrada solo trae filas cuando hay dato, el rango de años coincide con el período completo de las fechas de la planilla.
- Los años del inicio o del final de la serie que solo tengan valores vacíos no aparecen en las tablas.
- Si una estación no tiene ningún valor válido, el código se detiene con el mensaje "La serie no tiene valores validos."
