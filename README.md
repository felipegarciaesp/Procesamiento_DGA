# Procesamiento_DGA
Codigo para procesar los datos de estaciones DGA. Este codigo toma como inputs los outputs del codigo Datos_DGA

## ⚠️ Importante: variables a procesar y carpeta Outputs

Al inicio del script está el vector `variables_a_procesar`:

```r
variables_a_procesar <- c("PP")
```

**En este vector deben ir siempre TODAS las variables que se quieren procesar
en cada ejecución.** Por ejemplo, para procesar todo:

```r
variables_a_procesar <- c("PP", "TMEDIA", "TMAX", "TMIN")
```

Cómo funciona la carpeta `Outputs`:

- La carpeta **no se borra** al ejecutar el código; se crea solo si no existe.
- Cada variable del vector genera o **sobrescribe** su propio archivo:
  `"PP"` → `Resumen_PP.xlsx`, `"TMEDIA"` → `Resumen_TMEDIA.xlsx`, etc.
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
   gráficos. Para agregar una variable nueva basta con agregar un elemento a esta
   lista.
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

### Agregación según la variable

| Cálculo | `"suma"` (PP) | `"promedio"` (temperaturas) |
|---|---|---|
| Valor mensual a partir de los datos diarios | Suma | Promedio |
| Valor anual y semestral a partir de los meses | Suma | Promedio |
| "Promedio anual" a partir de los promedios mensuales | Suma | Promedio |
| Promedio mensual y promedio de las series anuales entre años | Promedio | Promedio |

Los valores sin datos se escriben en Excel como `#N/A` (valor de error nativo,
sin fórmulas).

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
