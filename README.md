# Procesamiento_DGA
Codigo para procesar los datos de estaciones DGA. Este codigo toma como inputs los outputs del codigo Datos_DGA

## Estructura del código

El script `Procesamiento_DGA.R` está organizado en secciones:

1. **Configuración del usuario**: año hidrológico (`mes_inicio_hidro`, `mes_fin_hidro`)
   y variables a procesar (`variables_a_procesar`).
2. **Configuración por variable** (`config_variables`): para cada variable se
   define cómo reconocer su planilla en `Inputs`, la columna de valores, el tipo
   de agregación (`"suma"` o `"promedio"`), la unidad y los rótulos de tablas y
   gráficos. Para agregar una variable nueva basta con agregar un elemento a esta
   lista.
3. **Lectura**: cada planilla de `Inputs` se carga una sola vez.
4. **Cálculo**: `calcular_periodo()` calcula todas las tablas de un periodo
   anual. Se usa dos veces por estación: año calendario (mes de inicio ENE) y año
   hidrológico (mes de inicio configurado, con sus dos semestres).
5. **Escritura en Excel** y **gráficos**, en funciones separadas.

Se genera un archivo `Resumen_<VARIABLE>.xlsx` por planilla procesada
(por ejemplo, `Resumen_PP.xlsx`).

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

En cada pestaña de estación, las columnas DO:DU contienen una serie diaria
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
