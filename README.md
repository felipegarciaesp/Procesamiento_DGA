# Procesamiento_DGA
Codigo para procesar los datos de estaciones DGA. Este codigo toma como inputs los outputs del codigo Datos_DGA

## Supuestos de funcionamiento

### Rango de años de las tablas de conteo

Las tablas "Cantidad de registros por mes - Año calendario" y "Cantidad de meses con información" se construyen a partir de la serie diaria de precipitación de cada estación.

**Cómo se define el vector de años:**

- Se consideran solo las fechas con un valor de precipitación **distinto de `NA`**. Las celdas vacías en la planilla de entrada se leen como `NA`, por lo que se tratan igual.
- El vector de años va desde el **primer** hasta el **último** año con al menos un valor de precipitación, es decir, `seq(min(anios), max(anios))`.
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
- Si una estación no tiene ningún valor de precipitación, el cálculo del rango falla, ya que el código no contempla ese caso.
