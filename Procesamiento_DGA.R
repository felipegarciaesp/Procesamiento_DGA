# =====================================================================
# Codigo para procesar los datos de estaciones DGA / Felipe Garcia

# Este codigo toma como inputs los outputs del codigo Datos_DGA
# =====================================================================

## Limpiar ambiente
#-------------------
rm(list=ls())
graphics.off()
cat("\014")

library(openxlsx2)
library(mschart)

# =====================================================================
# 1. CONFIGURACION DEL USUARIO
# =====================================================================

## Configuracion del año hidrologico
#-------------------
# Los meses se expresan como numero: 1 = ENE, 2 = FEB, 3 = MAR, 4 = ABR, ..., 12 = DIC
# El año hidrologico va desde el mes inicial del año X hasta el mes final del año X+1.
# Por defecto: 4 (ABR) del año X hasta 3 (MAR) del año X+1  ->  ejemplo: 1999/00
# Los semestres se definen a partir del mes inicial: con ABR-MAR se analizan
# ABR-SEP y OCT-MAR; con MAY-ABR se analizan MAY-OCT y NOV-ABR.
mes_inicio_hidro <- 4   # 4 = ABR (mes inicial, año X)
mes_fin_hidro    <- 3   # 3 = MAR (mes final, año X+1)

# Validacion: el mes final debe ser el mes anterior al mes inicial (12 meses en total)
if (mes_fin_hidro != ((mes_inicio_hidro + 10) %% 12) + 1) {
  stop("mes_fin_hidro debe ser el mes anterior a mes_inicio_hidro (el año hidrologico debe tener 12 meses).")
}

## Variables a procesar
#-------------------
# Codigos definidos en config_variables (seccion 2). Solo se procesan las
# planillas de Inputs cuya variable este en este vector.
# Ejemplo: c("PP", "TMEDIA", "TMAX", "TMIN")
variables_a_procesar <- c("PP")

## Directorio de trabajo
#-------------------
setwd("C:/Codigos/Procesamiento_DGA")
dir_inputs  <- file.path(getwd(), "Inputs")
dir_outputs <- file.path(getwd(), "Outputs")

## Carpeta de salida (se recrea en cada ejecucion)
#-------------------
if (dir.exists(dir_outputs)) unlink(dir_outputs, recursive = TRUE)  # borra resultados anteriores
dir.create(dir_outputs, recursive = TRUE)

# =====================================================================
# 2. CONFIGURACION POR VARIABLE
# =====================================================================
# Para agregar una variable nueva basta con agregar un elemento a esta lista.
#
#   patron_archivo           : expresion regular para reconocer la planilla de
#                              Inputs por su nombre de archivo (en minusculas)
#   patron_columna           : expresion regular para encontrar la columna de
#                              valores dentro de cada pestaña
#   agregacion               : "suma" o "promedio". Define como se pasa de
#                              diario a mensual, de mensual a anual/semestral, y
#                              como se obtiene el "Promedio anual" a partir de
#                              los promedios mensuales
#   unidad                   : unidad que se muestra en titulos y encabezados
#   nombre_tabla             : nombre en los titulos de las tablas
#   abrev                    : abreviatura en los encabezados de las columnas
#   nombre_grafico           : nombre en el eje Y de los graficos anuales y diario
#   grafico_promedio_mensual : rotulo del eje Y del grafico de promedios mensuales
#   titulo_diario            : titulo del grafico de la serie diaria
#   grafico_diario           : "barras" o "lineas"
config_variables <- list(

  PP = list(patron_archivo           = "^pp|precip",
            patron_columna           = "precip|^pp|valor",
            agregacion               = "suma",
            unidad                   = "mm",
            nombre_tabla             = "Precipitacion",
            abrev                    = "Pp",
            nombre_grafico           = "Precipitación",
            grafico_promedio_mensual = "Precipitación media",
            titulo_diario            = "Precipitaciones diarias",
            grafico_diario           = "barras"),

  TMEDIA = list(patron_archivo           = "tmed",
                patron_columna           = "temp|valor",
                agregacion               = "promedio",
                unidad                   = "°C",
                nombre_tabla             = "Temperatura media",
                abrev                    = "Tmedia",
                nombre_grafico           = "Temperatura media",
                grafico_promedio_mensual = "Temperatura media mensual",
                titulo_diario            = "Temperatura media diaria",
                grafico_diario           = "lineas"),

  TMAX = list(patron_archivo           = "tmax",
              patron_columna           = "temp|valor",
              agregacion               = "promedio",
              unidad                   = "°C",
              nombre_tabla             = "Temperatura máxima",
              abrev                    = "Tmax",
              nombre_grafico           = "Temperatura máxima",
              grafico_promedio_mensual = "Temperatura máxima media mensual",
              titulo_diario            = "Temperatura máxima diaria",
              grafico_diario           = "lineas"),

  TMIN = list(patron_archivo           = "tmin",
              patron_columna           = "temp|valor",
              agregacion               = "promedio",
              unidad                   = "°C",
              nombre_tabla             = "Temperatura mínima",
              abrev                    = "Tmin",
              nombre_grafico           = "Temperatura mínima",
              grafico_promedio_mensual = "Temperatura mínima media mensual",
              titulo_diario            = "Temperatura mínima diaria",
              grafico_diario           = "lineas")

  # Ejemplo para caudales (cuando exista la planilla de entrada):
  # , Q = list(patron_archivo = "^q|caudal", patron_columna = "caudal|valor",
  #            agregacion = "promedio", unidad = "m3/s",
  #            nombre_tabla = "Caudal", abrev = "Q", nombre_grafico = "Caudal",
  #            grafico_promedio_mensual = "Caudal medio mensual",
  #            titulo_diario = "Caudales diarios", grafico_diario = "lineas")
)

# Validacion de la configuracion
for (cod in names(config_variables)) {
  cfg <- config_variables[[cod]]
  if (!(cfg$agregacion %in% c("suma", "promedio"))) {
    stop("La variable ", cod, " tiene una agregacion no valida: ", cfg$agregacion)
  }
  if (!(cfg$grafico_diario %in% c("barras", "lineas"))) {
    stop("La variable ", cod, " tiene un grafico_diario no valido: ", cfg$grafico_diario)
  }
}
variables_desconocidas <- setdiff(variables_a_procesar, names(config_variables))
if (length(variables_desconocidas) > 0) {
  stop("Variables sin configuracion: ", paste(variables_desconocidas, collapse = ", "))
}

# =====================================================================
# 3. CONSTANTES
# =====================================================================
nombres_meses <- c("ENE", "FEB", "MAR", "ABR", "MAY", "JUN",
                   "JUL", "AGO", "SEP", "OCT", "NOV", "DIC")

# Rotulos y columnas de origen para la tabla de informacion de la estacion
info_rotulos  <- c("NOMBRE ESTACION", "CÓDIGO ESTACION", "REGIÓN",
                   "UTM ESTE (m)", "UTM NORTE (m)", "ALTITUD (m s.n.m.)")
info_columnas <- c("NOMBRE.ESTACION", "CODIGO.ESTACION", "REGION",
                   "UTM_ESTE", "UTM_NORTE", "ALTITUD")

# Columnas de inicio de cada bloque en las pestañas de salida
col_bloque_calendario <- 11    # K
col_bloque_hidro      <- 63    # BK
col_serie_continua    <- 119   # DO

# Formatos de texto (Calibri) y grilla de los graficos
txt_titulo   <- officer::fp_text(font.size = 14, font.family = "Calibri")
txt_rotulo   <- officer::fp_text(font.size = 10, font.family = "Calibri")
txt_ejes     <- officer::fp_text(font.size = 9,  font.family = "Calibri")
sin_linea    <- officer::fp_border(style = "none")
linea_grilla <- officer::fp_border(color = "#D9D9D9", width = 0.75)

# =====================================================================
# 4. FUNCIONES AUXILIARES
# =====================================================================

# Identifica la variable de una planilla segun su nombre de archivo.
# Devuelve el codigo de config_variables (ej: "PP") o NA.
identificar_variable <- function(nombre_archivo) {
  nombre <- tolower(nombre_archivo)
  for (cod in names(config_variables)) {
    if (grepl(config_variables[[cod]]$patron_archivo, nombre)) return(cod)
  }
  NA_character_
}

# Busca una columna por patron (ignora mayusculas, tildes y separadores)
buscar_columna <- function(df, patron, descripcion) {
  nombres_norm <- tolower(iconv(names(df), to = "ASCII//TRANSLIT"))
  idx <- grep(patron, nombres_norm)
  if (length(idx) == 0) {
    stop("No se encontro la columna de ", descripcion, ". Columnas disponibles: ",
         paste(names(df), collapse = ", "))
  }
  names(df)[idx[1]]
}

# Agrega un vector segun el tipo ("suma" o "promedio"), ignorando NA.
# Si todos los valores son NA, devuelve NA.
agregar <- function(v, tipo) {
  v <- v[!is.na(v)]
  if (length(v) == 0) return(NA_real_)
  if (tipo == "suma") sum(v) else mean(v)
}

# =====================================================================
# 5. LECTURA
# =====================================================================

# Lee una pestaña de estacion desde el libro ya cargado en memoria.
# Devuelve NULL si la pestaña no tiene datos.
leer_estacion <- function(wb_in, hoja, cfg) {

  # check_names = TRUE: los espacios de los rotulos pasan a "." (como en openxlsx)
  df <- wb_to_df(wb_in, sheet = hoja, detect_dates = TRUE, check_names = TRUE)
  if (is.null(df) || nrow(df) == 0) return(NULL)

  col_fecha <- buscar_columna(df, "fecha", "FECHA")
  col_valor <- buscar_columna(df, cfg$patron_columna, "valores")
  col_unid  <- buscar_columna(df, "unidad", "unidad")

  list(info         = unname(sapply(info_columnas, function(col) as.character(df[[col]][1]))),
       nombre_valor = col_valor,
       serie        = data.frame(FECHA  = as.Date(df[[col_fecha]]),
                                 VALOR  = df[[col_valor]],
                                 Unidad = df[[col_unid]],
                                 stringsAsFactors = FALSE))
}

# =====================================================================
# 6. CALCULO (sin Excel)
# =====================================================================

# Calcula todas las tablas de un periodo anual a partir de la serie diaria.
# Sirve para el año calendario (mes_inicio = 1, hidrologico = FALSE) y para
# el año hidrologico (mes_inicio = mes_inicio_hidro, hidrologico = TRUE).
#
# Solo el año hidrologico incluye los dos semestres, cuyos nombres dependen
# del mes de inicio (ej: ABR-SEP y OCT-MAR, o MAY-OCT y NOV-ABR).
calcular_periodo <- function(fechas, valores, mes_inicio, agregacion, hidrologico) {

  validos <- !is.na(fechas) & !is.na(valores)
  if (!any(validos)) stop("La serie no tiene valores validos.")
  f <- fechas[validos]
  v <- valores[validos]

  anio <- as.integer(format(f, "%Y"))
  mes  <- as.integer(format(f, "%m"))

  # Año del periodo de cada dato: inicia en X si mes >= mes de inicio, si no en X-1
  anio_per <- anio - as.integer(mes < mes_inicio)
  pos_mes  <- ((mes - mes_inicio) %% 12) + 1            # posicion del mes en el año
  orden    <- ((mes_inicio - 1 + 0:11) %% 12) + 1       # meses calendario en orden del periodo
  meses    <- nombres_meses[orden]

  # Vector continuo de años: desde el primer hasta el ultimo año con datos
  anios <- seq(min(anio_per), max(anio_per))
  n     <- length(anios)
  etiquetas <- if (hidrologico) sprintf("%d/%02d", anios, (anios + 1) %% 100) else anios

  f_anio <- factor(anio_per, levels = anios)
  f_mes  <- factor(pos_mes,  levels = 1:12)

  # Cantidad de registros por mes y meses con informacion
  conteo   <- matrix(as.integer(table(f_anio, f_mes)), nrow = n, ncol = 12)
  n_meses  <- rowSums(conteo > 0)

  # Valores mensuales (suma o promedio de los datos diarios); NA si no hay datos
  mensual <- tapply(v, list(f_anio, f_mes), agregar, tipo = agregacion)
  mensual <- matrix(as.numeric(mensual), nrow = n, ncol = 12)

  # Series anuales: año completo y, en el año hidrologico, los dos semestres
  grupos_meses   <- list(1:12)
  nombres_series <- "anual"
  nombres_prom   <- "promedio anual"
  nombres_resum  <- "Promedio anual"
  if (hidrologico) {
    lab_s1 <- paste0(meses[1], "-", meses[6])
    lab_s2 <- paste0(meses[7], "-", meses[12])
    grupos_meses   <- c(grupos_meses, list(1:6, 7:12))
    nombres_series <- c(nombres_series, paste("anual", lab_s1), paste("anual", lab_s2))
    nombres_prom   <- c(nombres_prom, paste("promedio", lab_s1), paste("promedio", lab_s2))
    nombres_resum  <- c(nombres_resum, paste("Promedio", lab_s1), paste("Promedio", lab_s2))
  }

  anual <- sapply(grupos_meses, function(idx) {
    apply(mensual[, idx, drop = FALSE], 1, agregar, tipo = agregacion)
  })
  anual <- matrix(anual, nrow = n)

  # Promedio de cada serie anual (siempre promedio entre años)
  prom_anual <- apply(anual, 2, agregar, tipo = "promedio")

  # Promedio mensual entre años y su resumen anual/semestral
  prom_mensual <- apply(mensual, 2, agregar, tipo = "promedio")
  resumen_prom <- sapply(grupos_meses, function(idx) agregar(prom_mensual[idx], agregacion))

  # Serie mensual en vector (una fila por mes)
  mes_vec     <- rep(orden, times = n)
  anio_vec    <- rep(anios, each = 12)
  anio_cal    <- anio_vec + as.integer(mes_vec < mes_inicio)
  vector_mens <- data.frame(FECHA = as.Date(sprintf("%d-%02d-01", anio_cal, mes_vec)),
                            MES   = mes_vec,
                            YEAR  = rep(etiquetas, each = 12),
                            VALOR = as.numeric(t(mensual)))

  list(hidrologico    = hidrologico,
       etiquetas      = etiquetas,
       meses          = meses,
       conteo         = conteo,
       n_meses        = n_meses,
       mensual        = mensual,
       vector_mens    = vector_mens,
       prom_mensual   = prom_mensual,
       resumen_prom   = resumen_prom,
       nombres_resum  = nombres_resum,
       anual          = anual,
       nombres_series = nombres_series,
       prom_anual     = prom_anual,
       nombres_prom   = nombres_prom)
}

# =====================================================================
# 7. ESCRITURA EN EXCEL
# =====================================================================
# Los valores NA se escriben como #N/A (valor de error nativo de Excel,
# sin formulas) mediante na = na_strings().

# Escribe un titulo con celdas unidas, en negrita y centrado (fila 1)
escribir_titulo <- function(wb, hoja, texto, col_ini, col_fin) {
  wb$add_data(sheet = hoja, x = texto, start_col = col_ini, start_row = 1, col_names = FALSE)
  wb$merge_cells(sheet = hoja, dims = wb_dims(rows = 1, cols = col_ini:col_fin))
  wb$add_font(sheet = hoja, dims = wb_dims(rows = 1, cols = col_ini), bold = TRUE)
  wb$add_cell_style(sheet = hoja, dims = wb_dims(rows = 1, cols = col_ini), horizontal = "center")
}

# Escribe una tabla con encabezado en negrita en la fila 2
escribir_tabla <- function(wb, hoja, df, col_ini) {
  wb$add_data(sheet = hoja, x = df, start_col = col_ini, start_row = 2, na = na_strings())
  wb$add_font(sheet = hoja, dims = wb_dims(rows = 2, cols = col_ini:(col_ini + ncol(df) - 1)),
              bold = TRUE)
}

# Tabla de informacion de la estacion (A1:B7)
escribir_info <- function(wb, hoja, info) {
  escribir_titulo(wb, hoja, "Informacion de la Estación", 1, 2)
  wb$add_data(sheet = hoja, x = info_rotulos, start_col = 1, start_row = 2, col_names = FALSE)
  wb$add_data(sheet = hoja, x = info,         start_col = 2, start_row = 2, col_names = FALSE)
}

# Serie diaria original (D1:F...)
escribir_serie_original <- function(wb, hoja, serie, nombre_valor) {
  tabla <- serie
  names(tabla) <- c("FECHA", nombre_valor, "Unidad")
  wb$add_data(sheet = hoja, x = tabla, start_col = 4, start_row = 1)
  wb$add_font(sheet = hoja, dims = wb_dims(rows = 1, cols = 4:6), bold = TRUE)
  wb$add_numfmt(sheet = hoja, dims = wb_dims(rows = 2:(nrow(tabla) + 1), cols = 4),
                numfmt = "dd/mm/yyyy")
  wb$set_col_widths(sheet = hoja, cols = 1:6, widths = c(22, 28, 3, 12, 20, 10))
}

# Serie diaria continua (DO1:DU...): todas las fechas entre la minima y la
# maxima de la serie original. DIA, MES, YEAR y Hydro YEAR se escriben como
# valores (ya no como formulas). Devuelve la ultima fila escrita.
escribir_serie_continua <- function(wb, hoja, serie, cfg, mes_inicio_hidro,
                                    col_ini = col_serie_continua) {
  fechas_validas <- serie$FECHA[!is.na(serie$FECHA)]
  if (length(fechas_validas) == 0) {
    stop("La pestana ", hoja, " no contiene fechas validas.")
  }

  fecha_ini <- min(fechas_validas)
  fecha_fin <- max(fechas_validas)
  n_dias <- as.integer(fecha_fin - fecha_ini) + 1L
  if (n_dias > 1048575L) {
    stop("La serie continua de la pestana ", hoja, " supera el limite de filas de Excel.")
  }

  fechas <- seq(fecha_ini, fecha_fin, by = "day")
  origen <- match(fechas, serie$FECHA)   # si hay fechas duplicadas se usa el primer registro
  mes    <- as.integer(format(fechas, "%m"))
  anio   <- as.integer(format(fechas, "%Y"))
  anio_h <- anio - as.integer(mes < mes_inicio_hidro)

  unidades <- serie$Unidad[origen]
  unidades[is.na(origen)] <- ""

  tabla <- data.frame(FECHA  = fechas,
                      DIA    = as.integer(format(fechas, "%d")),
                      MES    = mes,
                      YEAR   = anio,
                      HYDRO  = sprintf("%d/%02d", anio_h, (anio_h + 1L) %% 100L),
                      VALOR  = serie$VALOR[origen],
                      Unidad = unidades,
                      stringsAsFactors = FALSE)
  names(tabla) <- c("FECHA", "DIA", "MES", "YEAR", "Hydro YEAR",
                    paste(cfg$abrev, "diaria"), "Unidad")

  wb$add_data(sheet = hoja, x = tabla, start_col = col_ini, start_row = 1, na = na_strings())
  wb$add_font(sheet = hoja, dims = wb_dims(rows = 1, cols = col_ini:(col_ini + 6)), bold = TRUE)
  wb$add_numfmt(sheet = hoja, dims = wb_dims(rows = 2:(n_dias + 1L), cols = col_ini),
                numfmt = "dd/mm/yyyy")
  wb$set_col_widths(sheet = hoja, cols = col_ini:(col_ini + 6),
                    widths = c(12, 6, 6, 8, 12, 12, 10))

  n_dias + 1L
}

# Escribe las tablas de un periodo (calendario o hidrologico) a partir de
# col_ini. Devuelve las columnas de inicio de cada tabla (para los graficos).
escribir_periodo <- function(wb, hoja, res, cfg, col_ini) {

  u          <- cfg$unidad
  sufijo     <- if (res$hidrologico) " - Año hidrologico" else " - Año calendario"
  ancho_year <- if (res$hidrologico) 10 else 8
  n          <- length(res$etiquetas)
  k          <- length(res$nombres_series)   # 1 (anual) o 3 (anual + 2 semestres)

  pos <- list(conteo   = col_ini,
              meses    = col_ini + 14,
              mensual  = col_ini + 17,
              vector   = col_ini + 31,
              promedio = col_ini + 36,
              anual    = col_ini + 39)
  pos$graficos <- pos$anual + (1 + 2 * k) + 1

  # --- Cantidad de registros por mes ---
  tabla_conteo <- data.frame(res$etiquetas, res$conteo)
  names(tabla_conteo) <- c("YEAR", res$meses)

  escribir_titulo(wb, hoja, paste0("Cantidad de registros por mes", sufijo),
                  pos$conteo, pos$conteo + 12)
  escribir_tabla(wb, hoja, tabla_conteo, pos$conteo)
  wb$set_col_widths(sheet = hoja, cols = pos$conteo + 0:12, widths = c(ancho_year, rep(6, 12)))

  # Escala de colores: 0 = rojo, 15 = amarillo, 31 = verde oscuro
  wb$add_conditional_formatting(sheet = hoja,
                                dims = wb_dims(rows = 3:(n + 2), cols = pos$conteo + 1:12),
                                type = "colorScale",
                                style = c("#F8696B", "#FFEB84", "#006400"),
                                rule = c(0, 15, 31))

  # --- Cantidad de meses con informacion ---
  tabla_meses <- data.frame(YEAR = res$etiquetas, MESES = res$n_meses)

  escribir_titulo(wb, hoja, paste0("Cantidad de meses con información", sufijo),
                  pos$meses, pos$meses + 1)
  escribir_tabla(wb, hoja, tabla_meses, pos$meses)
  wb$set_col_widths(sheet = hoja, cols = pos$meses + 0:1, widths = c(ancho_year, 8))

  # --- Valores mensuales (tabla año x mes) ---
  tabla_mensual <- data.frame(res$etiquetas, res$mensual)
  names(tabla_mensual) <- c("YEAR", res$meses)

  escribir_titulo(wb, hoja, paste0(cfg$nombre_tabla, " mensual (", u, ")", sufijo),
                  pos$mensual, pos$mensual + 12)
  escribir_tabla(wb, hoja, tabla_mensual, pos$mensual)
  wb$set_col_widths(sheet = hoja, cols = pos$mensual + 0:12, widths = c(ancho_year, rep(8, 12)))

  # --- Valores mensuales en vector ---
  tabla_vector <- res$vector_mens
  names(tabla_vector) <- c("FECHA", "MES", "YEAR", paste0(cfg$abrev, " (", u, ")"))

  escribir_titulo(wb, hoja, paste0(cfg$abrev, " mensual (", u, ")", sufijo),
                  pos$vector, pos$vector + 3)
  escribir_tabla(wb, hoja, tabla_vector, pos$vector)
  wb$add_numfmt(sheet = hoja, dims = wb_dims(rows = 3:(nrow(tabla_vector) + 2), cols = pos$vector),
                numfmt = "dd/mm/yyyy")
  wb$set_col_widths(sheet = hoja, cols = pos$vector + 0:3, widths = c(12, 6, ancho_year, 10))

  # --- Promedio mensual entre años + resumen anual/semestral (fila 15 en adelante) ---
  tabla_prom <- data.frame(MES = res$meses, VALOR = res$prom_mensual)
  names(tabla_prom) <- c("MES", paste0(cfg$abrev, " promedio (", u, ")"))

  escribir_titulo(wb, hoja, paste0(cfg$nombre_tabla, " promedio mensual (", u, ")", sufijo),
                  pos$promedio, pos$promedio + 1)
  escribir_tabla(wb, hoja, tabla_prom, pos$promedio)

  wb$add_data(sheet = hoja, x = paste0(res$nombres_resum, " (", u, ")"),
              start_col = pos$promedio, start_row = 15, col_names = FALSE)
  wb$add_data(sheet = hoja, x = res$resumen_prom,
              start_col = pos$promedio + 1, start_row = 15, col_names = FALSE, na = na_strings())
  wb$add_font(sheet = hoja, dims = wb_dims(rows = 15:(14 + k), cols = pos$promedio + 0:1),
              bold = TRUE)
  wb$set_col_widths(sheet = hoja, cols = pos$promedio + 0:1,
                    widths = c(if (k > 1) 28 else 20, 16))

  # --- Series anuales y sus promedios ---
  tabla_anual <- data.frame(res$etiquetas, res$anual)
  names(tabla_anual) <- c("YEAR", paste0(cfg$abrev, " ", res$nombres_series, " (", u, ")"))

  escribir_titulo(wb, hoja, paste0(cfg$abrev, " anual y promedio anual (", u, ")", sufijo),
                  pos$anual, pos$anual + 2 * k)
  escribir_tabla(wb, hoja, tabla_anual, pos$anual)

  # Promedios: encabezado (fila 2) y un unico valor (fila 3) a la derecha de la tabla
  tabla_prom_anual <- as.data.frame(as.list(res$prom_anual))
  names(tabla_prom_anual) <- paste0(cfg$abrev, " ", res$nombres_prom, " (", u, ")")
  escribir_tabla(wb, hoja, tabla_prom_anual, pos$anual + 1 + k)

  wb$set_col_widths(sheet = hoja, cols = pos$anual + 0:(2 * k),
                    widths = c(ancho_year, 14, rep(20, k - 1), 22, rep(26, k - 1)))

  pos
}

# =====================================================================
# 8. GRAFICOS (nativos de Excel, enlazados a las tablas de la pestaña)
# =====================================================================

# Limites del eje Y: incluye siempre el 0 y redondea a valores "bonitos"
limites_eje <- function(v) {
  v <- v[!is.na(v)]
  if (length(v) == 0) return(c(0, 1))
  range(pretty(c(0, v)))
}

# Formato comun: sin leyenda, solo grilla horizontal mayor, sin titulo eje X
formatear_grafico <- function(g, titulo, ylab, limites) {
  g <- chart_labels(g, title = titulo, xlab = NULL, ylab = ylab)
  g <- chart_ax_y(g, limit_min = limites[1], limit_max = limites[2])
  g <- chart_theme(g,
                   main_title   = txt_titulo,
                   axis_title_y = txt_rotulo,
                   axis_text_x  = txt_ejes,
                   axis_text_y  = txt_ejes,
                   legend_position = "n",
                   grid_major_line_x = sin_linea,
                   grid_minor_line_x = sin_linea,
                   grid_minor_line_y = sin_linea,
                   grid_major_line_y = linea_grilla)
  g
}

# Graficos de un periodo: promedio mensual y una serie anual por columna
# (anual y, en el año hidrologico, los dos semestres con un eje Y comun)
graficar_periodo <- function(wb, hoja, res, cfg, pos, nombre_estacion) {

  u        <- cfg$unidad
  k        <- length(res$nombres_series)
  fila_fin <- length(res$etiquetas) + 2
  filas_graficos <- list(2:17, 19:34, 36:51, 53:68)
  cols_graficos  <- pos$graficos + 0:7

  # Promedio mensual (meses y valores, filas 2 a 14)
  datos <- wb_data(wb, sheet = hoja, dims = wb_dims(rows = 2:14, cols = pos$promedio + 0:1))
  datos$MES <- as.character(datos$MES)
  g <- ms_barchart(data = datos, x = "MES", y = names(datos)[2])
  g <- formatear_grafico(g, nombre_estacion,
                         paste0(cfg$grafico_promedio_mensual, " (", u, ")"),
                         limites_eje(res$prom_mensual))
  wb$add_mschart(sheet = hoja, dims = wb_dims(rows = filas_graficos[[1]], cols = cols_graficos),
                 graph = g)

  # Series anuales: eje Y comun para todas (anual, semestre 1 y semestre 2)
  limites_anual <- limites_eje(res$anual)
  for (j in seq_len(k)) {
    datos <- wb_data(wb, sheet = hoja, dims = wb_dims(rows = 2:fila_fin, cols = pos$anual + 0:j))
    datos <- datos[, c(1, ncol(datos))]               # columna YEAR y columna de valores
    nombre_y <- paste0(cfg$abrev, " ", res$nombres_series[j], " (", u, ")")
    names(datos) <- c("YEAR", nombre_y)
    datos$YEAR <- as.character(datos$YEAR)            # años como categorias (no 1, 2, 3...)
    g <- ms_barchart(data = datos, x = "YEAR", y = nombre_y)
    g <- formatear_grafico(g, nombre_estacion,
                           paste0(cfg$nombre_grafico, " ", res$nombres_series[j], " (", u, ")"),
                           limites_anual)
    g <- chart_ax_x(g, rotation = -90)                # texto del eje X rotado 90°
    wb$add_mschart(sheet = hoja, dims = wb_dims(rows = filas_graficos[[j + 1]], cols = cols_graficos),
                   graph = g)
  }
}

# Grafico de la serie diaria continua (fechas en DO, valores en DT),
# anclado en DW1 con tamaño 22.44 x 4.72 pulgadas
graficar_serie_diaria <- function(wb, hoja, cfg, fila_fin, nombre_estacion,
                                  col_ini = col_serie_continua) {
  datos <- wb_data(wb, sheet = hoja, dims = wb_dims(rows = 1:fila_fin, cols = col_ini + 0:5))
  datos <- datos[, c(1, ncol(datos))]
  nombre_y <- names(datos)[2]

  if (cfg$grafico_diario == "barras") {
    g <- ms_barchart(data = datos, x = "FECHA", y = nombre_y)
    g <- chart_settings(g, grouping = "clustered", dir = "vertical", gap_width = 0)
  } else {
    g <- ms_linechart(data = datos, x = "FECHA", y = nombre_y)
  }
  g <- formatear_grafico(g, paste0(cfg$titulo_diario, " - ", nombre_estacion),
                         paste0(cfg$nombre_grafico, " diaria (", cfg$unidad, ")"),
                         limites_eje(datos[[nombre_y]]))
  g <- chart_ax_x(g, rotation = -90, num_fmt = "mm/yyyy",
                  major_unit = 6, major_time_unit = "months",
                  minor_unit = 1, minor_time_unit = "months")

  cols_grafico <- (col_ini + 8):(col_ini + 40)       # DW:FC (33 columnas)
  wb$set_col_widths(sheet = hoja, cols = cols_grafico, widths = rep(8.43, 33))
  # Los offsets completan el tamaño sobre 33 columnas de 64 px y 21 filas de 16 pt.
  wb$add_mschart(sheet = hoja, dims = wb_dims(rows = 1:21, cols = cols_grafico), graph = g,
                 col_offset = c(0, 22.44 * 914400 - 33 * 64 * 9525),
                 row_offset = c(0, 4.72 * 914400 - 21 * 16 * 12700))
}

# =====================================================================
# 9. PROCESAMIENTO DE UNA PLANILLA
# =====================================================================
procesar_archivo <- function(file_xlsx, codigo) {

  cfg <- config_variables[[codigo]]

  # El archivo se lee una sola vez; las pestañas se extraen desde memoria
  wb_in         <- wb_load(file_xlsx)
  hojas         <- unname(wb_get_sheet_names(wb_in))
  df_estaciones <- wb_to_df(wb_in, sheet = "Estaciones")

  wb <- wb_workbook()

  for (i in seq_len(nrow(df_estaciones))) {

    hoja            <- df_estaciones$Pestana[i]
    nombre_estacion <- df_estaciones$Estacion[i]

    if (!(hoja %in% hojas)) {
      message("La pestana ", hoja, " (", nombre_estacion, ") no existe en ", basename(file_xlsx))
      next
    }

    # --- Lectura ---
    datos <- leer_estacion(wb_in, hoja, cfg)
    if (is.null(datos)) {
      message("La pestana ", hoja, " no tiene datos.")
      next
    }

    # --- Calculo ---
    res_cal <- calcular_periodo(datos$serie$FECHA, datos$serie$VALOR,
                                mes_inicio = 1, agregacion = cfg$agregacion,
                                hidrologico = FALSE)
    res_hid <- calcular_periodo(datos$serie$FECHA, datos$serie$VALOR,
                                mes_inicio = mes_inicio_hidro, agregacion = cfg$agregacion,
                                hidrologico = TRUE)

    # --- Escritura ---
    wb$add_worksheet(hoja)
    escribir_info(wb, hoja, datos$info)
    escribir_serie_original(wb, hoja, datos$serie, datos$nombre_valor)
    fila_fin_diaria <- escribir_serie_continua(wb, hoja, datos$serie, cfg, mes_inicio_hidro)
    pos_cal <- escribir_periodo(wb, hoja, res_cal, cfg, col_bloque_calendario)
    pos_hid <- escribir_periodo(wb, hoja, res_hid, cfg, col_bloque_hidro)

    # --- Graficos ---
    graficar_periodo(wb, hoja, res_cal, cfg, pos_cal, nombre_estacion)
    graficar_periodo(wb, hoja, res_hid, cfg, pos_hid, nombre_estacion)
    graficar_serie_diaria(wb, hoja, cfg, fila_fin_diaria, nombre_estacion)
  }

  archivo_salida <- paste0("Resumen_", codigo, ".xlsx")
  wb$save(file.path(dir_outputs, archivo_salida), overwrite = TRUE)
  message("Archivo generado: ", archivo_salida)
}

# =====================================================================
# 10. ITERACION SOBRE LAS PLANILLAS DE INPUTS
# =====================================================================
archivos <- list.files(dir_inputs, pattern = "\\.xlsx$", full.names = TRUE)
archivos <- archivos[!grepl("^~\\$", basename(archivos))]  # ignora temporales de Excel

for (file_xlsx in archivos) {

  codigo <- identificar_variable(basename(file_xlsx))

  if (is.na(codigo)) {
    message("No se pudo clasificar la planilla: ", basename(file_xlsx))
    next
  }

  if (!(codigo %in% variables_a_procesar)) {
    message("Planilla omitida (", codigo, " no esta en variables_a_procesar): ",
            basename(file_xlsx))
    next
  }

  message("Procesando ", codigo, ": ", basename(file_xlsx))
  t_ini <- Sys.time()
  procesar_archivo(file_xlsx, codigo)
  message("  Tiempo: ", round(as.numeric(difftime(Sys.time(), t_ini, units = "secs")), 1), " s")
}

message("El codigo ha terminado.")
