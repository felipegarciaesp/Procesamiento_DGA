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

## Pestaña "Resumen"
#-------------------
# Primer año que muestran las tablas de la pestaña Resumen (dias y meses con
# datos por año). Si alguna estacion tiene datos anteriores, el rango se
# extiende automaticamente hasta ese año. El ultimo año se determina a partir
# de los datos de todas las estaciones.
anio_inicio_resumen <- 1960

## Variables a procesar
#-------------------
# IMPORTANTE: en este vector deben ir SIEMPRE TODAS las variables que se
# quieren procesar en esta ejecucion. Solo se procesan las planillas de
# Inputs cuya variable este aqui; las demas se omiten.
#
#   - Cada variable procesada genera (o sobrescribe) su archivo en Outputs:
#     "PP" -> Outputs\PP\Resumen_PP.xlsx, "TMEDIA" -> Outputs\TMEDIA\Resumen_TMEDIA.xlsx, etc.
#   - Los archivos de las variables que NO estan en el vector NO se borran ni
#     se actualizan: quedan tal como estaban de una ejecucion anterior.
#   - Si se cambia algo que afecta a todas las variables (por ejemplo, el año
#     hidrologico o el formato de las tablas), se deben incluir TODAS las
#     variables en el vector. Si no, en Outputs quedaran archivos generados
#     con configuraciones distintas.
#
# Codigos disponibles: los nombres de config_variables (seccion 2).
# Ejemplo con todas las variables: c("PP", "TMEDIA", "TMAX", "TMIN")
variables_a_procesar <- c("PP", "TMEDIA", "TMAX", "TMIN")

## Directorio de trabajo
#-------------------
setwd("C:/Codigos/Procesamiento_DGA")
dir_inputs  <- file.path(getwd(), "Inputs")
dir_outputs <- file.path(getwd(), "Outputs")

## Carpeta de salida
#-------------------
# La carpeta Outputs NO se borra: se crea solo si no existe. Cada variable
# tiene su propia subcarpeta (Outputs\<VARIABLE>) y sobrescribe unicamente su
# archivo Resumen_<VARIABLE>.xlsx
# (ver "Variables a procesar" arriba).
if (!dir.exists(dir_outputs)) dir.create(dir_outputs, recursive = TRUE)

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
#   grafico_diario           : "barras" o "lineas" (tipo del grafico diario)
#   grafico_tipo             : "barras" o "lineas" (graficos mensual y anuales)
#   color_grafico            : color (hex) de relleno/borde de las barras o de la linea
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
            grafico_diario           = "barras",
            grafico_tipo             = "barras",
            color_grafico            = "#ADD8E6"),   # celeste (Light Blue)

  TMEDIA = list(patron_archivo           = "tmed",
                patron_columna           = "^tmed|temp|valor",
                agregacion               = "promedio",
                unidad                   = "°C",
                nombre_tabla             = "Temperatura media",
                abrev                    = "Tmedia",
                nombre_grafico           = "Temperatura media",
                grafico_promedio_mensual = "Temperatura media mensual",
                titulo_diario            = "Temperatura media diaria",
                grafico_diario           = "lineas",
                grafico_tipo             = "lineas",
                color_grafico            = "#00B050"),   # verde

  TMAX = list(patron_archivo           = "tmax",
              patron_columna           = "^tmax|temp|valor",
              agregacion               = "promedio",
              unidad                   = "°C",
              nombre_tabla             = "Temperatura máxima",
              abrev                    = "Tmax",
              nombre_grafico           = "Temperatura máxima",
              grafico_promedio_mensual = "Temperatura máxima media mensual",
              titulo_diario            = "Temperatura máxima diaria",
              grafico_diario           = "lineas",
              grafico_tipo             = "lineas",
              color_grafico            = "#FF0000"),   # rojo

  TMIN = list(patron_archivo           = "tmin",
              patron_columna           = "^tmin|temp|valor",
              agregacion               = "promedio",
              unidad                   = "°C",
              nombre_tabla             = "Temperatura mínima",
              abrev                    = "Tmin",
              nombre_grafico           = "Temperatura mínima",
              grafico_promedio_mensual = "Temperatura mínima media mensual",
              titulo_diario            = "Temperatura mínima diaria",
              grafico_diario           = "lineas",
              grafico_tipo             = "lineas",
              color_grafico            = "#0070C0")    # azul

  # Ejemplo para caudales (cuando exista la planilla de entrada):
  # , Q = list(patron_archivo = "^q|caudal", patron_columna = "caudal|valor",
  #            agregacion = "promedio", unidad = "m3/s",
  #            nombre_tabla = "Caudal", abrev = "Q", nombre_grafico = "Caudal",
  #            grafico_promedio_mensual = "Caudal medio mensual",
  #            titulo_diario = "Caudales diarios", grafico_diario = "lineas",
  #            grafico_tipo = "lineas", color_grafico = "#7030A0")
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

## Disposicion de las tablas en las pestañas de salida
# Las tablas y graficos se ubican automaticamente de izquierda a derecha, uno
# a continuacion del otro, en el orden definido en disponer_hoja() (seccion 7).
# Sus posiciones se calculan a partir del ancho real de cada tabla, por lo que
# agregar, quitar o ensanchar una tabla desplaza automaticamente a las demas.
# Para cambiar separaciones o tamaños basta con modificar estos valores.
separacion_tablas     <- 1    # columnas vacias entre tablas y bloques
separacion_tras_serie <- 4    # columnas vacias despues de la serie original (G:J)
ancho_grafico         <- 8    # columnas que ocupa cada grafico de periodo
alto_grafico          <- 16   # filas que ocupa cada grafico de periodo
separacion_graficos   <- 1    # filas vacias entre graficos de periodo
ancho_grafico_diario  <- 33   # columnas del grafico de la serie diaria
alto_grafico_diario   <- 21   # filas del grafico de la serie diaria

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
       anios          = anios,
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
# 7. ARMADO DE TABLAS Y DISPOSICION EN LA PESTAÑA (sin Excel)
# =====================================================================
# Primero se arman todas las tablas como data.frames; luego se calcula donde
# va cada una segun su ancho (disposicion automatica) y recien despues se
# escriben en Excel (seccion 8).

# Serie diaria original (columnas FECHA, valor, Unidad)
armar_serie_original <- function(serie, nombre_valor) {
  tabla <- serie
  names(tabla) <- c("FECHA", nombre_valor, "Unidad")
  tabla
}

# Serie diaria continua: todas las fechas entre la minima y la maxima de la
# serie original. DIA, MES, YEAR y Hydro YEAR se calculan aqui como valores.
armar_serie_continua <- function(serie, cfg, mes_inicio_hidro, hoja) {
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
  tabla
}

# Arma las tablas de un periodo (calendario o hidrologico) con sus encabezados
armar_tablas_periodo <- function(res, cfg) {
  u <- cfg$unidad
  t <- list()

  # Cantidad de registros por mes
  t$conteo <- data.frame(res$etiquetas, res$conteo)
  names(t$conteo) <- c("YEAR", res$meses)

  # Cantidad de meses con informacion
  t$meses <- data.frame(YEAR = res$etiquetas, MESES = res$n_meses)

  # Valores mensuales (tabla año x mes)
  t$mensual <- data.frame(res$etiquetas, res$mensual)
  names(t$mensual) <- c("YEAR", res$meses)

  # Valores mensuales en vector
  t$vector <- res$vector_mens
  names(t$vector) <- c("FECHA", "MES", "YEAR", paste0(cfg$abrev, " (", u, ")"))

  # Promedio mensual entre años
  t$promedio <- data.frame(MES = res$meses, VALOR = res$prom_mensual)
  names(t$promedio) <- c("MES", paste0(cfg$abrev, " promedio (", u, ")"))

  # Resumen bajo la tabla de promedios (anual y, si corresponde, semestres)
  t$resumen <- list(etiquetas = paste0(res$nombres_resum, " (", u, ")"),
                    valores   = res$resumen_prom)

  # Series anuales y sus promedios (estos van a la derecha de la tabla anual)
  t$anual <- data.frame(res$etiquetas, res$anual)
  names(t$anual) <- c("YEAR", paste0(cfg$abrev, " ", res$nombres_series, " (", u, ")"))

  t$prom_anual <- as.data.frame(as.list(res$prom_anual))
  names(t$prom_anual) <- paste0(cfg$abrev, " ", res$nombres_prom, " (", u, ")")

  t
}

# --- Disposicion automatica ---
# Un "cursor" recorre las columnas de izquierda a derecha. Cada tabla reserva
# tantas columnas como tenga, mas la separacion, y el cursor avanza.
nueva_disposicion <- function(col_ini = 1) {
  disp <- new.env()
  disp$cursor <- col_ini
  disp
}

reservar_columnas <- function(disp, ancho, separacion = separacion_tablas) {
  col <- disp$cursor
  disp$cursor <- col + ancho + separacion
  col
}

# Columnas de inicio de cada tabla de un periodo y de sus graficos
disponer_periodo <- function(disp, tablas) {
  pos <- list()
  pos$conteo   <- reservar_columnas(disp, ncol(tablas$conteo))
  pos$meses    <- reservar_columnas(disp, ncol(tablas$meses))
  pos$mensual  <- reservar_columnas(disp, ncol(tablas$mensual))
  pos$vector   <- reservar_columnas(disp, ncol(tablas$vector))
  pos$promedio <- reservar_columnas(disp, ncol(tablas$promedio))
  pos$anual    <- reservar_columnas(disp, ncol(tablas$anual) + ncol(tablas$prom_anual))
  pos$graficos <- reservar_columnas(disp, ancho_grafico)
  pos
}

# Orden de los bloques en la pestaña, de izquierda a derecha.
# Para reordenar los bloques basta con cambiar el orden de estas lineas.
disponer_hoja <- function(tabla_orig, tabla_cont, tablas_cal, tablas_hid) {
  disp <- nueva_disposicion(col_ini = 1)
  pos <- list()
  pos$info           <- reservar_columnas(disp, 2)    # rotulos + valores
  pos$serie          <- reservar_columnas(disp, ncol(tabla_orig), separacion_tras_serie)
  pos$cal            <- disponer_periodo(disp, tablas_cal)
  pos$hid            <- disponer_periodo(disp, tablas_hid)
  pos$continua       <- reservar_columnas(disp, ncol(tabla_cont))
  pos$grafico_diario <- reservar_columnas(disp, ancho_grafico_diario)
  pos
}

# =====================================================================
# 8. ESCRITURA EN EXCEL
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

# Escribe una tabla con titulo opcional (fila 1) y encabezado en negrita (fila 2)
escribir_tabla <- function(wb, hoja, df, col_ini, titulo = NULL, col_fin_titulo = NULL) {
  col_fin <- col_ini + ncol(df) - 1
  if (!is.null(titulo)) {
    escribir_titulo(wb, hoja, titulo, col_ini, if (is.null(col_fin_titulo)) col_fin else col_fin_titulo)
  }
  wb$add_data(sheet = hoja, x = df, start_col = col_ini, start_row = 2, na = na_strings())
  wb$add_font(sheet = hoja, dims = wb_dims(rows = 2, cols = col_ini:col_fin), bold = TRUE)
}

# Tabla de informacion de la estacion (2 columnas: rotulos y valores)
escribir_info <- function(wb, hoja, info, col) {
  escribir_titulo(wb, hoja, "Informacion de la Estación", col, col + 1)
  wb$add_data(sheet = hoja, x = info_rotulos, start_col = col,     start_row = 2, col_names = FALSE)
  wb$add_data(sheet = hoja, x = info,         start_col = col + 1, start_row = 2, col_names = FALSE)
  # Anchos de las 2 columnas y de la columna separadora angosta a su derecha
  wb$set_col_widths(sheet = hoja, cols = col + 0:2, widths = c(22, 28, 3))
}

# Serie diaria original (encabezado en la fila 1)
escribir_serie_original <- function(wb, hoja, tabla, col) {
  wb$add_data(sheet = hoja, x = tabla, start_col = col, start_row = 1)
  wb$add_font(sheet = hoja, dims = wb_dims(rows = 1, cols = col + 0:(ncol(tabla) - 1)), bold = TRUE)
  wb$add_numfmt(sheet = hoja, dims = wb_dims(rows = 2:(nrow(tabla) + 1), cols = col),
                numfmt = "dd/mm/yyyy")
  wb$set_col_widths(sheet = hoja, cols = col + 0:(ncol(tabla) - 1), widths = c(12, 20, 10))
}

# Serie diaria continua (encabezado en la fila 1)
escribir_serie_continua <- function(wb, hoja, tabla, col) {
  wb$add_data(sheet = hoja, x = tabla, start_col = col, start_row = 1, na = na_strings())
  wb$add_font(sheet = hoja, dims = wb_dims(rows = 1, cols = col + 0:(ncol(tabla) - 1)), bold = TRUE)
  wb$add_numfmt(sheet = hoja, dims = wb_dims(rows = 2:(nrow(tabla) + 1), cols = col),
                numfmt = "dd/mm/yyyy")
  wb$set_col_widths(sheet = hoja, cols = col + 0:(ncol(tabla) - 1),
                    widths = c(12, 6, 6, 8, 12, 12, 10))
}

# Escribe las tablas de un periodo (calendario o hidrologico) en las
# posiciones calculadas por disponer_periodo()
escribir_periodo <- function(wb, hoja, res, tablas, cfg, pos) {

  u          <- cfg$unidad
  sufijo     <- if (res$hidrologico) " - Año hidrologico" else " - Año calendario"
  ancho_year <- if (res$hidrologico) 10 else 8
  n          <- nrow(tablas$conteo)
  k          <- ncol(tablas$anual) - 1   # 1 (anual) o 3 (anual + 2 semestres)

  # --- Cantidad de registros por mes ---
  escribir_tabla(wb, hoja, tablas$conteo, pos$conteo,
                 paste0("Cantidad de registros por mes", sufijo))
  wb$set_col_widths(sheet = hoja, cols = pos$conteo + 0:(ncol(tablas$conteo) - 1),
                    widths = c(ancho_year, rep(6, ncol(tablas$conteo) - 1)))

  # Escala de colores: 0 = rojo, 15 = amarillo, 31 = verde oscuro
  wb$add_conditional_formatting(sheet = hoja,
                                dims = wb_dims(rows = 3:(n + 2), cols = pos$conteo + 1:(ncol(tablas$conteo) - 1)),
                                type = "colorScale",
                                style = c("#F8696B", "#FFEB84", "#006400"),
                                rule = c(0, 15, 31))

  # --- Cantidad de meses con informacion ---
  escribir_tabla(wb, hoja, tablas$meses, pos$meses,
                 paste0("Cantidad de meses con información", sufijo))
  wb$set_col_widths(sheet = hoja, cols = pos$meses + 0:1, widths = c(ancho_year, 8))

  # --- Valores mensuales (tabla año x mes) ---
  escribir_tabla(wb, hoja, tablas$mensual, pos$mensual,
                 paste0(cfg$nombre_tabla, " mensual (", u, ")", sufijo))
  wb$set_col_widths(sheet = hoja, cols = pos$mensual + 0:(ncol(tablas$mensual) - 1),
                    widths = c(ancho_year, rep(8, ncol(tablas$mensual) - 1)))

  # --- Valores mensuales en vector ---
  escribir_tabla(wb, hoja, tablas$vector, pos$vector,
                 paste0(cfg$abrev, " mensual (", u, ")", sufijo))
  wb$add_numfmt(sheet = hoja, dims = wb_dims(rows = 3:(nrow(tablas$vector) + 2), cols = pos$vector),
                numfmt = "dd/mm/yyyy")
  wb$set_col_widths(sheet = hoja, cols = pos$vector + 0:3, widths = c(12, 6, ancho_year, 10))

  # --- Promedio mensual entre años + resumen anual/semestral debajo ---
  escribir_tabla(wb, hoja, tablas$promedio, pos$promedio,
                 paste0(cfg$nombre_tabla, " promedio mensual (", u, ")", sufijo))

  fila_resumen <- 2 + nrow(tablas$promedio) + 1   # fila siguiente a la tabla (15)
  wb$add_data(sheet = hoja, x = tablas$resumen$etiquetas,
              start_col = pos$promedio, start_row = fila_resumen, col_names = FALSE)
  wb$add_data(sheet = hoja, x = tablas$resumen$valores,
              start_col = pos$promedio + 1, start_row = fila_resumen, col_names = FALSE,
              na = na_strings())
  wb$add_font(sheet = hoja,
              dims = wb_dims(rows = fila_resumen + 0:(k - 1), cols = pos$promedio + 0:1),
              bold = TRUE)
  wb$set_col_widths(sheet = hoja, cols = pos$promedio + 0:1,
                    widths = c(if (k > 1) 28 else 20, 16))

  # --- Series anuales y, a su derecha, sus promedios (un unico valor) ---
  ancho_total <- ncol(tablas$anual) + ncol(tablas$prom_anual)
  escribir_tabla(wb, hoja, tablas$anual, pos$anual,
                 paste0(cfg$abrev, " anual y promedio anual (", u, ")", sufijo),
                 col_fin_titulo = pos$anual + ancho_total - 1)
  escribir_tabla(wb, hoja, tablas$prom_anual, pos$anual + ncol(tablas$anual))
  wb$set_col_widths(sheet = hoja, cols = pos$anual + 0:(ancho_total - 1),
                    widths = c(ancho_year, 14, rep(20, k - 1), 22, rep(26, k - 1)))
}

# =====================================================================
# 9. GRAFICOS (nativos de Excel, enlazados a las tablas de la pestaña)
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

# Crea el grafico (barras o lineas) y le aplica el color de la variable:
# relleno y borde en barras, trazo sin marcadores en lineas
crear_grafico <- function(datos, x, y, tipo, color) {
  if (tipo == "barras") {
    g <- ms_barchart(data = datos, x = x, y = y)
    g <- chart_data_fill(g, values = setNames(color, y))
    g <- chart_data_stroke(g, values = setNames(color, y))
  } else {
    g <- ms_linechart(data = datos, x = x, y = y)
    g <- chart_data_stroke(g, values = setNames(color, y))
    g <- chart_data_line_width(g, values = setNames(1.5, y))
    g <- chart_data_symbol(g, values = setNames("none", y))
  }
  g
}

# Filas que ocupa el grafico numero "i" de la columna de graficos (1, 2, 3...)
filas_grafico <- function(i) {
  fila_ini <- 2 + (i - 1) * (alto_grafico + separacion_graficos)
  fila_ini:(fila_ini + alto_grafico - 1)
}

# Graficos de un periodo, apilados verticalmente: promedio mensual y una serie
# anual por columna (anual y, en el año hidrologico, los dos semestres con un
# eje Y comun)
graficar_periodo <- function(wb, hoja, res, tablas, cfg, pos, nombre_estacion) {

  u        <- cfg$unidad
  k        <- ncol(tablas$anual) - 1
  cols_graficos <- pos$graficos + 0:(ancho_grafico - 1)

  # Promedio mensual (meses y valores)
  fila_fin <- 2 + nrow(tablas$promedio)
  datos <- wb_data(wb, sheet = hoja, dims = wb_dims(rows = 2:fila_fin, cols = pos$promedio + 0:1))
  datos$MES <- as.character(datos$MES)
  names(datos) <- c("MES", names(tablas$promedio)[2])   # evita diferencias de codificacion en "°C"
  g <- crear_grafico(datos, "MES", names(datos)[2], cfg$grafico_tipo, cfg$color_grafico)
  g <- formatear_grafico(g, nombre_estacion,
                         paste0(cfg$grafico_promedio_mensual, " (", u, ")"),
                         limites_eje(res$prom_mensual))
  wb$add_mschart(sheet = hoja, dims = wb_dims(rows = filas_grafico(1), cols = cols_graficos),
                 graph = g)

  # Series anuales: eje Y comun para todas (anual, semestre 1 y semestre 2)
  fila_fin <- 2 + nrow(tablas$anual)
  limites_anual <- limites_eje(res$anual)
  for (j in seq_len(k)) {
    datos <- wb_data(wb, sheet = hoja, dims = wb_dims(rows = 2:fila_fin, cols = pos$anual + 0:j))
    datos <- datos[, c(1, ncol(datos))]               # columna YEAR y columna de valores
    nombre_y <- names(tablas$anual)[j + 1]
    names(datos) <- c("YEAR", nombre_y)
    datos$YEAR <- as.character(datos$YEAR)            # años como categorias (no 1, 2, 3...)
    g <- crear_grafico(datos, "YEAR", nombre_y, cfg$grafico_tipo, cfg$color_grafico)
    g <- formatear_grafico(g, nombre_estacion,
                           paste0(cfg$nombre_grafico, " ", res$nombres_series[j], " (", u, ")"),
                           limites_anual)
    g <- chart_ax_x(g, rotation = -90)                # texto del eje X rotado 90°
    wb$add_mschart(sheet = hoja, dims = wb_dims(rows = filas_grafico(j + 1), cols = cols_graficos),
                   graph = g)
  }
}

# Grafico de la serie diaria continua, con tamaño 22.44 x 4.72 pulgadas
graficar_serie_diaria <- function(wb, hoja, cfg, tabla, col_serie, col_grafico, nombre_estacion) {

  # Columnas FECHA (primera) y valor diario dentro de la serie continua
  nombre_y  <- paste(cfg$abrev, "diaria")
  col_valor <- col_serie + which(names(tabla) == nombre_y) - 1
  fila_fin  <- nrow(tabla) + 1

  datos <- wb_data(wb, sheet = hoja, dims = wb_dims(rows = 1:fila_fin, cols = col_serie:col_valor))
  datos <- datos[, c(1, ncol(datos))]

  g <- crear_grafico(datos, "FECHA", nombre_y, cfg$grafico_diario, cfg$color_grafico)
  if (cfg$grafico_diario == "barras") {
    g <- chart_settings(g, grouping = "clustered", dir = "vertical", gap_width = 0)
  }
  g <- formatear_grafico(g, paste0(cfg$titulo_diario, " - ", nombre_estacion),
                         paste0(cfg$nombre_grafico, " diaria (", cfg$unidad, ")"),
                         limites_eje(datos[[nombre_y]]))
  g <- chart_ax_x(g, rotation = -90, num_fmt = "mm/yyyy",
                  major_unit = 6, major_time_unit = "months",
                  minor_unit = 1, minor_time_unit = "months")

  cols_grafico <- col_grafico + 0:(ancho_grafico_diario - 1)
  wb$set_col_widths(sheet = hoja, cols = cols_grafico, widths = rep(8.43, ancho_grafico_diario))
  # Los offsets completan el tamaño sobre columnas de 64 px y filas de 16 pt.
  wb$add_mschart(sheet = hoja, dims = wb_dims(rows = 1:alto_grafico_diario, cols = cols_grafico),
                 graph = g,
                 col_offset = c(0, 22.44 * 914400 - ancho_grafico_diario * 64 * 9525),
                 row_offset = c(0, 4.72 * 914400 - alto_grafico_diario * 16 * 12700))
}

# =====================================================================
# 9b. PLANILLA DE CRITERIOS (Criterios_<VARIABLE>.xlsx)
# =====================================================================
# Planilla que completa el usuario DESPUES de revisar Resumen_<VARIABLE>.xlsx.
# Un codigo posterior leera Resumen + Criterios y generara la planilla filtrada.
#
# Tiene una pestaña por estacion, con el mismo nombre que en Resumen (si una
# estacion no esta en Resumen, tampoco esta en Criterios). Columnas:
#
#   A  "Año inicio" / B "Año fin"      : periodo de analisis (años calendario,
#                                        inclusive). Se analiza desde enero del
#                                        año de inicio hasta diciembre del año
#                                        final. Vacio = toda la serie.
#   D  "Dias minimos"                  : un solo numero (fila 2). Dias con dato
#                                        que debe tener un mes para ser valido.
#                                        Si tiene menos, se descartan TODOS los
#                                        datos de ese mes.
#   E  "Meses a aplicar"               : numeros de mes (1-12), uno por fila,
#                                        en los que aplica "Dias minimos".
#                                        Vacio = no se aplica el filtro.
#   F  "Meses minimos para anual"      : un solo numero (fila 2). Cantidad
#                                        minima de meses validos para calcular
#                                        el valor anual; si no, queda #N/A.
#                                        Vacio = no se exige minimo (el anual se
#                                        calcula ignorando los #N/A).
#   G  "Meses obligatorios para anual" : numeros de mes (1-12), uno por fila.
#                                        Todos deben tener valor para calcular el
#                                        valor anual; si falta alguno, queda #N/A.
#                                        Vacio = da lo mismo cuales meses esten
#                                        completos. (F y G se pueden usar juntas:
#                                        el anual se calcula solo si se cumplen ambas.)
#   I  "Fechas a eliminar"             : fechas puntuales (dd/mm/yyyy), una por fila.
#   J  "Rango: inicio" / K "Rango: fin": periodo completo a eliminar (inclusive).
#   L  "Comentario"                    : motivo de la eliminacion (trazabilidad).
#
# En I:L cada fila es UNA eliminacion: o bien una fecha puntual (I) o bien un
# rango (J y K), con su comentario en L.
#
# NOTAS PARA EL CODIGO QUE USE ESTA PLANILLA:
#   1. Orden del filtro: primero se eliminan las fechas puntuales y los rangos
#      y RECIEN DESPUES se cuentan los dias del mes. Ej.: si en un mes se
#      eliminan 12 datos anomalos y quedan 18 dias, ese mes queda descartado
#      por el minimo de 20.
#   2. Un dato eliminado queda como faltante (#N/A), NO como 0.
#   3. Los criterios del valor anual (F y G) se evaluan sobre los meses que
#      queden validos despues del paso 1, y se aplican igual al año calendario
#      y al año hidrologico (en este ultimo, los meses son los del año hidrologico).
#   4. Periodo de analisis (A y B): corresponde a enero del año de inicio hasta
#      diciembre del año final. Para el año hidrologico (con el ejemplo ABR-MAR),
#      el primer año hidrologico es X-1/X y el ultimo es Y/Y+1; ej.: 1980-2026
#      -> 1979/80 a 2026/27. En general, se incluyen los años hidrologicos que
#      contienen algun mes del periodo; si mes_inicio_hidro = 1 no hay desfase.
#
# Si Criterios_<VARIABLE>.xlsx ya existe NO se sobrescribe, para no perder lo
# que el usuario haya ingresado. Para regenerarlo, borrar el archivo.
crear_criterios <- function(codigo, hojas) {

  archivo <- paste0("Criterios_", codigo, ".xlsx")
  ruta    <- file.path(dir_outputs, codigo, archivo)

  if (file.exists(ruta)) {
    message("Se mantiene ", archivo, " existente (no se sobrescribe).")
    return(invisible(NULL))
  }

  encabezados <- c(A = "Año inicio",
                   B = "Año fin",
                   D = "Dias minimos",
                   E = "Meses a aplicar",
                   F = "Meses minimos para anual",
                   G = "Meses obligatorios para anual",
                   I = "Fechas a eliminar",
                   J = "Rango: inicio",
                   K = "Rango: fin",
                   L = "Comentario")
  columnas <- match(names(encabezados), LETTERS)
  anchos   <- c(12, 12, 3, 14, 17, 26, 30, 3, 18, 15, 15, 45)
  filas_validacion <- 2:1001

  validar_entero <- function(wb, hoja, filas, col, minimo, maximo, titulo, mensaje) {
    wb$add_data_validation(sheet = hoja, dims = wb_dims(rows = filas, cols = col),
                           type = "whole", operator = "between",
                           value = c(minimo, maximo),
                           error_title = titulo, error = mensaje)
  }

  wb <- wb_workbook()
  for (hoja in hojas) {
    wb$add_worksheet(hoja)
    for (k in seq_along(encabezados)) {
      wb$add_data(sheet = hoja, x = encabezados[[k]], start_col = columnas[k],
                  start_row = 1, col_names = FALSE)
    }
    wb$add_font(sheet = hoja, dims = wb_dims(rows = 1, cols = columnas), bold = TRUE)
    wb$add_cell_style(sheet = hoja, dims = wb_dims(rows = 1, cols = columnas),
                      horizontal = "center", wrap_text = TRUE)
    wb$set_col_widths(sheet = hoja, cols = 1:12, widths = anchos)
    wb$add_numfmt(sheet = hoja, dims = wb_dims(rows = filas_validacion, cols = 9:11),
                  numfmt = "dd/mm/yyyy")
    wb$freeze_pane(sheet = hoja, first_row = TRUE)

    # Validaciones de entrada
    validar_entero(wb, hoja, 2, 1:2, 1800, 2200, "Año no valido",
                   "Ingrese un año de 4 digitos (por ejemplo, 1980).")
    validar_entero(wb, hoja, 2, 4, 1, 31, "Valor no valido",
                   "Ingrese un numero entero de dias (1 a 31).")
    validar_entero(wb, hoja, filas_validacion, 5, 1, 12, "Mes no valido",
                   "Ingrese el numero de mes (1 a 12).")
    validar_entero(wb, hoja, 2, 6, 1, 12, "Valor no valido",
                   "Ingrese un numero entero de meses (1 a 12).")
    validar_entero(wb, hoja, filas_validacion, 7, 1, 12, "Mes no valido",
                   "Ingrese el numero de mes (1 a 12).")
  }

  tryCatch(
    wb$save(ruta, overwrite = TRUE),
    error = function(e) {
      stop("No se pudo guardar ", archivo, ": ", conditionMessage(e), call. = FALSE)
    }
  )
  message("Archivo generado: ", archivo)
}
# =====================================================================
# 9c. PESTAÑA "Resumen" (primera pestaña de Resumen_<VARIABLE>.xlsx)
# =====================================================================
# Contiene, de izquierda a derecha:
#   - Tabla de estaciones (ID, nombre, codigo, coordenadas, altitud y periodo
#     de registro). El ID es el nombre de la pestaña de cada estacion.
#   - Dias con dato por año, en año calendario y, justo debajo, en año
#     hidrologico (escala de color rojo-amarillo-verde de 0 a 366 dias).
#   - Meses con dato por año, tambien en año calendario y año hidrologico
#     (semaforo: verde = 12 meses, amarillo >= 7, rojo < 7, vacio = sin datos)
#     y su leyenda al final.
#
# Todo se arma a partir de lo que ya se calculo para cada estacion (no se
# vuelve a leer ninguna pestaña). Las columnas se ubican automaticamente.
# En el año hidrologico, cada columna queda bajo el año calendario en que
# termina (1959/60 bajo 1960). Su ultimo año es el ultimo con datos.

# Datos de una estacion para la pestaña Resumen
datos_resumen_estacion <- function(hoja, nombre_estacion, datos, res_cal, res_hid) {
  validas <- datos$serie$FECHA[!is.na(datos$serie$FECHA) & !is.na(datos$serie$VALOR)]
  f <- range(validas)
  mes_anio <- function(x) paste0(as.integer(format(x, "%m")), "/", format(x, "%Y"))
  list(id      = hoja,
       nombre  = nombre_estacion,
       codigo  = datos$info[2],
       este    = datos$info[4],
       norte   = datos$info[5],
       altitud = datos$info[6],
       periodo = paste(mes_anio(f[1]), "-", mes_anio(f[2])),
       cal = list(anios = res_cal$anios, dias = rowSums(res_cal$conteo), meses = res_cal$n_meses),
       hid = list(anios = res_hid$anios, dias = rowSums(res_hid$conteo), meses = res_hid$n_meses))
}

# Matriz estaciones x años (0 donde la estacion no tiene datos)
matriz_resumen <- function(est, periodo, campo, anios_cols) {
  m <- matrix(0L, nrow = length(est), ncol = length(anios_cols))
  for (i in seq_along(est)) {
    r   <- est[[i]][[periodo]]
    pos <- match(r$anios, anios_cols)
    ok  <- !is.na(pos)
    m[i, pos[ok]] <- as.integer(r[[campo]][ok])
  }
  m
}

# Convierte a numero si todos los valores lo permiten; si no, deja el texto
columna_numerica <- function(x) {
  v <- suppressWarnings(as.numeric(x))
  if (all(!is.na(v) | is.na(x))) v else x
}

# Escribe una tabla estaciones x años con encabezados por decada.
#   tipo = "dias"  -> escala de color de 0 a 366
#   tipo = "meses" -> semaforo; las celdas con 0 quedan vacias
escribir_tabla_resumen <- function(wb, hoja, fila, col_etq, col_ini, etq_filas,
                                   etq_anios, decadas, valores, titulo, tipo) {
  n <- length(etq_filas)
  k <- length(etq_anios)
  f_hdr <- fila + 1
  f_ini <- fila + 2
  f_fin <- fila + 1 + n
  cols  <- col_ini + 0:(k - 1)

  # Titulo y rotulos de la columna de estaciones
  wb$add_data(sheet = hoja, x = titulo, start_col = col_etq, start_row = fila, col_names = FALSE)
  wb$add_cell_style(sheet = hoja, dims = wb_dims(rows = fila, cols = col_etq),
                    horizontal = "center", vertical = "bottom", wrap_text = TRUE)
  wb$add_data(sheet = hoja, x = "Estación", start_col = col_etq, start_row = f_hdr, col_names = FALSE)
  wb$add_data(sheet = hoja, x = etq_filas,  start_col = col_etq, start_row = f_ini, col_names = FALSE)
  wb$add_font(sheet = hoja, dims = wb_dims(rows = c(fila, f_hdr), cols = col_etq), bold = TRUE)
  wb$set_col_widths(sheet = hoja, cols = col_etq, widths = max(12, nchar(etq_filas)) + 2)
  wb$add_border(sheet = hoja, dims = wb_dims(rows = f_hdr:f_fin, cols = col_etq),
                top_border = "medium", bottom_border = "medium",
                left_border = "medium", right_border = "medium", inner_hgrid = "thin")

  # Fila de años (rotados) y valores
  wb$add_data(sheet = hoja, x = matrix(etq_anios, nrow = 1), start_col = col_ini,
              start_row = fila, col_names = FALSE)
  wb$add_cell_style(sheet = hoja, dims = wb_dims(rows = fila, cols = cols),
                    horizontal = "center", vertical = "bottom", text_rotation = 90)
  if (tipo == "meses") valores[valores == 0] <- NA
  # Las celdas sin dato (NA) no se escriben, para que queden realmente vacias
  if (tipo == "dias") {
    wb$add_data(sheet = hoja, x = valores, start_col = col_ini, start_row = f_ini, col_names = FALSE)
  } else {
    for (i in seq_len(nrow(valores))) {
      ok <- which(!is.na(valores[i, ]))
      if (length(ok) == 0) next
      pos <- split(ok, cumsum(c(1, diff(ok) != 1)))   # tramos contiguos
      for (p in pos) {
        wb$add_data(sheet = hoja, x = matrix(valores[i, p], nrow = 1), start_col = col_ini + p[1] - 1,
                    start_row = f_ini + i - 1, col_names = FALSE)
      }
    }
  }
  wb$add_cell_style(sheet = hoja, dims = wb_dims(rows = f_ini:f_fin, cols = cols),
                    horizontal = "center")

  # Una caja por decada, con su rotulo "primer año-ultimo año"
  r <- rle(decadas)
  fin_g <- cumsum(r$lengths)
  ini_g <- fin_g - r$lengths + 1
  for (g in seq_along(r$lengths)) {
    cols_g <- col_ini + (ini_g[g] - 1):(fin_g[g] - 1)
    rotulo <- if (r$lengths[g] > 1) paste0(etq_anios[ini_g[g]], "-", etq_anios[fin_g[g]]) else
      as.character(etq_anios[ini_g[g]])
    wb$add_data(sheet = hoja, x = rotulo, start_col = cols_g[1], start_row = f_hdr, col_names = FALSE)
    if (length(cols_g) > 1) wb$merge_cells(sheet = hoja, dims = wb_dims(rows = f_hdr, cols = cols_g))
    wb$add_font(sheet = hoja, dims = wb_dims(rows = f_hdr, cols = cols_g[1]), bold = TRUE)
    wb$add_cell_style(sheet = hoja, dims = wb_dims(rows = f_hdr, cols = cols_g[1]),
                      horizontal = "center")
    wb$add_border(sheet = hoja, dims = wb_dims(rows = f_hdr:f_fin, cols = cols_g),
                  top_border = "medium", bottom_border = "medium",
                  left_border = "medium", right_border = "medium",
                  inner_hgrid = "thin", inner_vgrid = "thin")
  }

  # Formato condicional
  dims_valores <- wb_dims(rows = f_ini:f_fin, cols = cols)
  if (tipo == "dias") {
    wb$add_conditional_formatting(sheet = hoja, dims = dims_valores, type = "colorScale",
                                  style = c("#F8696B", "#FFEB84", "#63BE7B"),
                                  rule = c(0, 183, 366))
  } else {
    wb$add_conditional_formatting(sheet = hoja, dims = dims_valores, type = "iconSet",
                                  rule = c(0, 7, 12),
                                  params = list(showValue = FALSE, iconSet = "3TrafficLights1"))
  }
}

# Arma y escribe toda la pestaña Resumen
escribir_resumen <- function(wb, hoja, est) {
  n  <- length(est)
  d  <- as.integer(mes_inicio_hidro > 1)   # desfase del año hidrologico
  ids     <- vapply(est, function(e) e$id,     character(1))
  nombres <- vapply(est, function(e) e$nombre, character(1))

  # Años de las tablas: calendario X; hidrologico X/X+1 bajo el año en que termina
  anios_cal <- unlist(lapply(est, function(e) e$cal$anios))
  anios_hid <- unlist(lapply(est, function(e) e$hid$anios))
  a_ini     <- min(anio_inicio_resumen, anios_cal)
  cols_cal  <- a_ini:max(anios_cal)
  cols_hid  <- (a_ini - d):max(anios_hid, a_ini - d)
  etq_cal   <- cols_cal
  etq_hid   <- sprintf("%d/%02d", cols_hid, (cols_hid + 1) %% 100)
  dec_cal   <- cols_cal %/% 10 * 10
  dec_hid   <- (cols_hid + d) %/% 10 * 10

  # Disposicion
  ncol_info <- 7
  col_etq1  <- ncol_info + 2                               # tabla de dias
  col_ini1  <- col_etq1 + 1
  ancho     <- max(length(cols_cal), length(cols_hid))
  col_etq2  <- col_ini1 + ancho + separacion_tablas        # tabla de meses
  col_ini2  <- col_etq2 + 1
  alto      <- n + 2
  fila_cal  <- 1
  fila_hid  <- fila_cal + alto + 1
  fila_ley  <- fila_hid + alto + 1

  # Tabla de estaciones
  encabezados <- list(c("ID", 1), c("Nombre", 2), c("Código", 3), c("Altitud (m s.n.m.)", 6),
                      c("Periodo de Registro", 7))
  for (h in encabezados) {
    col <- as.integer(h[2])
    wb$add_data(sheet = hoja, x = h[1], start_col = col, start_row = 1, col_names = FALSE)
    wb$merge_cells(sheet = hoja, dims = wb_dims(rows = 1:2, cols = col))
  }
  wb$add_data(sheet = hoja, x = "Coordenadas UTM", start_col = 4, start_row = 1, col_names = FALSE)
  wb$merge_cells(sheet = hoja, dims = wb_dims(rows = 1, cols = 4:5))
  wb$add_data(sheet = hoja, x = matrix(c("UTM NORTE (m)", "UTM ESTE (m)"), nrow = 1),
              start_col = 4, start_row = 2, col_names = FALSE)
  wb$add_cell_style(sheet = hoja, dims = wb_dims(rows = 1:2, cols = 1:ncol_info),
                    horizontal = "center", vertical = "center", wrap_text = TRUE)

  columnas <- list(ids, nombres,
                   columna_numerica(vapply(est, function(e) e$codigo, character(1))),
                   columna_numerica(vapply(est, function(e) e$norte, character(1))),
                   columna_numerica(vapply(est, function(e) e$este, character(1))),
                   columna_numerica(vapply(est, function(e) e$altitud, character(1))),
                   vapply(est, function(e) e$periodo, character(1)))
  for (j in seq_along(columnas)) {
    wb$add_data(sheet = hoja, x = columnas[[j]], start_col = j, start_row = 3, col_names = FALSE)
  }
  wb$add_cell_style(sheet = hoja, dims = wb_dims(rows = 3:(n + 2), cols = 1:ncol_info),
                    horizontal = "center")
  wb$set_col_widths(sheet = hoja, cols = 1:(ncol_info + 1),
                    widths = c(6, max(12, nchar(nombres)) + 2, 12, 14, 14, 18, 20, 3))

  # Tablas de dias y de meses con datos
  wb$set_col_widths(sheet = hoja, cols = col_ini1 + 0:(ancho - 1), widths = 3.78)
  wb$set_col_widths(sheet = hoja, cols = col_ini2 + 0:(ancho - 1), widths = 2.78)
  wb$set_row_heights(sheet = hoja, rows = c(fila_cal, fila_hid), heights = 45)

  tablas <- list(
    list(col_etq1, col_ini1, fila_cal, "calendario", "dias",  "Días con dato\nAño calendario"),
    list(col_etq1, col_ini1, fila_hid, "hidrologico", "dias", "Días con dato\nAño hidrológico"),
    list(col_etq2, col_ini2, fila_cal, "calendario", "meses", "Meses con dato\nAño calendario"),
    list(col_etq2, col_ini2, fila_hid, "hidrologico", "meses", "Meses con dato\nAño hidrológico"))
  for (t in tablas) {
    hid <- t[[4]] == "hidrologico"
    escribir_tabla_resumen(
      wb, hoja, fila = t[[3]], col_etq = t[[1]], col_ini = t[[2]], etq_filas = nombres,
      etq_anios = if (hid) etq_hid else etq_cal,
      decadas   = if (hid) dec_hid else dec_cal,
      valores   = matriz_resumen(est, if (hid) "hid" else "cal", if (t[[5]] == "dias") "dias" else "meses",
                                 if (hid) cols_hid else cols_cal),
      titulo = t[[6]], tipo = t[[5]])
  }

  # Leyenda del semaforo
  wb$add_data(sheet = hoja, x = c(12, 8, 5), start_col = col_ini2, start_row = fila_ley, col_names = FALSE)
  wb$add_data(sheet = hoja, start_col = col_ini2 + 1, start_row = fila_ley, col_names = FALSE,
              x = c(": Año completo", ": \u2265 7 meses con información",
                    ": < 7 meses con información", ": Sin información"))
  wb$add_conditional_formatting(sheet = hoja, dims = wb_dims(rows = fila_ley + 0:2, cols = col_ini2),
                                type = "iconSet", rule = c(0, 7, 12),
                                params = list(showValue = FALSE, iconSet = "3TrafficLights1"))
}

# =====================================================================
# 10. PROCESAMIENTO DE UNA PLANILLA
# =====================================================================
procesar_archivo <- function(file_xlsx, codigo) {

  cfg <- config_variables[[codigo]]

  # El archivo se lee una sola vez; las pestañas se extraen desde memoria
  wb_in         <- wb_load(file_xlsx)
  hojas         <- unname(wb_get_sheet_names(wb_in))
  df_estaciones <- wb_to_df(wb_in, sheet = "Estaciones")

  dir_variable <- file.path(dir_outputs, codigo)
  if (!dir.exists(dir_variable)) dir.create(dir_variable, recursive = TRUE)

  wb <- wb_workbook()
  wb$add_worksheet("Resumen")   # primera pestaña; se completa al final
  estaciones_resumen <- list()
  hojas_resumen <- character(0)  # pestañas efectivamente escritas en Resumen

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
    estaciones_resumen[[length(estaciones_resumen) + 1]] <-
      datos_resumen_estacion(hoja, nombre_estacion, datos, res_cal, res_hid)

    # --- Armado de tablas y disposicion ---
    tabla_orig <- armar_serie_original(datos$serie, datos$nombre_valor)
    tabla_cont <- armar_serie_continua(datos$serie, cfg, mes_inicio_hidro, hoja)
    tablas_cal <- armar_tablas_periodo(res_cal, cfg)
    tablas_hid <- armar_tablas_periodo(res_hid, cfg)
    pos        <- disponer_hoja(tabla_orig, tabla_cont, tablas_cal, tablas_hid)

    # --- Escritura ---
    wb$add_worksheet(hoja)
    hojas_resumen <- c(hojas_resumen, hoja)
    escribir_info(wb, hoja, datos$info, pos$info)
    escribir_serie_original(wb, hoja, tabla_orig, pos$serie)
    escribir_serie_continua(wb, hoja, tabla_cont, pos$continua)
    escribir_periodo(wb, hoja, res_cal, tablas_cal, cfg, pos$cal)
    escribir_periodo(wb, hoja, res_hid, tablas_hid, cfg, pos$hid)

    # --- Graficos ---
    graficar_periodo(wb, hoja, res_cal, tablas_cal, cfg, pos$cal, nombre_estacion)
    graficar_periodo(wb, hoja, res_hid, tablas_hid, cfg, pos$hid, nombre_estacion)
    graficar_serie_diaria(wb, hoja, cfg, tabla_cont, pos$continua, pos$grafico_diario,
                          nombre_estacion)
  }

  if (length(estaciones_resumen) > 0) escribir_resumen(wb, "Resumen", estaciones_resumen)

  # Sobrescribe solo el archivo de esta variable. Si esta abierto en Excel,
  # Windows no permite reemplazarlo: se avisa con un mensaje claro.
  archivo_salida <- paste0("Resumen_", codigo, ".xlsx")
  ruta_salida    <- file.path(dir_variable, archivo_salida)
  tryCatch(
    wb$save(ruta_salida, overwrite = TRUE),
    error = function(e) {
      stop("No se pudo guardar ", archivo_salida,
           ". Si esta abierto en Excel, cierrelo y vuelva a ejecutar el codigo.\n",
           "Detalle: ", conditionMessage(e), call. = FALSE)
    }
  )
  message("Archivo generado: ", archivo_salida)

  crear_criterios(codigo, hojas_resumen)
}

# =====================================================================
# 11. ITERACION SOBRE LAS PLANILLAS DE INPUTS
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

# Aviso si alguna variable de variables_a_procesar no tiene planilla en Inputs
codigos_en_inputs <- vapply(basename(archivos), identificar_variable, character(1))
sin_planilla <- setdiff(variables_a_procesar, codigos_en_inputs)
if (length(sin_planilla) > 0) {
  warning("Variables en variables_a_procesar sin planilla en Inputs: ",
          paste(sin_planilla, collapse = ", "), call. = FALSE)
}

message("El codigo ha terminado.")
