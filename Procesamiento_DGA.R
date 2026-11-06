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

## Configuracion del año hidrologico
#-------------------
# Los meses se expresan como numero: 1 = ENE, 2 = FEB, 3 = MAR, 4 = ABR, ..., 12 = DIC
# El año hidrologico va desde el mes inicial del año X hasta el mes final del año X+1.
# Por defecto: 4 (ABR) del año X hasta 3 (MAR) del año X+1  ->  ejemplo: 1999/00
mes_inicio_hidro <- 4   # 4 = ABR (mes inicial, año X)
mes_fin_hidro    <- 3   # 3 = MAR (mes final, año X+1)

# Validacion: el mes final debe ser el mes anterior al mes inicial (12 meses en total)
if (mes_fin_hidro != ((mes_inicio_hidro + 10) %% 12) + 1) {
  stop("mes_fin_hidro debe ser el mes anterior a mes_inicio_hidro (el año hidrologico debe tener 12 meses).")
}

## Directorio de trabajo
#-------------------
setwd("C:/Codigos/Procesamiento_DGA")
dir_inputs  <- file.path(getwd(), "Inputs")
dir_outputs <- file.path(getwd(), "Outputs")

## Carpeta de salida (se recrea en cada ejecucion)
#-------------------
if (dir.exists(dir_outputs)) unlink(dir_outputs, recursive = TRUE)  # borra resultados anteriores
dir.create(dir_outputs, recursive = TRUE)

## Funciones auxiliares
#-------------------
# Clasifica la planilla segun su nombre: "PP", "TEMP" o NA
clasificar_variable <- function(nombre) {
  nombre <- tolower(nombre)
  if (grepl("^pp|precip", nombre))             return("PP")
  if (grepl("tmin|tmax|tmedia|temp", nombre))  return("TEMP")
  return(NA_character_)
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

# Escribe un titulo con celdas unidas, en negrita y centrado (fila 1)
escribir_titulo <- function(wb, hoja, texto, col_ini, col_fin) {
  wb$add_data(sheet = hoja, x = texto, start_col = col_ini, start_row = 1, col_names = FALSE)
  wb$merge_cells(sheet = hoja, dims = wb_dims(rows = 1, cols = col_ini:col_fin))
  wb$add_font(sheet = hoja, dims = wb_dims(rows = 1, cols = col_ini), bold = TRUE)
  wb$add_cell_style(sheet = hoja, dims = wb_dims(rows = 1, cols = col_ini), horizontal = "center")
}

# Escribe una tabla con encabezado en negrita en la fila 2
escribir_tabla <- function(wb, hoja, df, col_ini) {
  wb$add_data(sheet = hoja, x = df, start_col = col_ini, start_row = 2)
  wb$add_font(sheet = hoja, dims = wb_dims(rows = 2, cols = col_ini:(col_ini + ncol(df) - 1)),
              bold = TRUE)
}

# Escribe #N/A mediante la formula =NA()
escribir_na <- function(wb, hoja, fila, col) {
  wb$add_formula(sheet = hoja, x = "NA()", dims = wb_dims(rows = fila, cols = col))
}

## Procesamiento de precipitaciones
#-------------------
procesar_pp <- function(file_xlsx) {

  hojas <- unname(wb_get_sheet_names(wb_load(file_xlsx)))
  df_estaciones <- read_xlsx(file_xlsx, sheet = "Estaciones")

  # Rotulos y columnas de origen para la tabla de informacion
  info_rotulos  <- c("NOMBRE ESTACION", "CÓDIGO ESTACION", "REGIÓN",
                     "UTM ESTE (m)", "UTM NORTE (m)", "ALTITUD (m s.n.m.)")
  info_columnas <- c("NOMBRE.ESTACION", "CODIGO.ESTACION", "REGION",
                     "UTM_ESTE", "UTM_NORTE", "ALTITUD")

  nombres_meses <- c("ENE", "FEB", "MAR", "ABR", "MAY", "JUN",
                     "JUL", "AGO", "SEP", "OCT", "NOV", "DIC")

  wb <- wb_workbook()

  for (i in seq_len(nrow(df_estaciones))) {

    hoja <- df_estaciones$Pestana[i]

    if (!(hoja %in% hojas)) {
      message("La pestana ", hoja, " (", df_estaciones$Estacion[i],
              ") no existe en ", basename(file_xlsx))
      next
    }

    # check_names = TRUE: los espacios de los rotulos pasan a "." (como en openxlsx)
    df <- read_xlsx(file_xlsx, sheet = hoja, detect_dates = TRUE, check_names = TRUE)

    if (nrow(df) == 0) {
      message("La pestana ", hoja, " no tiene datos.")
      next
    }

    # Columnas de la serie
    col_fecha <- buscar_columna(df, "fecha", "FECHA")
    col_valor <- buscar_columna(df, "precip|^pp|valor", "precipitacion")
    col_unid  <- buscar_columna(df, "unidad", "unidad")

    wb$add_worksheet(hoja)

    # --- Tabla de informacion de la estacion (A1:B7) ---
    escribir_titulo(wb, hoja, "Informacion de la Estación", 1, 2)

    valores_info <- sapply(info_columnas, function(col) as.character(df[[col]][1]))
    wb$add_data(sheet = hoja, x = info_rotulos,  start_col = 1, start_row = 2, col_names = FALSE)
    wb$add_data(sheet = hoja, x = unname(valores_info), start_col = 2, start_row = 2, col_names = FALSE)

    # --- Serie diaria (D1:F...) ---
    serie <- data.frame(FECHA                = as.Date(df[[col_fecha]]),
                        Precipitación_diaria = df[[col_valor]],
                        Unidad               = df[[col_unid]],
                        stringsAsFactors     = FALSE)
    wb$add_data(sheet = hoja, x = serie, start_col = 4, start_row = 1)
    wb$add_font(sheet = hoja, dims = wb_dims(rows = 1, cols = 4:6), bold = TRUE)
    wb$add_numfmt(sheet = hoja, dims = wb_dims(rows = 2:(nrow(serie) + 1), cols = 4),
                  numfmt = "dd/mm/yyyy")

    wb$set_col_widths(sheet = hoja, cols = 1:6, widths = c(22, 28, 3, 12, 20, 10))

    # --- Cantidad de registros por mes - Año calendario (K1:W...) ---
    fechas  <- as.Date(serie$FECHA)
    validos <- !is.na(fechas) & !is.na(serie$Precipitación_diaria)
    anios   <- as.integer(format(fechas[validos], "%Y"))
    meses   <- as.integer(format(fechas[validos], "%m"))

    # Vector continuo de años: desde el primer hasta el ultimo año con datos
    anios_unicos <- seq(min(anios), max(anios))

    conteo <- table(factor(anios, levels = anios_unicos),
                    factor(meses, levels = 1:12))

    tabla_conteo <- data.frame(YEAR = anios_unicos,
                               matrix(as.integer(conteo), nrow = length(anios_unicos), ncol = 12))
    names(tabla_conteo) <- c("YEAR", nombres_meses)

    escribir_titulo(wb, hoja, "Cantidad de registros por mes - Año calendario", 11, 23)
    escribir_tabla(wb, hoja, tabla_conteo, 11)
    wb$set_col_widths(sheet = hoja, cols = 11:23, widths = c(8, rep(6, 12)))

    # Escala de colores: 0 = rojo, 15 = amarillo, 31 = verde oscuro
    wb$add_conditional_formatting(sheet = hoja,
                                  dims = wb_dims(rows = 3:(nrow(tabla_conteo) + 2), cols = 12:23),
                                  type = "colorScale",
                                  style = c("#F8696B", "#FFEB84", "#006400"),
                                  rule = c(0, 15, 31))

    # --- Cantidad de meses con información (Y1:Z...) ---
    tabla_meses <- data.frame(YEAR  = tabla_conteo$YEAR,
                              MESES = rowSums(tabla_conteo[, 2:13] > 0))

    escribir_titulo(wb, hoja, "Cantidad de meses con información", 25, 26)
    escribir_tabla(wb, hoja, tabla_meses, 25)
    wb$set_col_widths(sheet = hoja, cols = 25:26, widths = c(8, 8))

    # --- Precipitacion mensual (mm) - Año calendario (AB1:AN...) ---
    suma_mensual <- tapply(serie$Precipitación_diaria[validos],
                           list(factor(anios, levels = anios_unicos),
                                factor(meses, levels = 1:12)),
                           sum)

    tabla_pp <- data.frame(YEAR = anios_unicos,
                           matrix(as.numeric(suma_mensual),
                                  nrow = length(anios_unicos), ncol = 12))
    names(tabla_pp) <- c("YEAR", nombres_meses)

    escribir_titulo(wb, hoja, "Precipitacion mensual (mm) - Año calendario", 28, 40)
    escribir_tabla(wb, hoja, tabla_pp, 28)
    wb$set_col_widths(sheet = hoja, cols = 28:40, widths = c(8, rep(8, 12)))

    # Meses sin datos: formula =NA()
    pos_na <- which(is.na(tabla_pp[, 2:13]), arr.ind = TRUE)
    if (nrow(pos_na) > 0) {
      for (k in seq_len(nrow(pos_na))) {
        escribir_na(wb, hoja, 2 + pos_na[k, 1], 28 + pos_na[k, 2])
      }
    }

    # --- Pp mensual (mm) - Año calendario, en vector (AP1:AS...) ---
    n_anios  <- length(anios_unicos)
    vec_year <- rep(anios_unicos, each = 12)
    vec_mes  <- rep(1:12, times = n_anios)

    tabla_vec <- data.frame(FECHA     = as.Date(sprintf("%d-%02d-01", vec_year, vec_mes)),
                            MES       = vec_mes,
                            YEAR      = vec_year,
                            `Pp (mm)` = as.numeric(t(as.matrix(tabla_pp[, 2:13]))),
                            check.names = FALSE)

    escribir_titulo(wb, hoja, "Pp mensual (mm) - Año calendario", 42, 45)
    escribir_tabla(wb, hoja, tabla_vec, 42)
    wb$add_numfmt(sheet = hoja, dims = wb_dims(rows = 3:(nrow(tabla_vec) + 2), cols = 42),
                  numfmt = "dd/mm/yyyy")
    wb$set_col_widths(sheet = hoja, cols = 42:45, widths = c(12, 6, 8, 10))

    for (r in which(is.na(tabla_vec$`Pp (mm)`))) {
      escribir_na(wb, hoja, 2 + r, 45)
    }

    # --- Precipitacion promedio mensual (mm) - Año calendario (AU1:AV...) ---
    promedio_mensual <- sapply(2:13, function(j) {
      v <- tabla_pp[[j]]
      if (all(is.na(v))) NA_real_ else mean(v, na.rm = TRUE)
    })

    tabla_prom <- data.frame(MES = nombres_meses,
                             `Pp promedio (mm)` = promedio_mensual,
                             check.names = FALSE)

    escribir_titulo(wb, hoja, "Precipitacion promedio mensual (mm) - Año calendario", 47, 48)
    escribir_tabla(wb, hoja, tabla_prom, 47)

    for (m in which(is.na(promedio_mensual))) {
      escribir_na(wb, hoja, 2 + m, 48)
    }

    # Promedio anual (mm): suma de ENE a DIC ignorando NA (fila 15)
    promedio_anual <- sum(promedio_mensual, na.rm = TRUE)

    wb$add_data(sheet = hoja, x = "Promedio anual (mm)", start_col = 47, start_row = 15, col_names = FALSE)
    wb$add_data(sheet = hoja, x = promedio_anual, start_col = 48, start_row = 15, col_names = FALSE)
    wb$add_font(sheet = hoja, dims = wb_dims(rows = 15, cols = 47:48), bold = TRUE)
    wb$set_col_widths(sheet = hoja, cols = 47:48, widths = c(20, 16))

    # --- Pp anual y promedio anual (mm) - Año calendario (AX1:AZ...) ---
    pp_anual <- apply(tabla_pp[, 2:13], 1, function(v) {
      if (all(is.na(v))) NA_real_ else sum(v, na.rm = TRUE)
    })
    pp_prom_anual <- if (all(is.na(pp_anual))) NA_real_ else mean(pp_anual, na.rm = TRUE)

    tabla_anual <- data.frame(YEAR = anios_unicos,
                              `Pp anual (mm)` = as.numeric(pp_anual),
                              check.names = FALSE)

    escribir_titulo(wb, hoja, "Pp anual y promedio anual (mm) - Año calendario", 50, 52)
    escribir_tabla(wb, hoja, tabla_anual, 50)

    # Pp promedio anual (mm): encabezado y un unico valor en la primera fila de datos
    wb$add_data(sheet = hoja, x = "Pp promedio anual (mm)", start_col = 52, start_row = 2,
                col_names = FALSE)
    wb$add_font(sheet = hoja, dims = wb_dims(rows = 2, cols = 52), bold = TRUE)
    if (is.na(pp_prom_anual)) {
      escribir_na(wb, hoja, 3, 52)
    } else {
      wb$add_data(sheet = hoja, x = pp_prom_anual, start_col = 52, start_row = 3,
                  col_names = FALSE)
    }
    wb$set_col_widths(sheet = hoja, cols = 50:52, widths = c(8, 14, 22))

    for (r in which(is.na(tabla_anual[["Pp anual (mm)"]]))) {
      escribir_na(wb, hoja, 2 + r, 51)
    }
    if (is.na(pp_prom_anual)) {
      escribir_na(wb, hoja, 3, 52)
    }

    # =================================================================
    # Graficos (nativos de Excel, enlazados a las tablas de la pestaña)
    # =================================================================
    nombre_estacion <- df_estaciones$Estacion[i]
    fila_fin_anual  <- nrow(tabla_anual) + 2

    # Formatos de texto (Calibri) y grilla
    txt_titulo  <- officer::fp_text(font.size = 14, font.family = "Calibri")
    txt_rotulo  <- officer::fp_text(font.size = 10, font.family = "Calibri")
    txt_ejes    <- officer::fp_text(font.size = 9,  font.family = "Calibri")
    sin_linea   <- officer::fp_border(style = "none")
    linea_grilla <- officer::fp_border(color = "#D9D9D9", width = 0.75)

    # Aplica formato comun: sin leyenda, solo grilla horizontal mayor, sin titulo eje X
    formatear_grafico <- function(g, ylab, ymax) {
      g <- chart_labels(g, title = nombre_estacion, xlab = NULL, ylab = ylab)
      g <- chart_ax_y(g, limits = c(0, ymax))            # eje vertical desde 0
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

    maximo_nice <- function(v) {
      v <- v[!is.na(v)]
      if (length(v) == 0) return(1)
      max(pretty(c(0, max(v))))
    }

    # Grafico 1: Precipitacion media (columnas agrupadas), datos AU2:AV14 (sin Promedio anual)
    datos_g1 <- wb_data(wb, sheet = hoja, dims = "AU2:AV14")
    datos_g1$MES <- as.character(datos_g1$MES)
    g1 <- ms_barchart(data = datos_g1, x = "MES", y = "Pp promedio (mm)")
    g1 <- formatear_grafico(g1, "Precipitación media (mm)",
                            maximo_nice(datos_g1[["Pp promedio (mm)"]]))
    wb$add_mschart(sheet = hoja, dims = "BB2:BI17", graph = g1)

    # Grafico 2: Precipitacion anual (columnas agrupadas), datos AX2:AY...
    datos_g2 <- wb_data(wb, sheet = hoja, dims = paste0("AX2:AY", fila_fin_anual))
    datos_g2$YEAR <- as.character(datos_g2$YEAR)   # años como categorias (no 1, 2, 3...)
    g2 <- ms_barchart(data = datos_g2, x = "YEAR", y = "Pp anual (mm)")
    g2 <- formatear_grafico(g2, "Precipitación anual (mm)",
                            maximo_nice(datos_g2[["Pp anual (mm)"]]))
    g2 <- chart_ax_y(g2, limits = c(0, maximo_nice(datos_g2[["Pp anual (mm)"]])))  # minimo = 0
    g2 <- chart_ax_x(g2, rotate = -270)              # texto del eje X rotado 270°
    wb$add_mschart(sheet = hoja, dims = "BB19:BI34", graph = g2)
  }

  wb$save(file.path(dir_outputs, "Resumen_PP.xlsx"), overwrite = TRUE)
  message("Archivo generado: Resumen_PP.xlsx")
}

## Iteracion sobre las planillas de Inputs
#-------------------
archivos <- list.files(dir_inputs, pattern = "\\.xlsx$", full.names = TRUE)
archivos <- archivos[!grepl("^~\\$", basename(archivos))]  # ignora temporales de Excel

for (file_xlsx in archivos) {

  tipo <- clasificar_variable(basename(file_xlsx))

  if (is.na(tipo)) {
    message("No se pudo clasificar la planilla: ", basename(file_xlsx))
    next
  }

  if (tipo == "PP") {
    message("Procesando precipitaciones: ", basename(file_xlsx))
    procesar_pp(file_xlsx)
  } else {
    message("Planilla de temperatura omitida por ahora: ", basename(file_xlsx))
  }
}

message("El codigo ha terminado.")