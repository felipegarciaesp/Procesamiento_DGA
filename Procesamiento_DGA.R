# =====================================================================
# Codigo para procesar los datos de estaciones DGA / Felipe Garcia

# Este codigo toma como inputs los outputs del codigo Datos_DGA
# =====================================================================

## Limpiar ambiente
#-------------------
rm(list=ls())
graphics.off()
cat("\014")

## Directorio de trabajo
#-------------------
setwd("C:/Codigos/Datos_DGA")
dir_inputs  <- file.path(getwd(), "Inputs")
dir_outputs <- file.path(getwd(), "Outputs")