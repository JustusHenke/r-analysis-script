# Pakete laden
library(haven)

# Verzeichnis des Skripts setzen
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

# SPSS-Datei einlesen
# Pfad zur Rohdatendatei anpassen
data <- read_sav("Rohdaten/DATEINAME.sav")

# Value Labels anschauen (Beispiel)
# Variablenname anpassen
attr(data$VARIABLENNAME, "labels")

# Variable Label anschauen (Beispiel)
# Variablenname anpassen
attr(data$VARIABLENNAME, "label")

# RDS-Datei speichern (Labels bleiben erhalten)
# Ausgabedateiname anpassen
saveRDS(data, "Rohdaten/DATEINAME.rds")


# RDS-Datei einlesen
# Dateiname anpassen
d <- readRDS("Rohdaten/DATEINAME.rds")

# Alle Variablennamen
names(d)

# Alle Variablenlabels auflisten
sapply(d, function(x) attr(x, "label"))

# Beispiel: Value Labels für eine Variable
# Variablenname anpassen
attr(d$VARIABLENNAME, "labels")
