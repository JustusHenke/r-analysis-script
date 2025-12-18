# Pakete laden
library(haven)

# Verzeichnis des Skripts setzen
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

# SPSS-Datei einlesen
# Pfad zur Rohdatendatei anpassen
data <- read_sav("PATH/FILE.sav")


# RDS-Datei speichern (Labels bleiben erhalten)
# Ausgabedateiname anpassen
saveRDS(data, "PATH/FILE.rds")


# Alle Variablennamen
names(data)

# Alle Variablenlabels auflisten
sapply(data, function(x) attr(x, "label"))


