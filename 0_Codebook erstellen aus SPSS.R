library(haven)
library(dplyr)

# Verzeichnis des Skripts setzen
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

# Daten einlesen
daten <- read_sav("PATH/FILE.sav")

# Variablenübersicht erstellen
codebook_df <- data.frame(
  Variable = names(daten),
  Label = sapply(daten, function(x) {
    label <- attr(x, "label")
    if (is.null(label)) "" else label
  }),
  Typ = sapply(daten, function(x) paste(class(x), collapse = ", ")), 
  Wertelabels = sapply(daten, function(x) {
    labels <- attr(x, "labels")
    if (!is.null(labels)) {
      paste(names(labels), "=", labels, collapse = "; ")
    } else {
      "keine"
    }
  }),
  stringsAsFactors = FALSE
)


# Oder als Excel
library(writexl)
write_xlsx(codebook_df, "PATH/FILE.xlsx")