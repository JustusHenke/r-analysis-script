library(haven)
library(dplyr)
library(writexl)

# Verzeichnis des Skripts setzen
try(setwd(dirname(rstudioapi::getActiveDocumentContext()$path)))

# =============================================================================
# KONFIGURATION
# =============================================================================

SPSS_FILE <- "PATH/FILE.sav"  # Pfad zur SPSS-Datei
CODEBOOK_EXCEL <- "codebook.xlsx"  # Codebook als Excel (für Dokumentation)

# =============================================================================
# CODEBOOK ERSTELLEN
# =============================================================================

cat("Lade SPSS-Daten...\n")
daten <- read_sav(SPSS_FILE)
cat("✓ Daten geladen:", nrow(daten), "Zeilen,", ncol(daten), "Spalten\n\n")

cat("Erstelle Codebook...\n")

# Variablenübersicht erstellen
codebook_df <- data.frame(
  Variable = names(daten),
  Label = sapply(daten, function(x) {
    label <- attr(x, "label")
    if (is.null(label)) "" else as.character(label)
  }),
  Typ = sapply(daten, function(x) paste(class(x), collapse = ", ")),
  stringsAsFactors = FALSE
)

# Wertelabels KORREKT extrahieren (Code = Label Format)
codebook_df$Wertelabels <- sapply(daten, function(x) {
  labels <- attr(x, "labels")
  if (!is.null(labels) && length(labels) > 0) {
    # WICHTIG: SPSS speichert als c("Weiblich" = "AO01")
    # Wir brauchen aber: c("AO01" = "Weiblich")
    
    label_names <- names(labels)  # "Weiblich", "Männlich", ...
    label_values <- as.character(labels)  # "AO01", "AO02", ...
    
    # Prüfe ob umgekehrt (Heuristik: Werte sind kürzer als Namen)
    avg_name_len <- mean(nchar(label_names))
    avg_value_len <- mean(nchar(label_values))
    
    if (avg_name_len > avg_value_len && avg_value_len <= 10) {
      # Umkehren: Code = Label
      paste(label_values, "=", label_names, collapse = "; ")
    } else {
      # Normal: verwende wie sie sind
      paste(label_names, "=", label_values, collapse = "; ")
    }
  } else {
    "keine"
  }
})

cat("✓ Codebook erstellt für", nrow(codebook_df), "Variablen\n\n")

# Zeige Beispiel
cat("Beispiel (erste 3 Variablen mit Labels):\n")
with_labels <- codebook_df[codebook_df$Wertelabels != "keine", ]
if (nrow(with_labels) > 0) {
  print(head(with_labels[, c("Variable", "Label", "Wertelabels")], 3))
}
cat("\n")

# =============================================================================
# CODEBOOK SPEICHERN
# =============================================================================

cat("Speichere Codebook...\n")

# Als Excel (für Dokumentation)
write_xlsx(codebook_df, CODEBOOK_EXCEL)
cat("✓ Excel gespeichert:", CODEBOOK_EXCEL, "\n")

cat("\n=== FERTIG ===\n")

