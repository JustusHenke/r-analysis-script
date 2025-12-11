# =============================================================================
# ANALYSIS-COCKPIT - HAUPTEINSTIEGSPUNKT
# =============================================================================
# Version: 1.3.0
# Datum: 12.10.2025, 13:45 UTC
# Beschreibung: Hauptscript zur Steuerung der automatisierten Survey-Analyse
#               Lädt Konfiguration, Daten und ruft Analysefunktionen auf
# =============================================================================

library(dplyr)

try(setwd(dirname(rstudioapi::getActiveDocumentContext()$path)))


# =============================================================================
# Konfiguration
# =============================================================================


# Dateinamen
CONFIG_FILE <- "Analysis-Config.xlsx"
DATA_FILE <- "PATHTO/DATAFILE.rds"  # oder .csv oder .xlsx
OUTPUT_FILE <- "PATHTO/ANALYSISFILE.xlsx"
LOG <- TRUE  # Logging aktivieren/deaktivieren

SAVE_FINAL_DATASET <- TRUE  # oder FALSE
FINAL_DATASET_FILE <- "PATHTO/ANALYZEDDATAFILE.rds"


# Gewichtungseinstellungen
WEIGHTS <- FALSE  # TRUE oder FALSE
WEIGHT_VAR <- "weight"  # Name der Gewichtungsvariable im Datensatz

# Weitere Einstellungen
ALPHA_LEVEL <- 0.05
DIGITS_ROUND <- 2
INCLUDE_MISSING_DEFAULT <- FALSE


# =============================================================================
# Zusätzliche Variablen
# =============================================================================


# Metadaten-Variablen, die entfernt werden sollen
meta_vars_to_remove <- c("id", "lastpage", "seed", "submitdate", "startlanguage", "token")


add_custom_vars <- function(data) {
  # HIER KÖNNEN SIE EIGENE VARIABLEN ERSTELLEN
  # 
  # Beispiele:
  #
  # # Dichotome Variable aus kategorialer Variable
  # data <- data %>%
  #   mutate(
  #     geschlecht_maennl = case_when(
  #       geschlecht == "Männlich" ~ 1,
  #       geschlecht == "Weiblich" ~ 0,
  #       TRUE ~ NA_real_
  #     )
  #   )
  #
  # # Kategorisierung numerischer Variablen
  # data <- data %>%
  #   mutate(
  #     noten_kategorie = cut(
  #       note,
  #       breaks = c(-Inf, 1.5, 2.5, 3.5, 4.5, Inf),
  #       labels = c("sehr gut", "gut", "befriedigend", "ausreichend", "mangelhaft")
  #     )
  #   )
  #
  # # Umkodierung von Variablen
  # data <- data %>%
  #   mutate(
  #     region_ost_west = case_when(
  #       bundesland %in% c("BB", "MV", "SN", "ST", "TH") ~ "Ost",
  #       bundesland %in% c("BW", "BY", "HE", "NW", "RP", "SH", "SL", "NI") ~ "West",
  #       TRUE ~ "Unbekannt"
  #     )
  #   )
  #
  # # Gruppierungen mit group_by (z.B. Hochschul-Aggregationen)
  # data <- data %>%
  #   group_by(hochschul_id) %>%
  #   mutate(
  #     durchschnitt_hs = mean(zufriedenheit, na.rm = TRUE)
  #   ) %>%
  #   ungroup()
  
  return(data)
}

# Hilfsfunktion für Index-Definitionen
generate_index_definition <- function(name, label, prefix, range, binary = FALSE) {
  vars_original <- paste0(prefix, "[", sprintf("%03d", range), "]")
  list(
    name = name,
    label = label,
    vars_original = vars_original,
    binary = binary  
  )
}

# Indices bilden aus mehreren Matrix-Items
index_definitions <- list(
  # HIER KÖNNEN SIE INDIZES DEFINIEREN
  #
  # Beispiele:
  #
  # # Standard-Index aus Matrix-Items (Mittelwert)
  # generate_index_definition(
  #   name = "zufriedenheit_index", 
  #   label = "Zufriedenheits-Index", 
  #   prefix = "ZF01",  # Variablenname-Präfix
  #   range = 1:5       # Items ZF01[001] bis ZF01[005]
  # ),
  #
  # # Spezifische Items auswählen
  # generate_index_definition(
  #   name = "motivation_index",
  #   label = "Motivations-Index",
  #   prefix = "MO01",
  #   range = c(1,3,5,7)  # Nur Items 1,3,5,7
  # ),
  #
  # # Binäre Matrix (z.B. Checkbox-Grids)
  # list(
  #   name = "netzwerk_nutzung_index",
  #   label = "Netzwerk-Nutzungs-Index", 
  #   vars_original = paste0("NW01[", sprintf("%03d", 1:20), "_SQ002]"),
  #   binary = TRUE  # Wichtig für Checkbox-Grids!
  # )
)

# Gebildete Merkmale - Variable Labels
custom_var_labels <- c(
  # HIER KÖNNEN SIE LABELS FÜR IHRE VARIABLEN DEFINIEREN
  #
  # Format: variablenname = "Beschreibung"
  #
  # Beispiele:
  # alter = "Alter in Jahren",
  # geschlecht = "Geschlecht",
  # zufriedenheit_index = "Zufriedenheits-Index (Skala 1-5)"
)

# Value Labels für kategoriale Variablen
custom_val_labels <- list(
  # HIER KÖNNEN SIE WERTE-LABELS FÜR KATEGORIALE VARIABLEN DEFINIEREN
  #
  # Format: variablenname = c("wert1" = "Label1", "wert2" = "Label2")
  #
  # Beispiele:
  # geschlecht = c("1" = "Weiblich", "2" = "Männlich", "3" = "Divers"),
  # bildung = c("1" = "Hauptschule", "2" = "Realschule", "3" = "Abitur"),
  # zufriedenheit = c("1" = "Sehr unzufrieden", "2" = "Unzufrieden", 
  #                   "3" = "Neutral", "4" = "Zufrieden", "5" = "Sehr zufrieden")
)


# =============================================================================
# Analyse starten
# =============================================================================


# Extern hereinholen
source("__AnalysisFunctions.R")


# Script ausführen
if (interactive()) {
  # Nur ausführen wenn interaktiv (nicht beim Sourcen)
  results <- main()
} else {
  cat("Script geladen. Führen Sie main() aus, um die Analyse zu starten.\n")
}
