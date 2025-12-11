# Test der Log-Datei-Generierung
source("__AnalysisFunctions.R", local = TRUE)

# Simuliere globale Variablen wie im Cockpit
OUTPUT_FILE <- "test-output.xlsx"
LOG <- TRUE

# Rufe main() auf, aber nur den Logging-Teil
# Wir extrahieren den relevanten Teil
cat("=== Test Log-Datei-Generierung ===\n")
cat("OUTPUT_FILE:", OUTPUT_FILE, "\n")
cat("LOG:", LOG, "\n")

# Prüfe, ob LOG_FILE generiert wird
if (exists("LOG") && LOG) {
  if (!exists("LOG_FILE") || is.null(LOG_FILE) || LOG_FILE == "") {
    if (exists("OUTPUT_FILE") && !is.null(OUTPUT_FILE) && OUTPUT_FILE != "") {
      LOG_FILE <- sub("\\.xlsx$", ".log", OUTPUT_FILE)
      LOG_FILE <- sub("\\.log$", paste0("-", format(Sys.time(), "%Y%m%d-%H%M%S"), ".log"), LOG_FILE)
      cat("Generierte LOG_FILE:", LOG_FILE, "\n")
    }
  }
}

# Prüfe, ob setup_logging funktioniert
cat("\n=== Test setup_logging ===\n")
setup_logging("test-log.txt")
log_cat("Testnachricht\n")
close_logging()

cat("\n=== Test abgeschlossen ===\n")