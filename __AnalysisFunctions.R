# =============================================================================
# SURVEY DATENAUSWERTUNG MIT KONFIGURIERBARER EXCEL-STEUERUNG
# =============================================================================
# Autor: Survey Analysis Script
# Version: 1.3.0
# Datum: 12.10.2025, 13:45 UTC
# Letzte Änderung: 51d1896 (Intelligent label truncation)
# Beschreibung: Automatisierte Auswertung von Survey-Daten basierend auf 
#               Excel-Konfiguration mit deskriptiven Statistiken, Kreuztabellen
#               und Regressionsanalysen
#
# CHANGELOG - Letzte Änderungen:
# ─────────────────────────────────────────────────────────────────────────
# v1.3.0 (12.10.2025)
#   • Matrix-Fragen erhalten nun auch statistische Tests 
#   • Filter-Funktionalität: Individuelle Filter für Variablen & Analysen
#   • Filter-Syntax: Einfache R/logische Ausdrücke (z.B. 'SD01==1')
#   • Sicherer Parser: parse_filter_expression() mit Validation
#   • Excel-Konfiguration: Filter-Spalte in allen Sheets
#   • Integration: Filter in deskriptive, Kreuztabellen, Regressionen, Textantworten
#   • Berichterstattung: Filter-Info in Excel-Export dokumentiert
#
# v1.2.0 (14.11.2025)
#   • Intelligent label truncation: Bessere Item-Labels für numerische Matrizen
#   • Duplicate frequency entries fixed: Korrekte Zählung mit which()
#   • Defensive checks: Leere result_rows vor rbind() Aufrufen
#   • Optional custom_var_labels: exists() Check für Crosstab-Kontext
#   • Label length control: Intelligentes Truncating bei >80 Zeichen
#
# v1.1.0 (Vorherige Phase)
#   • Redundancy elimination: 250 Zeilen duplizierter Code entfernt (74%)
#   • Helper functions: get_matrix_labels(), map_response_labels()
#   • Dokumentation: CRUSH.md, REFACTORING_SUMMARY.md
# ─────────────────────────────────────────────────────────────────────────

# =============================================================================
# LOGGING FUNKTIONEN
# =============================================================================

# Globale Log-Verbindung
log_connection <- NULL

setup_logging <- function(log_file) {
  "Startet das Logging mit dualer Ausgabe"
  
  # Verzeichnis erstellen falls nicht vorhanden
  dir.create(dirname(log_file), showWarnings = FALSE, recursive = TRUE)
  
  # Log-Verbindung öffnen
  log_connection <<- file(log_file, open = "wt", encoding = "UTF-8")
  
  cat("Logging gestartet:", log_file, "\n")
}

log_cat <- function(...) {
  "Schreibt sowohl in Console als auch Log-Datei"
  
  # Text zusammenfügen
  text <- paste(..., sep = "")
  
  # In Console ausgeben
  cat(text)
  
  # In Log-Datei schreiben
  if (!is.null(log_connection)) {
    writeLines(text, log_connection, sep = "")
    flush(log_connection)
  }
}

close_logging <- function() {
  "Schließt Log-Datei"
  
  if (!is.null(log_connection)) {
    close(log_connection)
    log_connection <<- NULL
  }
  
  cat("Log-Datei geschlossen.\n")
}

# =============================================================================
# PACKAGES LADEN
# =============================================================================

# Funktion zum sicheren Laden von Packages
load_packages <- function() {
  required_packages <- c(
    "readxl",      # Excel lesen
    "openxlsx",    # Excel schreiben mit Formatierung
    "dplyr",       # Datenmanipulation
    "tidyr",       # Datenumformung
    "stringr",     # String-Operationen
    "psych",       # Deskriptive Statistiken
    "survey",      # Gewichtete Analysen
    "haven",       # SPSS/Stata Files (für RDS mit Labels)
    "labelled",    # Label-Handling
    "lme4"         # Mehrebenenmodelle
  )
  
  
  # Prüfen welche Packages fehlen
  missing_packages <- required_packages[!required_packages %in% installed.packages()[,"Package"]]
  
  # Fehlende Packages installieren
  if(length(missing_packages) > 0) {
    cat("Installiere fehlende Packages:", paste(missing_packages, collapse = ", "), "\n")
    install.packages(missing_packages, dependencies = TRUE)
  }
  
  # Packages laden
  suppressMessages({
    lapply(required_packages, library, character.only = TRUE)
  })
  
  cat("Alle benötigten Packages erfolgreich geladen.\n")
}

# =============================================================================
# HILFSFUNKTIONEN
# =============================================================================



# Funktion um zu prüfen ob Datei existiert
check_file_exists <- function(filename) {
  if (!file.exists(filename)) {
    stop(paste("Datei nicht gefunden:", filename, "\nBitte prüfen Sie den Dateinamen und Pfad."))
  }
}

# Funktion um Dateiendung zu ermitteln
get_file_extension <- function(filename) {
  tolower(tools::file_ext(filename))
}

# =============================================================================
# FILTER FUNKTIONEN
# =============================================================================

parse_filter_expression <- function(filter_string, data_columns = NULL) {
  "Parst und validiert einen Filterausdruck für die Datenanalyse
  
  Args:
    filter_string: Filterausdruck als String (z.B. 'SD01==1', 'ALTER>=18')
    data_columns: Vektor von Variablennamen zur Validierung (optional)
  
  Returns:
    Parsed filter expression als expression object oder NULL bei leerem Filter
    Stoppt bei ungültigen Ausdrücken oder unbekannten Variablen
  "
  
  # Leerer Filter
  if (is.na(filter_string) || filter_string == "" || filter_string == "NA") {
    return(NULL)
  }
  
  cat("Parsing Filter:", filter_string, "\n")
  
  # Sicherheitsprüfung: Verbotene Funktionen
  forbidden_patterns <- c(
    "\\blibrary\\(",
    "\\bsource\\(",
    "\\bsystem\\(",
    "\\bfile\\(",
    "\\bread\\(",
    "\\bwrite\\(",
    "\\beval\\(",
    "\\bparse\\(",
    "\\bget\\(",
    "\\bsetwd\\(",
    "\\bdir\\(",
    "\\bsetwd\\("
  )
  
  for (pattern in forbidden_patterns) {
    if (grepl(pattern, filter_string, ignore.case = TRUE)) {
      stop("Filter enthält unerlaubte Funktion: ", filter_string)
    }
  }
  
  # Erlaubte Operatoren und Funktionen
  allowed_patterns <- c(
    # Operatoren
    "==", "!=", ">", "<", ">=", "<=", "&", "\\|", "!", "\\(", "\\)",
    # Funktionen
    "\\bis\\.na\\(",
    "\\bnchar\\(",
    "\\bgrepl\\(",
    "\\bsubstr\\(",
    "\\bstr_detect\\(",
    "\\b%in%",
    "\\bis\\.numeric\\(",
    "\\bis\\.character\\("
  )
  
  # Validierung der Syntax durch Parsing-Versuch
  parsed_expr <- tryCatch({
    parse(text = filter_string)
  }, error = function(e) {
    stop("Filter-Syntaxfehler in '", filter_string, "': ", e$message)
  })
  
  # Extrahiere alle Variablen aus dem Ausdruck
  variables_in_expr <- all.vars(parsed_expr)
  
  # Validiere Variablen gegen Daten (falls data_columns gegeben)
  if (!is.null(data_columns) && length(variables_in_expr) > 0) {
    invalid_vars <- variables_in_expr[!variables_in_expr %in% data_columns]
    if (length(invalid_vars) > 0) {
      stop("Unbekannte Variable(n) im Filter '", filter_string, "': ", 
           paste(invalid_vars, collapse = ", "))
    }
  }
  
  # Erlaubte Zeichen überprüfen (zusätzliche Sicherheit)
  allowed_chars <- "[a-zA-Z0-9._><=!&|()\"'\\[\\]\\s\\-+]"
  illegal_chars <- gsub(allowed_chars, "", filter_string)
  illegal_chars <- gsub(" ", "", illegal_chars)  # Leerzeichen ignorieren
  illegal_chars <- gsub("==", "", illegal_chars)  # Doppelgleich entfernen
  illegal_chars <- gsub("!=", "", illegal_chars)  # Ungleich entfernen
  illegal_chars <- gsub("<=", "", illegal_chars)  # <= entfernen
  illegal_chars <- gsub(">=", "", illegal_chars)  # >= entfernen
  
  if (nchar(illegal_chars) > 0) {
    warning("Potentiell unsichere Zeichen im Filter '", filter_string, "': '", 
            paste(unique(strsplit(illegal_chars, "")[[1]]), collapse = "', '"), "'")
  }
  
  return(parsed_expr)
}

apply_filter <- function(data, filter_expr) {
  "Wendet einen geparsten Filterausdruck auf einen Datensatz an
  
  Args:
    data: Datensatz (data.frame)
    filter_expr: Geparster Filterausdruck (von parse_filter_expression)
  
  Returns:
    Gefilterter Datensatz oder Original bei NULL
  "
  
  if (is.null(filter_expr)) {
    return(data)
  }
  
  tryCatch({
    # Evaluierung im Datensatz-Kontext
    filter_result <- eval(filter_expr, envir = data)
    
    # Sicherstellen, dass Ergebnis logisch ist
    if (!is.logical(filter_result)) {
      stop("Filter-Ausdruck muss logische Werte (TRUE/FALSE) zurückgeben")
    }
    
    if (length(filter_result) != nrow(data)) {
      stop("Filter-Ergebnis hat falsche Länge: ", length(filter_result), 
           " statt ", nrow(data))
    }
    
    # Datensatz filtern
    filtered_data <- data[filter_result & !is.na(filter_result), , drop = FALSE]
    
    cat("Filter angewendet: ", nrow(filtered_data), " von ", nrow(data), 
        " Fällen (", round(nrow(filtered_data)/nrow(data)*100, 1), "%)\n")
    
    return(filtered_data)
    
  }, error = function(e) {
    stop("Fehler beim Anwenden des Filters: ", e$message)
  })
}

validate_filter_config <- function(filter_string, variable_names) {
  "Validiert einen Filterausdruck für eine Konfiguration
  
  Args:
    filter_string: Filterausdruck als String
    variable_names: Verfügbare Variablennamen im Datensatz
  
  Returns:
    TRUE wenn Filter gültig, sonst Fehlermeldung
  "
  
  if (is.na(filter_string) || filter_string == "" || filter_string == "NA") {
    return(TRUE)
  }
  
  # Versuche zu parsen
  parsed_expr <- tryCatch({
    parse(text = filter_string)
  }, error = function(e) {
    return(paste("Syntaxfehler:", e$message))
  })
  
  # Extrahiere Variablen
  expr_vars <- tryCatch({
    all.vars(parsed_expr)
  }, error = function(e) {
    return(character(0))
  })
  
  # Prüfe auf unbekannte Variablen
  if (length(expr_vars) > 0) {
    unknown_vars <- expr_vars[!expr_vars %in% variable_names]
    if (length(unknown_vars) > 0) {
      return(paste("Unbekannte Variable(n):", paste(unknown_vars, collapse = ", ")))
    }
  }
  
  return(TRUE)
}

apply_row_filter <- function(data, config_row, variable_names = NULL) {
  "Wendet den Filter aus einer Config-Zeile auf den Datensatz an
  
  Args:
    data: Datensatz (data.frame)
    config_row: Einzelne Zeile aus Config (z.B. config$variablen[i,])
    variable_names: Optionale Liste verfügbarer Variablennamen zur Validierung
  
  Returns:
    Gefilterter Datensatz oder Original wenn kein Filter
  "
  
  # Filter-String extrahieren (kann 'filter' oder 'Filter' Spalte sein)
  filter_string <- NULL
  if ("filter" %in% names(config_row)) {
    filter_string <- config_row$filter[1]
  } else if ("Filter" %in% names(config_row)) {
    filter_string <- config_row$Filter[1]
  }
  
  # Kein Filter definiert
  if (is.null(filter_string) || is.na(filter_string) || filter_string == "") {
    return(list(
      data = data,
      filtered = FALSE,
      filter_info = NULL
    ))
  }
  
  # Validieren (falls variable_names gegeben)
  if (!is.null(variable_names)) {
    validation_result <- validate_filter_config(filter_string, variable_names)
    if (validation_result != TRUE) {
      warning("Ungültiger Filter '", filter_string, "': ", validation_result, ". Filter wird ignoriert.")
      return(list(
        data = data,
        filtered = FALSE,
        filter_info = NULL
      ))
    }
  }
  
  # Filter parsen und anwenden
  tryCatch({
    filter_expr <- parse_filter_expression(filter_string, variable_names)
    if (is.null(filter_expr)) {
      return(list(
        data = data,
        filtered = FALSE,
        filter_info = NULL
      ))
    }
    
    filtered_data <- apply_filter(data, filter_expr)
    
    return(list(
      data = filtered_data,
      filtered = TRUE,
      filter_info = list(
        filter_string = filter_string,
        original_n = nrow(data),
        filtered_n = nrow(filtered_data)
      )
    ))
    
  }, error = function(e) {
    warning("Fehler beim Anwenden des Filters '", filter_string, "': ", e$message, ". Filter wird ignoriert.")
    return(list(
      data = data,
      filtered = FALSE,
      filter_info = NULL
    ))
  })
}

# =============================================================================
# KONFIGURATION LADEN UND VALIDIEREN
# =============================================================================

load_config <- function() {
  cat("Lade Konfiguration aus:", CONFIG_FILE, "\n")
  
  # Prüfen ob Konfig-Datei existiert
  check_file_exists(CONFIG_FILE)
  
  # Alle Sheets laden
  sheet_names <- excel_sheets(CONFIG_FILE)
  cat("Gefundene Sheets:", paste(sheet_names, collapse = ", "), "\n")
  
  config <- list()
  
  # Sheet 1: Variablen (obligatorisch)
  if (!"Variablen" %in% sheet_names) {
    stop("Sheet 'Variablen' fehlt in der Konfigurationsdatei!")
  }
  
  config$variablen <- read_excel(CONFIG_FILE, sheet = "Variablen", col_types = "text") %>%
    mutate(
      # Konvertiere logische Spalten
      reverse_coding = as.logical(reverse_coding),
      use_NA = if("use_NA" %in% names(.)) as.logical(use_NA) else INCLUDE_MISSING_DEFAULT,
      # Konvertiere numerische Spalten
      min_value = as.numeric(min_value),
      max_value = as.numeric(max_value),
      # Filter-Spalte (optional) als String behalten
      filter = if("filter" %in% names(.)) filter else NA_character_
    ) %>%
    # Entferne leere Zeilen
    filter(!is.na(variable_name) & variable_name != "")
  
  # Validierung der Variablen-Konfiguration
  validate_variable_config(config$variablen)
  
  # Sheet 2: Kreuztabellen (optional)
  if ("Kreuztabellen" %in% sheet_names) {
    config$kreuztabellen <- read_excel(CONFIG_FILE, sheet = "Kreuztabellen", col_types = "text") %>%
      mutate(
        filter = if("filter" %in% names(.)) filter else NA_character_
      ) %>%
      filter(!is.na(analysis_name) & analysis_name != "")
    cat("Kreuztabellen-Konfiguration geladen:", nrow(config$kreuztabellen), "Analysen\n")
  } else {
    config$kreuztabellen <- data.frame()
    cat("Keine Kreuztabellen-Konfiguration gefunden (optional)\n")
  }
  
  # Sheet 3: Regressionen (optional)
  if ("Regressionen" %in% sheet_names) {
    config$regressionen <- read_excel(CONFIG_FILE, sheet = "Regressionen", col_types = "text") %>%
      mutate(
        filter = if("filter" %in% names(.)) filter else NA_character_
      ) %>%
      filter(!is.na(regression_name) & regression_name != "")
    cat("Regressions-Konfiguration geladen:", nrow(config$regressionen), "Analysen\n")
  } else {
    config$regressionen <- data.frame()
    cat("Keine Regressions-Konfiguration gefunden (optional)\n")
  }
  
  # Sheet 4: Textantworten (optional) - NEUE ERGÄNZUNG
  if ("Textantworten" %in% sheet_names) {
    config$textantworten <- read_excel(CONFIG_FILE, sheet = "Textantworten", col_types = "text") %>%
      mutate(
        min_length = as.numeric(ifelse(is.na(min_length), 3, min_length)),
        include_empty = as.logical(ifelse(is.na(include_empty), FALSE, include_empty))
      ) %>%
      mutate(
        filter = if("filter" %in% names(.)) filter else NA_character_
      ) %>%
      filter(!is.na(analysis_name) & analysis_name != "")
    cat("Textantworten-Konfiguration geladen:", nrow(config$textantworten), "Analysen\n")
  } else {
    config$textantworten <- data.frame()
    cat("Keine Textantworten-Konfiguration gefunden (optional)\n")
  }
  
  cat("Konfiguration erfolgreich geladen.\n")
  cat("Anzahl Variablen:", nrow(config$variablen), "\n")
  
  return(config)
}

# Validierung der Variablen-Konfiguration
validate_variable_config <- function(variablen_config) {
  required_cols <- c("variable_name", "data_type")
  missing_cols <- required_cols[!required_cols %in% names(variablen_config)]
  
  if (length(missing_cols) > 0) {
    stop(paste("Fehlende Spalten in Variablen-Sheet:", paste(missing_cols, collapse = ", ")))
  }
  
  # Prüfe gültige data_type Werte
  valid_types <- c("numeric", "nominal_coded", "nominal", "nominal_text", "ordinal", "dichotom", "matrix")
  invalid_types <- variablen_config$data_type[!variablen_config$data_type %in% valid_types]
  
  if (length(invalid_types) > 0) {
    stop(paste("Ungültige data_type Werte gefunden:", paste(unique(invalid_types), collapse = ", "),
               "\nGültige Werte:", paste(valid_types, collapse = ", ")))
  }
  
  # Prüfe ob coding bei nominal_coded/ordinal vorhanden
  needs_coding <- variablen_config %>%
    filter(data_type %in% c("nominal_coded", "ordinal")) %>%
    filter(is.na(coding) | coding == "")
  
  if (nrow(needs_coding) > 0) {
    warning(paste("Variablen ohne coding gefunden:", 
                  paste(needs_coding$variable_name, collapse = ", ")))
  }
  
  cat("Variablen-Konfiguration validiert.\n")
}


# =============================================================================
# VEREINFACHTE MATRIX-BEHANDLUNG
# =============================================================================

extract_numeric_from_matrix_coding <- function(data_values, coding_string, min_value = NA, max_value = NA) {
  "Extrahiert numerische Werte aus Matrix-Items basierend auf Kodierung und filtert nach min/max"
  
  if (is.na(coding_string) || coding_string == "") {
    return(as.numeric(data_values))
  }
  
  # cat("DEBUG: Parsing coding:", coding_string, "\n")
  if (!is.na(min_value) && !is.na(max_value)) {
    # cat("DEBUG: Gültiger Wertebereich:", min_value, "bis", max_value, "\n")
  }
  
  # Parse coding - verbesserte Version für Matrix-Format
  labels <- parse_coding(coding_string)
  
  if (is.null(labels) || length(labels) == 0) {
    # cat("DEBUG: Keine Labels gefunden, versuche direkte Konvertierung\n")
    return(as.numeric(data_values))
  }
  
  # cat("DEBUG: Gefundene Labels:", paste(names(labels), "=", labels, collapse = ", "), "\n")
  # cat("DEBUG: Beispiel data_values:", paste(head(data_values, 10), collapse = ", "), "\n")
  # 
  # Konvertiere Werte basierend auf Kodierung
  numeric_values <- rep(NA, length(data_values))
  
  for (i in seq_along(data_values)) {
    if (!is.na(data_values[i])) {
      current_value <- as.character(data_values[i])
      extracted_number <- NA
      
      # *** NEUE STRATEGIE 1: AO01, AO02 Pattern direkt verarbeiten ***
      if (grepl("^AO\\d+$", current_value)) {
        # Extrahiere Nummer aus AO01 -> 01 -> 1
        ao_number <- gsub("^AO0*", "", current_value)  # Entferne AO und führende Nullen
        if (ao_number != "" && !is.na(suppressWarnings(as.numeric(ao_number)))) {
          extracted_number <- as.numeric(ao_number)
        }
      }
      
      # *** NEUE STRATEGIE 2: Suche nach AO-Pattern in Labels ***
      if (is.na(extracted_number) && current_value %in% names(labels)) {
        # Extrahiere numerischen Code aus dem Key
        ao_match <- gsub("^AO0*", "", current_value)
        if (ao_match != "" && !is.na(suppressWarnings(as.numeric(ao_match)))) {
          extracted_number <- as.numeric(ao_match)
        }
      }
      
      # Strategie 3: Direkte Label-Übereinstimmung (bestehende Logik)
      if (is.na(extracted_number)) {
        matching_code <- names(labels)[labels == current_value]
        if (length(matching_code) > 0) {
          extracted_number <- as.numeric(matching_code[1])
        }
      }
      
      # Strategie 4: Suche nach numerischem Präfix im data_value
      if (is.na(extracted_number)) {
        numeric_match <- str_extract(current_value, "^\\d+")
        if (!is.na(numeric_match) && numeric_match %in% names(labels)) {
          extracted_number <- as.numeric(numeric_match)
        }
      }
      
      # Strategie 5: Fallback - direkte Konvertierung
      if (is.na(extracted_number)) {
        extracted_number <- suppressWarnings(as.numeric(current_value))
      }
      
      # *** NEUE VALIDIERUNG: Prüfe min/max Bereich ***
      if (!is.na(extracted_number)) {
        if (!is.na(min_value) && extracted_number < min_value) {
          cat("DEBUG: Wert", extracted_number, "unter Minimum", min_value, "-> auf NA gesetzt\n")
          extracted_number <- NA
        } else if (!is.na(max_value) && extracted_number > max_value) {
          cat("DEBUG: Wert", extracted_number, "über Maximum", max_value, "-> auf NA gesetzt\n")
          extracted_number <- NA
        }
      }
      
      numeric_values[i] <- extracted_number
    }
  }
  
  successful_conversions <- sum(!is.na(numeric_values))
  total_values <- length(data_values[!is.na(data_values)])
  # cat("DEBUG: Erfolgreich konvertiert:", successful_conversions, "von", total_values, "Werten\n")
  
  return(numeric_values)
}

# NEUE HILFSFUNKTION: Verbesserte Kodierung-Parser
parse_coding <- function(coding_string) {
  "Verbesserte Parsing-Funktion für Matrix-Kodierungen - unterstützt auch 'Zahl (Text)' Format"
  
  if (is.na(coding_string) || coding_string == "") {
    return(NULL)
  }
  
  # Split bei Semikolon und trim whitespace
  parts <- str_split(coding_string, ";")[[1]]
  parts <- str_trim(parts)
  
  labels <- c()
  
  for (part in parts) {
    if (str_detect(part, "=")) {
      # Format: "1=Label" oder "1 = Label"
      split_part <- str_split(part, "=", n = 2)[[1]]
      if (length(split_part) == 2) {
        code <- str_trim(split_part[1])
        label <- str_trim(split_part[2])
        labels[code] <- label
      }
    } else if (str_detect(part, "^\\d+\\s*\\(")) {
      # Format: "1 (sehr unzufrieden)" oder "1(Label)" oder "5 (stimme voll und ganz zu)"
      code_match <- str_extract(part, "^\\d+")
      label_match <- str_extract(part, "\\(([^)]+)\\)")
      
      if (!is.na(code_match) && !is.na(label_match)) {
        # Entferne Klammern vom Label
        label_clean <- str_remove_all(label_match, "\\(|\\)")
        labels[code_match] <- label_clean
      }
    } else if (str_detect(part, "^\\d+")) {
      # Format: "1 Label ohne Klammern"
      tokens <- str_split(part, "\\s+", n = 2)[[1]]
      if (length(tokens) >= 2) {
        code <- tokens[1]
        label <- paste(tokens[-1], collapse = " ")
        labels[code] <- label
      }
    }
  }
  
  # *** NEUER FALLBACK für Daten-Werte ***
  # Falls keine Labels aus Config gefunden, versuche aus den Datenwerten zu extrahieren
  if (length(labels) == 0) {
    # Prüfe ob coding_string einzelne Werte im "Zahl (Text)" Format enthält
    unique_values <- unique(str_trim(parts))
    
    for (value in unique_values) {
      if (str_detect(value, "^\\d+\\s*\\(")) {
        # Extrahiere "5 (stimme voll und ganz zu)" -> code="5", label="stimme voll und ganz zu"
        code_match <- str_extract(value, "^\\d+")
        label_match <- str_extract(value, "\\(([^)]+)\\)")
        
        if (!is.na(code_match) && !is.na(label_match)) {
          label_clean <- str_remove_all(label_match, "\\(|\\)")
          labels[code_match] <- label_clean
        }
      }
    }
  }
  
  # *** NEUER CODE: Sortiere Labels nach ordinaler Logik ***
  if (length(labels) > 0) {
    # Extrahiere numerische Codes und sortiere danach
    label_codes <- names(labels)
    numeric_codes <- suppressWarnings(as.numeric(label_codes))
    
    # Falls numerische Codes vorhanden sind, sortiere nach diesen
    if (!any(is.na(numeric_codes))) {
      sorted_indices <- order(numeric_codes)
      labels <- labels[sorted_indices]
      cat("  Labels nach numerischen Codes sortiert\n")
    } else {
      # Fallback: Sortiere nach bekannten ordinalen Mustern in den Label-Texten
      label_values <- as.character(labels)
      sorted_values <- sort_response_categories(label_values)
      
      # Rekonstruiere die Code-zu-Label Zuordnung in der sortierten Reihenfolge
      sorted_labels <- c()
      for (sorted_value in sorted_values) {
        # Finde den entsprechenden Code für diesen Label-Wert
        matching_code <- names(labels)[labels == sorted_value][1]
        if (!is.na(matching_code)) {
          sorted_labels[matching_code] <- sorted_value
        }
      }
      if (length(sorted_labels) > 0) {
        labels <- sorted_labels
        cat("  Labels nach ordinalen Mustern sortiert\n")
      }
    }
  }
  
  return(labels)
}


create_matrix_table <- function(data, var_config, use_na, survey_obj = NULL) {
  matrix_name <- var_config$variable_name
  question_text <- var_config$question_text
  
  cat("💫 Verarbeite Matrix:", matrix_name, "\n")
  
  # Finde alle Matrix-Items mit verschiedenen Trennern
  matrix_patterns <- c(
    paste0("^", matrix_name, "\\[.+\\]$"),     # Original: ZS01[001]
    paste0("^", matrix_name, "\\..+\\.$"),     # Sanitized: ZS01.001.
    paste0("^", matrix_name, "_.+$"),          # Underscore: ZS01_001
    paste0("^", matrix_name, "-.+$")           # Dash: ZS01-001
  )
  
  matrix_vars <- c()
  for (pattern in matrix_patterns) {
    found_vars <- names(data)[grepl(pattern, names(data))]
    matrix_vars <- c(matrix_vars, found_vars)
  }
  
  # FILTER OUT [other] variables
  matrix_vars <- matrix_vars[!grepl("other", matrix_vars, ignore.case = TRUE)]
  
  # Duplikate entfernen und sortieren
  matrix_vars <- unique(matrix_vars)
  matrix_vars <- sort(matrix_vars)
  
  if (length(matrix_vars) == 0) {
    cat("WARNUNG: Keine Matrix-Items gefunden für", matrix_name, "\n")
    cat("Gesucht nach Mustern:", paste(matrix_patterns, collapse = ", "), "\n")
    cat("Verfügbare Variablen mit", matrix_name, ":", 
        paste(names(data)[grepl(matrix_name, names(data))], collapse = ", "), "\n")
    
    # Returniere NULL statt Fehler zu werfen
    return(NULL)
  }
  
  cat("Gefundene Matrix-Items:", length(matrix_vars), "\n")
  cat("Items:", paste(matrix_vars, collapse = ", "), "\n")
  
  # SCHRITT 1: Alle möglichen Antwortkategorien sammeln
  all_responses <- c()
  for (var in matrix_vars) {
    # IMMER die tatsächlichen Werte sammeln (unabhängig von Labels)
    if (!use_na) {
      var_responses <- unique(data[[var]][!is.na(data[[var]])])
    } else {
      var_responses <- unique(data[[var]])
    }
    all_responses <- c(all_responses, var_responses)
  }
  
  # Eindeutige Kategorien ermitteln und sortieren
  unique_responses <- unique(all_responses)
  unique_responses <- unique_responses[!is.na(unique_responses)]
  # NEUER FIX: Entferne leere Strings
  unique_responses <- unique_responses[unique_responses != "" & !is.null(unique_responses)]
  
  # Sortierung: Versuche intelligente Sortierung für ordinale Daten
  unique_responses <- sort_response_categories(unique_responses)
  
  cat("Gefundene Antwortkategorien:", paste(unique_responses, collapse = ", "), "\n")
  
  # *** VEREINFACHTE LABEL-EXTRAKTION MIT NEUER HILFSFUNKTION ***
  # Hole Labels mit zentraler Funktion (eliminiert Redundanz)
  labels <- get_matrix_labels(data, matrix_vars, matrix_name, var_config, var_config$coding)
  
  if (!is.null(labels) && length(labels) > 0) {
    cat("Labels für Matrix-Responses gefunden:", length(labels), "Labels\n")
  } else {
    cat("Keine Labels für Matrix-Responses gefunden\n")
  }
  
  # Mappe Response-Werte auf Labels (zentralisierte Funktion)
  response_labels <- map_response_labels(unique_responses, labels, verbose = TRUE)
  
  
  
  # Bestimme Matrix-Typ basierend auf Kodierung und Daten
  has_coding <- !is.na(var_config$coding) && var_config$coding != ""
  is_dichotomous_matrix <- FALSE
  is_ordinal_matrix <- FALSE
  is_numeric_matrix <- FALSE
  
  # Prüfe ob dichotome Matrix (nur "1" oder leer in Daten)
  if (all(unique_responses %in% c("", "1")) || all(unique_responses %in% c("1"))) {
    is_dichotomous_matrix <- TRUE
    cat("Dichotome Matrix erkannt (nur '1' Werte)\n")
  }
  
  # Prüfe ob ordinale Matrix (Kodierung vorhanden und nicht dichotom)
  if (has_coding && !is_dichotomous_matrix) {
    is_ordinal_matrix <- TRUE
    cat("Ordinale Matrix erkannt (Kodierung vorhanden)\n")
  }
  
  # Prüfe ob numerische Matrix (keine Kodierung, numerische Werte)
  if (!has_coding && !is_dichotomous_matrix) {
    test_numeric <- suppressWarnings(as.numeric(unique_responses))
    if (all(!is.na(test_numeric))) {
      is_numeric_matrix <- TRUE
      cat("Numerische Matrix erkannt (numerische Werte ohne Kodierung)\n")
    }
  }
  
  # Numerische Statistiken
  
  # Prüfe ob Kodierung verfügbar ist
  if (!is.na(var_config$coding) && var_config$coding != "") {
    # *** VEREINFACHTE LABEL-EXTRAKTION ***
    labels <- get_matrix_labels(data, matrix_vars, matrix_name, var_config, var_config$coding)
    
    if (!is.null(labels) && length(labels) > 0) {
      cat("Kodierung gefunden:", paste(names(labels), "=", labels, collapse = ", "), "\n")
    }
    
    # Erkenne dichotome Matrix (Y/N, 1/0, etc.)
    if (!is.null(labels) && length(labels) <= 3) {  # Max 3 Kategorien für dichotom
      label_keys <- names(labels)
      label_values <- unique_responses  # Tatsächlich vorhandene Werte
      
      # Pattern 1: Y/N in Kodierung
      if (any(c("Y", "N") %in% label_keys) || any(c("1", "0") %in% label_keys)) {
        is_dichotomous_matrix <- TRUE
        cat("Dichotome Matrix erkannt (Y/N Pattern in Kodierung)\n")
      }
      # Pattern 2: Nur "1" und leere Werte in Daten (typisch für Checkboxen)
      else if (all(label_values %in% c("", "1")) || all(label_values %in% c("1"))) {
        is_dichotomous_matrix <- TRUE
        cat("Dichotome Matrix erkannt (1/leer Pattern in Daten)\n")
      }
    }
    
    # *** VEREINFACHTES LABEL-MAPPING ***
    if (!is.null(labels) && length(labels) > 0) {
      # Verwende zentralisierte Mapping-Funktion
      response_labels <- map_response_labels(unique_responses, labels, verbose = FALSE)
    }
  }
  # *** ENDE NEUER CODE ***
  
  # SCHRITT 2: Dynamische Spalten für alle Kategorien erstellen
  result_rows <- list()
  
  # SPEZIELLE BEHANDLUNG FÜR DICHOTOME MATRIX
  if (is_dichotomous_matrix) {
    cat("Erstelle kategoriale Tabelle für dichotome Matrix\n")
    
    # WICHTIG: Verwende ALLE Fälle für Gesamtzahl, nicht nur gefilterte
    total_n <- nrow(data)  # Gesamtstichprobe statt gefilterte Daten
    
    for (var in matrix_vars) {
      # Label extrahieren - VERWENDE GLEICHE LOGIK WIE BEI ORDINALEN MATRIX
      item_label <- extract_item_label(data, var, matrix_name)
      
      # FALLBACK: Falls Label schlecht ist, versuche bessere Extraktion
      if (grepl("^(Item|Item:|Subquestion)", item_label)) {
        # Suche nach besseren Labels in custom_var_labels oder Attributen
        if (var %in% names(custom_var_labels) && !is.na(custom_var_labels[[var]])) {
          item_label <- custom_var_labels[[var]]
        } else {
          # Verbessertes Fallback-Label
          clean_part <- gsub(paste0("^", matrix_name, "[._\\[\\]-]*"), "", var)
          clean_part <- gsub("[._\\]]*$", "", clean_part)
          if (clean_part != "" && clean_part != var) {
            item_label <- clean_part
          }
        }
      }
      
      cat("Variable:", var, "-> Label:", item_label, "\n")
      
      # Für dichotome Matrix: Zähle 1en und 0en/leere GEGEN GESAMTSTICHPROBE
      var_data <- data[[var]]
      
      # Zähle explizit
      count_1 <- sum(var_data == "1", na.rm = TRUE)
      count_0_or_empty <- total_n - count_1  # Alle anderen sind "nicht ausgewählt"
      
      cat("  Counts für", var, ": 1er =", count_1, ", 0/leer =", count_0_or_empty, ", total =", total_n, "\n")
      
      # Ergebnis-Zeile für dichotome Matrix (KORRIGIERT)
      result_row <- data.frame(
        Item = item_label,
        Ausgewählt_absolut = count_1,
        Nicht_ausgewählt_absolut = count_0_or_empty,
        Ausgewählt_prozent = round(count_1 / total_n * 100, DIGITS_ROUND),
        Nicht_ausgewählt_prozent = round(count_0_or_empty / total_n * 100, DIGITS_ROUND),
        Gesamt = total_n,
        stringsAsFactors = FALSE
      )
      
      result_rows[[var]] <- result_row
    }
  } else {
    cat("Erstelle kategoriale Tabelle für normale Matrix\n")
    
    for (var in matrix_vars) {
      # Label extrahieren
      item_label <- extract_item_label(data, var, matrix_name)
      cat("Variable:", var, "-> Label:", item_label, "\n")
      
      # Daten für dieses Item filtern
      if (!use_na) {
        item_data <- data[!is.na(data[[var]]), ]
      } else {
        item_data <- data
      }
      
      # Häufigkeiten berechnen
      if (!is.null(survey_obj) && WEIGHTS) {
        # Gewichtete Häufigkeiten
        if (!use_na) {
          survey_obj_filtered <- subset(survey_obj, !is.na(get(var)))
        } else {
          survey_obj_filtered <- survey_obj
        }
        
        freq_table <- svytable(as.formula(paste("~", var)), survey_obj_filtered)
        freq_df <- data.frame(
          response = names(freq_table),
          count = as.numeric(freq_table),
          stringsAsFactors = FALSE
        )
      } else {
        # Ungewichtete Häufigkeiten
        freq_table <- table(item_data[[var]], useNA = if(use_na) "always" else "no")
        freq_df <- data.frame(
          response = names(freq_table),
          count = as.numeric(freq_table),
          stringsAsFactors = FALSE
        )
      }
      
      # Gesamtzahl für Prozente
      total_count <- sum(freq_df$count)
      
      # Ergebnis-Zeile initialisieren
      result_row <- data.frame(Item = item_label, stringsAsFactors = FALSE)
      
      # SCHRITT 1: Erst alle absoluten Werte sammeln
      absolut_values <- list()
      prozent_values <- list()
      
      for (response in unique_responses) {
        # *** SICHER: Extrahiere COUNT mit which() um Duplikate zu handlen ***
        matching_idx <- which(freq_df$response == response)
        
        if (length(matching_idx) == 0) {
          count <- 0
        } else if (length(matching_idx) == 1) {
          count <- freq_df$count[matching_idx]
        } else {
          # FEHLERFALL: Mehrere Häufigkeitszeilen für gleichen Response
          # Summiere sie (sollte nicht vorkommen, aber defensiv)
          cat("    WARNUNG: Mehrere Häufigkeitszeilen für Response '", response, "' - summiere\n")
          count <- sum(freq_df$count[matching_idx])
        }
        
        percent <- if (total_count > 0) round(count / total_count * 100, DIGITS_ROUND) else 0
        
        # *** SICHERERE INDEXIERUNG: Verwende which() oder [1] um sicherzustellen, dass genau 1 Wert zurückkommt ***
        response_char <- as.character(response)
        response_label <- NA_character_
        
        # Versuche direkten Match in response_labels names
        label_matching_idx <- which(names(response_labels) == response_char)
        if (length(label_matching_idx) > 0) {
          response_label <- response_labels[label_matching_idx[1]]
        } else {
          # Fallback zu rauem Wert
          response_label <- response_char
        }
        
        clean_response <- make_clean_colname(response_label)
        
        # Werte sammeln statt direkt zuweisen
        absolut_values[[paste0(clean_response, "_absolut")]] <- count
        prozent_values[[paste0(clean_response, "_prozent")]] <- percent
      }
      
      # SCHRITT 2: Erst alle absoluten Spalten hinzufügen
      for (col_name in names(absolut_values)) {
        result_row[[col_name]] <- absolut_values[[col_name]]
      }
      
      # SCHRITT 3: Dann alle Prozent-Spalten hinzufügen  
      for (col_name in names(prozent_values)) {
        result_row[[col_name]] <- prozent_values[[col_name]]
      }
      
      # N/A Spalten hinzufügen wenn use_na = TRUE
      if (use_na) {
        na_count <- freq_df$count[is.na(freq_df$response)]
        if (length(na_count) == 0) na_count <- 0
        na_percent <- if (total_count > 0) round(na_count / total_count * 100, DIGITS_ROUND) else 0
        
        result_row$NA_absolut <- na_count
        result_row$NA_prozent <- na_percent
      }
      
      result_row$Gesamt <- total_count
      
      result_rows[[var]] <- result_row
      cat("    ✓ Zeile hinzugefügt für", var, "- Spalten:", ncol(result_row), "\n")
    }
  }
  
  # Alle Zeilen zusammenfügen
  # *** SICHER: Prüfe ob result_rows leer ist ***
  if (length(result_rows) == 0) {
    cat("WARNUNG: Keine Zeilen für kategoriale Tabelle erstellt!\n")
    result_table <- data.frame()  # Leere Tabelle
  } else {
    result_table <- do.call(rbind, result_rows)
    rownames(result_table) <- NULL
  }
  
  # PRÜFE OB KODIERUNG VORHANDEN IST (ordinal behandeln) ODER DICHOTOM ERKANNT
  has_coding <- !is.na(var_config$coding) && var_config$coding != ""
  
  # NEUE LOGIK: Erkenne ordinale Matrix basierend auf Kodierung
  is_ordinal_matrix <- FALSE
  if (has_coding) {
    labels <- parse_coding(var_config$coding)  
    if (!is.null(labels) && length(labels) > 2) {
      # Prüfe ob Labels numerische Codes haben (ordinal)
      numeric_codes <- suppressWarnings(as.numeric(names(labels)))
      if (!any(is.na(numeric_codes)) && length(unique(numeric_codes)) > 2) {
        is_ordinal_matrix <- TRUE
        cat("Ordinale Matrix erkannt basierend auf numerischen Codes in Kodierung\n")
      }
    }
  }
  
  # AUTOMATISCHE ERKENNUNG: Prüfe ob die tatsächlichen Werte numerisch sind
  is_numeric_matrix <- FALSE
  if (!has_coding && !is_dichotomous_matrix && !is_ordinal_matrix) {
    # Sammle Stichprobe von Werten aus allen Matrix-Items
    sample_values <- c()
    for (var in matrix_vars[1:min(3, length(matrix_vars))]) {  # Prüfe max. 3 Items
      var_values <- data[[var]][!is.na(data[[var]]) & data[[var]] != ""]
      sample_values <- c(sample_values, as.character(var_values[1:min(20, length(var_values))]))
    }
    
    # Prüfe ob die Werte numerisch konvertierbar sind
    numeric_test <- suppressWarnings(as.numeric(sample_values))
    proportion_numeric <- sum(!is.na(numeric_test)) / length(numeric_test)
    
    # Wenn > 80% der Werte numerisch sind, behandle als numerische Matrix
    if (proportion_numeric > 0.8 && length(unique(numeric_test[!is.na(numeric_test)])) > 2) {
      is_numeric_matrix <- TRUE
      cat("Numerische Matrix automatisch erkannt (", round(proportion_numeric * 100, 1), 
          "% numerische Werte)\n", sep = "")
    }
  }
  
  if (has_coding || is_dichotomous_matrix || is_ordinal_matrix || is_numeric_matrix) {
    cat("Matrix hat Kodierung oder ist dichotom - erstelle zusätzliche numerische Statistiken\n")
    
    # Erstelle numerische Statistik-Tabelle
    numeric_stats_rows <- list()
    
    for (var in matrix_vars) {
      item_label <- extract_item_label(data, var, matrix_name)
      
      # Daten für dieses Item
      if (!use_na) {
        item_data <- data[!is.na(data[[var]]), ]
        item_values <- item_data[[var]]
      } else {
        item_values <- data[[var]]
      }
      
      # NEUE LOGIK: Unterscheide zwischen dichotom und ordinal basierend auf Kodierung
      if (is.na(var_config$coding) || var_config$coding == "") {  
        # Keine Kodierung - verwende Rohwerte (funktioniert jetzt auch für automatisch erkannte numerische Matrizen)
        numeric_values <- suppressWarnings(as.numeric(as.character(item_values)))
        if (is_numeric_matrix) {
          cat("  Automatische numerische Konvertierung für", var, "\n")
        }
      } else {
        # Kodierung vorhanden - prüfe Typ
        labels <- parse_coding(var_config$coding)
        
        if (!is.null(labels) && length(labels) <= 2) {
          # BINÄRE MATRIX: Leere Werte zu 0, "1"/Y zu 1, andere zu 0
          cat("  Binäre Matrix-Konvertierung für", var, "\n")
          numeric_values <- rep(0, nrow(data))  # Default: 0 für ALLE Zeilen
          
          # Bearbeite ALLE Zeilen im Original-Datensatz
          for (i in seq_len(nrow(data))) {
            val <- data[[var]][i] 
            if (!is.na(val) && val != "") {
              if (val %in% c("Y", "1")) {
                numeric_values[i] <- 1
              }
              # Andere Werte bleiben 0
            }
          }
          
          cat("    Binäre Werte:", sum(numeric_values == 1), "von", length(numeric_values), "= 1\n")
          
        } else {
          # ORDINALE MATRIX: Verwende bestehende Kodierungs-Extraktion
          cat("  Ordinale Matrix-Konvertierung für", var, "\n")
          numeric_values <- extract_numeric_from_matrix_coding(
            item_values, 
            var_config$coding, 
            var_config$min_value,  
            var_config$max_value  
          )
        }
      }
      
      
      # Prüfe Gültigkeit der numerischen Werte
      valid_numeric_values <- numeric_values[!is.na(numeric_values)]
      
      if (length(valid_numeric_values) > 0 || is_dichotomous_matrix) {
        if (is_dichotomous_matrix) {
          # Für dichotome Matrix: Alle Werte verwenden (inkl. 0) 
          # KORREKTUR: N = Gesamtstichprobe, nicht nur die mit Werten
          stats_row <- data.frame(
            Item = item_label,
            N = nrow(data),  # KORREKTUR: Gesamtstichprobe statt length(numeric_values)
            Anteil_Ja = round(mean(numeric_values, na.rm = TRUE), DIGITS_ROUND),
            Anzahl_Ja = sum(numeric_values == 1, na.rm = TRUE),
            Anzahl_Nein = sum(numeric_values == 0, na.rm = TRUE),
            stringsAsFactors = FALSE
          )
        } else {
          # Für ordinale Matrix: Verwende nur gültige Werte für Statistiken (unverändert)
          stats_row <- data.frame(
            Item = item_label,
            N = length(valid_numeric_values),
            Mittelwert = round(mean(valid_numeric_values, na.rm = TRUE), DIGITS_ROUND),
            Median = round(median(valid_numeric_values, na.rm = TRUE), DIGITS_ROUND),
            Q1 = round(as.numeric(quantile(valid_numeric_values, 0.25, na.rm = TRUE)), DIGITS_ROUND),
            Q3 = round(as.numeric(quantile(valid_numeric_values, 0.75, na.rm = TRUE)), DIGITS_ROUND),
            Min = min(valid_numeric_values, na.rm = TRUE),
            Max = max(valid_numeric_values, na.rm = TRUE),
            SD = round(sd(valid_numeric_values, na.rm = TRUE), DIGITS_ROUND),
            stringsAsFactors = FALSE
          )
        }
        
        numeric_stats_rows[[var]] <- stats_row
      }
    }
    
    # Kombiniere numerische Statistiken
    if (length(numeric_stats_rows) > 0) {
      numeric_stats_table <- do.call(rbind, numeric_stats_rows)
      rownames(numeric_stats_table) <- NULL
      
      # Bestimme Rückgabe-Typ
      matrix_type <- if (is_dichotomous_matrix) "matrix_dichotomous" 
      else if (is_ordinal_matrix) "matrix_ordinal" 
      else if (is_numeric_matrix) "matrix_numeric"  # Automatisch erkannte numerische Matrix
      else "matrix_ordinal"  # Default für Matrices mit Kodierung
      
      # Füge numerische Tabelle zum Rückgabeobjekt hinzu
      return(list(
        table_categorical = result_table,
        table_numeric = numeric_stats_table,
        variable = matrix_name,
        question = question_text,
        type = matrix_type,
        matrix_items = matrix_vars,
        n_items = length(matrix_vars),
        response_categories = unique_responses,
        response_labels = response_labels,
        weighted = !is.null(survey_obj) && WEIGHTS,
        has_coding = TRUE,
        is_dichotomous = is_dichotomous_matrix
      ))
    }
    
    if (length(numeric_stats_rows) == 0 && is_dichotomous_matrix) {
      cat("WARNUNG: Dichotome Matrix erkannt, aber keine numerischen Statistiken erzeugt\n")
      
      return(list(
        table = result_table,
        variable = matrix_name,
        question = question_text,
        type = "matrix_dichotomous",
        matrix_items = matrix_vars,
        n_items = length(matrix_vars),
        response_categories = unique_responses,
        response_labels = response_labels,
        weighted = !is.null(survey_obj) && WEIGHTS,
        has_coding = has_coding,
        is_dichotomous = TRUE
      ))
    }
  }
  
  # Fallback falls numeric_stats_rows leer bleibt
  return(list(
    table = result_table,
    variable = matrix_name,
    question = question_text,
    type = if (is_dichotomous_matrix) "matrix_dichotomous" else "matrix",
    matrix_items = matrix_vars,
    n_items = length(matrix_vars),
    response_categories = unique_responses,
    response_labels = response_labels,
    weighted = !is.null(survey_obj) && WEIGHTS,
    has_coding = has_coding,
    is_dichotomous = is_dichotomous_matrix
  ))
}



# Hilfsfunktion: Intelligente Sortierung von Antwortkategorien
sort_response_categories <- function(responses) {
  # Definiere bekannte ordinale Reihenfolgen
  ordinal_patterns <- list(
    # Zustimmungsskala (negativ -> positiv)
    c("Stimme gar nicht zu", "Stimme eher nicht zu", "Teils/teils", 
      "Stimme eher zu", "Stimme voll zu"),
    
    # Wichtigkeit (negativ -> positiv)
    c("Unwichtig", "Weniger wichtig", "Wichtig", "Sehr wichtig"),
    
    # Schulnoten (schlecht -> gut)
    c("Ungenügend", "Mangelhaft", "Ausreichend", "Befriedigend", "Gut", "Sehr gut"),
    
    # Häufigkeit (selten -> häufig)
    c("Nie", "Selten", "Manchmal", "Oft", "Immer"),
    c("Nie", "Selten", "Gelegentlich", "Häufig", "Sehr häufig"),
    
    # Zustimmung kurzform (negativ -> positiv)
    c("Trifft gar nicht zu", "Trifft eher nicht zu", "Teils/teils", "Trifft eher zu", "Trifft voll zu"),
    
    # Häufigkeiten (selten -> häufig)
    c("Nie", "Seltener", "Einmal pro Monat", "Mehrmals pro Monat", 
      "Einmal pro Woche", "Mehrmals pro Woche", "Täglich"),
    
    # Ja/Nein (Nein = 0, Ja = 1)
    c("Nein", "Ja"),
    
    # Bildungsabschlüsse (niedrig -> hoch)
    c("Kein Schulabschluss", "Hauptschulabschluss", "Realschulabschluss", "Abitur"),
    c("Kein Berufsabschluss", "Lehre", "Fachschule", "Fachhochschule", "Universität")
  )
  
  # Prüfe, ob responses zu einem bekannten ordinalen Muster passt
  for (pattern in ordinal_patterns) {
    if (all(responses %in% pattern)) {
      # Sortiere nach dem bekannten Muster
      return(pattern[pattern %in% responses])
    }
  }
  
  # Falls kein Muster erkannt wird, alphabetisch sortieren
  return(sort(responses))
}

# Hilfsfunktion: Bereinige Spaltennamen für R
make_clean_colname <- function(text) {
  # Entferne/ersetze problematische Zeichen
  clean <- gsub("[^A-Za-z0-9_]", "_", text)
  clean <- gsub("_{2,}", "_", clean)  # Mehrfache Unterstriche reduzieren
  clean <- gsub("^_|_$", "", clean)   # Führende/nachfolgende Unterstriche entfernen
  
  # Falls leer oder nur Zahlen, Präfix hinzufügen
  if (clean == "" || grepl("^[0-9]+$", clean)) {
    clean <- paste0("Kategorie_", clean)
  }
  
  return(clean)
}


# =============================================================================
# NEUE VEREINFACHTE HILFSFUNKTIONEN
# =============================================================================

update_config_variable_names <- function(config, data) {
  "Passt alle Variablennamen in der Config an die sanitierten Daten an"
  
  cat("Aktualisiere Config für sanitierte Variablennamen...\n")
  
  # 1. VARIABLEN-SHEET
  config$variablen$variable_name <- update_variable_list(config$variablen$variable_name, names(data))
  
  # 2. KREUZTABELLEN-SHEET
  if (nrow(config$kreuztabellen) > 0) {
    config$kreuztabellen$variable_1 <- update_variable_list(config$kreuztabellen$variable_1, names(data))
    config$kreuztabellen$variable_2 <- update_variable_list(config$kreuztabellen$variable_2, names(data))
  }
  
  # 3. REGRESSIONS-SHEET
  if (nrow(config$regressionen) > 0) {
    config$regressionen$dependent_var <- update_variable_list(config$regressionen$dependent_var, names(data))
    
    # Unabhängige Variablen (durch ; getrennt)
    for (i in 1:nrow(config$regressionen)) {
      indep_vars <- str_split(config$regressionen$independent_vars[i], ";")[[1]]
      indep_vars <- str_trim(indep_vars)
      updated_indep_vars <- update_variable_list(indep_vars, names(data))
      config$regressionen$independent_vars[i] <- paste(updated_indep_vars, collapse = ";")
    }
  }
  
  # 4. TEXTANTWORTEN-SHEET (FEHLTE!)
  if (nrow(config$textantworten) > 0) {
    config$textantworten$text_variable <- update_variable_list(config$textantworten$text_variable, names(data))
    config$textantworten$sort_variable <- update_variable_list(config$textantworten$sort_variable, names(data))
  }
  
  return(config)
}

update_variable_list <- function(config_vars, data_vars) {
  "Aktualisiert eine Liste von Variablennamen basierend auf verfügbaren Daten - MIT INTERAKTIONSTERM-SUPPORT"
  
  updated_vars <- character(length(config_vars))
  
  for (i in seq_along(config_vars)) {
    original_var <- str_trim(config_vars[i])
    
    # NEUE LOGIK: Prüfe auf Interaktionsterm ZUERST
    if (grepl("\\*", original_var)) {
      cat("  Interaktionsterm erkannt:", original_var, "\n")
      
      # Extrahiere beide Variablen des Interaktionsterms
      interaction_vars <- str_split(original_var, "\\*")[[1]]
      interaction_vars <- str_trim(interaction_vars)
      
      # Prüfe ob beide Variablen existieren
      updated_interaction_vars <- c()
      all_vars_found <- TRUE
      
      for (int_var in interaction_vars) {
        # Versuche normale Variable zu finden
        found_var <- find_single_variable_simple(int_var, data_vars)
        if (!is.null(found_var) && found_var %in% data_vars) {
          updated_interaction_vars <- c(updated_interaction_vars, found_var)
          cat("    ", int_var, "→", found_var, "\n")
        } else {
          cat("    FEHLER:", int_var, "nicht gefunden\n")
          all_vars_found <- FALSE
          break
        }
      }
      
      if (all_vars_found) {
        # Rekonstruiere Interaktionsterm mit korrekten Variablennamen
        updated_vars[i] <- paste(updated_interaction_vars, collapse = "*")
        cat("  Interaktionsterm aktualisiert:", original_var, "→", updated_vars[i], "\n")
        next
      } else {
        cat("  WARNUNG: Interaktionsterm", original_var, "- nicht alle Variablen gefunden\n")
        updated_vars[i] <- original_var  # Behalte original
        next
      }
    }
    
    # NORMALE VARIABLE (bestehende Logik)
    # 1. Direkte Übereinstimmung nach Sanitization
    sanitized_var <- make.names(original_var)
    if (sanitized_var %in% data_vars) {
      updated_vars[i] <- sanitized_var
      if (sanitized_var != original_var) {
        cat("  ", original_var, "→", sanitized_var, "\n")
      }
      next
    }
    
    # 2. [other] Variablen behandeln
    if (grepl("\\[other\\]$", original_var)) {
      other_var <- find_other_variable_simple(original_var, data_vars)
      if (!is.null(other_var)) {
        updated_vars[i] <- other_var
        cat("  [other]", original_var, "→", other_var, "\n")
        next
      }
    }
    
    # 3. Matrix-Items direkt suchen
    matrix_match <- find_exact_matrix_item(original_var, data_vars)
    if (!is.null(matrix_match)) {
      updated_vars[i] <- matrix_match
      cat("  Matrix-Item:", original_var, "→", matrix_match, "\n")
      next
    }
    
    # 4. Matrix-Variablen (bleiben unverändert für spätere Behandlung)
    if (any(grepl(paste0("^", make.names(original_var), "[\\.\\[_-]"), data_vars))) {
      updated_vars[i] <- make.names(original_var)  # Sanitized aber nicht gefunden = Matrix
      cat("  Matrix:", original_var, "→ wird später verarbeitet\n")
      next
    }
    
    # 5. Variable nicht gefunden
    cat("  WARNUNG:", original_var, "nicht gefunden\n")
    updated_vars[i] <- make.names(original_var)  # Behalte sanitized version
  }
  
  return(updated_vars)
}

# Neue Hilfsfunktion für einfache Variablensuche
find_single_variable_simple <- function(target_var, data_vars) {
  "Findet eine einzelne Variable in den Daten"
  
  # 1. Direkte Übereinstimmung
  if (target_var %in% data_vars) {
    return(target_var)
  }
  
  # 2. Sanitized Version
  sanitized_var <- make.names(target_var)
  if (sanitized_var %in% data_vars) {
    return(sanitized_var)
  }
  
  # 3. [other] Behandlung
  if (grepl("\\[other\\]$", target_var)) {
    return(find_other_variable_simple(target_var, data_vars))
  }
  
  # 4. Matrix-Item
  matrix_match <- find_exact_matrix_item(target_var, data_vars)
  if (!is.null(matrix_match)) {
    return(matrix_match)
  }
  
  return(NULL)
}

# NEUE HILFSFUNKTION: Matrix-Items direkt finden
find_exact_matrix_item <- function(target_var, data_vars) {
  "Findet exakte Matrix-Items in den Daten"
  
  # Verschiedene Matrix-Item Formate testen
  candidates <- c(
    make.names(target_var),                    # AS03.other.
    gsub("\\[", ".", target_var),             # AS03[other] -> AS03.other
    gsub("\\]", ".", gsub("\\[", ".", target_var))  # AS03[other] -> AS03.other.
  )
  
  for (candidate in candidates) {
    if (candidate %in% data_vars) {
      return(candidate)
    }
  }
  
  return(NULL)
}

find_other_variable_simple <- function(target_var, data_vars) {
  "Einfache Suche nach [other] Variablen"
  
  # Extrahiere Basis-Variable (entferne [other])
  base_var <- gsub("\\[other\\]$", "", target_var)
  base_var_sanitized <- make.names(base_var)
  
  # Suche nach common [other] patterns
  patterns <- paste0("^", base_var_sanitized, c("\\.other\\.", "_other$", "\\.other$"))
  
  for (pattern in patterns) {
    matches <- data_vars[grepl(pattern, data_vars)]
    if (length(matches) > 0) {
      return(matches[1])
    }
  }
  
  return(NULL)
}

apply_variable_labels <- function(data, custom_var_labels = NULL, custom_val_labels = NULL) {
  "Wendet Variable Labels BEFORE Faktor-Konvertierung an"
  
  cat("Wende Variable Labels an (vor Datentyp-Konvertierung)...\n")
  
  # Eigene Variable Labels als Attribut setzen (nur wenn vorhanden)
  if (!is.null(custom_var_labels) && length(custom_var_labels) > 0) {
    attr(data, "var.labels") <- custom_var_labels
  }
  
  # Value Labels nur für Variablen die existieren UND NOCH NICHT FAKTOREN SIND
  if (requireNamespace("labelled", quietly = TRUE) && !is.null(custom_val_labels)) {
    for (var in names(custom_val_labels)) {
      if (var %in% names(data)) {
        # NUR wenn Variable noch kein Faktor ist
        if (!is.factor(data[[var]])) {
          labels <- custom_val_labels[[var]]
          data <- safe_apply_labels(data, var, labels)
          cat("  Labels angewendet für:", var, "\n")
        } else {
          cat("  Übersprungen (bereits Faktor):", var, "\n")
        }
      }
    }
  }
  
  return(data)
}

# Datentypen vorbereiten
prepare_variable_types <- function(data, config) {
  cat("Bereite Variablentypen vor...\n")
  
  for (i in 1:nrow(config$variablen)) {
    var_name <- config$variablen$variable_name[i]
    var_type <- config$variablen$data_type[i]
    
    if (var_name %in% names(data)) {
      # Labels beibehalten falls vorhanden (für RDS mit gelabelten Daten)
      original_labels <- attr(data[[var_name]], "labels")
      
      # Datentyp setzen
      if (var_type == "numeric") {
        data[[var_name]] <- as.numeric(data[[var_name]])
      } else if (var_type %in% c("nominal_coded", "ordinal", "dichotom")) {
        data[[var_name]] <- as.factor(data[[var_name]])
      } else if (var_type %in% c("nominal_nominal", "matrix")) {
        data[[var_name]] <- as.character(data[[var_name]])
      }
      
      # Labels wieder setzen
      if (!is.null(original_labels)) {
        attr(data[[var_name]], "labels") <- original_labels
      }
    }
  }
  
  return(data)
}

prepare_variable_types_minimal <- function(data, config) {
  cat("Setze Variablentypen (mit Label-Erhaltung)...\n")
  
  for (i in 1:nrow(config$variablen)) {
    var_name <- config$variablen$variable_name[i]
    var_type <- config$variablen$data_type[i]
    
    if (var_name %in% names(data)) {
      # *** NEU: Labels vor Konvertierung sichern ***
      original_labels <- attr(data[[var_name]], "labels")
      original_label <- attr(data[[var_name]], "label")
      original_format <- attr(data[[var_name]], "format")
      
      # NUR numerische Konvertierung, keine Factors
      if (var_type == "numeric") {
        # Sichere numerische Konvertierung
        tryCatch({
          original_values <- data[[var_name]]
          numeric_values <- suppressWarnings(as.numeric(original_values))
          
          # Prüfe Konvertierungserfolg
          successful_conversions <- sum(!is.na(numeric_values))
          total_values <- sum(!is.na(original_values))
          
          if (total_values > 0 && successful_conversions / total_values >= 0.8) {
            data[[var_name]] <- numeric_values
            
            # *** NEU: Labels wiederherstellen ***
            if (!is.null(original_labels)) {
              attr(data[[var_name]], "labels") <- original_labels
            }
            if (!is.null(original_label)) {
              attr(data[[var_name]], "label") <- original_label
            }
            if (!is.null(original_format)) {
              attr(data[[var_name]], "format") <- original_format
            }
            
            cat("  ", var_name, "→ numerisch (", successful_conversions, "/", total_values, "erfolgreich) + Labels erhalten\n")
          } else {
            cat("  WARNUNG:", var_name, "- numerische Konvertierung teilweise fehlgeschlagen\n")
            data[[var_name]] <- numeric_values  # Trotzdem zuweisen, aber warnen
            
            # Labels trotzdem wiederherstellen
            if (!is.null(original_labels)) {
              attr(data[[var_name]], "labels") <- original_labels
            }
            if (!is.null(original_label)) {
              attr(data[[var_name]], "label") <- original_label
            }
          }
        }, error = function(e) {
          cat("  FEHLER:", var_name, "- numerische Konvertierung fehlgeschlagen:", e$message, "\n")
        })
      }
      # Factors werden in Analyse-Funktionen erstellt (Labels bleiben erhalten)
    }
  }
  
  return(data)
}

# Hilfsfunktion um numerische Werte aus ordinalen Textvariablen zu extrahieren
extract_numeric_from_ordinal <- function(x) {
  if (is.numeric(x)) return(x)
  if (is.character(x) || is.factor(x)) {
    # Extrahiere Zahl am Anfang des Strings (z.B. "5 (sehr zufrieden)" -> 5)
    numeric_values <- as.numeric(str_extract(as.character(x), "^\\d+"))
    return(numeric_values)
  }
  return(as.numeric(x))
}

prepare_numeric_data_safe <- function(data_subset) {
  if(is.null(data_subset) || nrow(data_subset) == 0 || ncol(data_subset) == 0) {
    return(NULL)
  }
  
  tryCatch({
    numeric_data <- data_subset
    
    for(col in names(numeric_data)) {
      # Versuche Konvertierung zu numerisch
      if(is.factor(numeric_data[[col]]) || is.character(numeric_data[[col]])) {
        # Falls Faktor mit numerischen Levels
        if(is.factor(numeric_data[[col]])) {
          numeric_data[[col]] <- as.numeric(as.character(numeric_data[[col]]))
        } else {
          numeric_data[[col]] <- as.numeric(numeric_data[[col]])
        }
      }
    }
    
    return(numeric_data)
    
  }, error = function(e) {
    cat("Fehler bei numerischer Konvertierung:", e$message, "\n")
    return(NULL)
  })
}


convert_text_nas <- function(data, config) {
  cat("Konvertiere Text-NAs zu echten NAs (mit Label-Erhaltung)...\n")
  
  na_patterns <- c("N/A", "n/a", "NA", "NULL", "", " ", "missing", "Missing")
  
  for (i in 1:nrow(config$variablen)) {
    var_name <- config$variablen$variable_name[i]
    var_type <- config$variablen$data_type[i]
    
    if (var_name %in% names(data) && var_type %in% c("nominal_text", "nominal_coded", "nominal", "ordinal")) {
      # *** NEU: Labels vor Konvertierung sichern ***
      original_labels <- attr(data[[var_name]], "labels")
      original_label <- attr(data[[var_name]], "label")
      
      # Text-NAs finden und zu echten NAs konvertieren
      na_mask <- data[[var_name]] %in% na_patterns
      if (any(na_mask, na.rm = TRUE)) {
        data[[var_name]][na_mask] <- NA
        
        # *** NEU: Labels wiederherstellen ***
        if (!is.null(original_labels)) {
          attr(data[[var_name]], "labels") <- original_labels
        }
        if (!is.null(original_label)) {
          attr(data[[var_name]], "label") <- original_label
        }
        
        cat("  -", var_name, ":", sum(na_mask, na.rm = TRUE), "Text-NAs konvertiert + Labels erhalten\n")
      }
    }
  }
  
  return(data)
}

# Reverse Coding anwenden
apply_reverse_coding <- function(data, config) {
  reverse_vars <- config$variablen %>% 
    filter(reverse_coding == TRUE & !is.na(min_value) & !is.na(max_value))
  
  if (nrow(reverse_vars) > 0) {
    cat("Wende Reverse Coding an für:", nrow(reverse_vars), "Variablen\n")
    
    for (i in 1:nrow(reverse_vars)) {
      var <- reverse_vars$variable_name[i]
      min_val <- reverse_vars$min_value[i]
      max_val <- reverse_vars$max_value[i]
      
      if (var %in% names(data)) {
        # Reverse coding anwenden
        data[[var]] <- (min_val + max_val) - as.numeric(data[[var]])
        
        # Dokumentation hinzufügen
        attr(data[[var]], "reverse_coded") <- TRUE
        cat("  -", var, ": umkodiert (", min_val, "-", max_val, ")\n")
      }
    }
  }
  
  return(data)
}

# Automatische Kategorienerkennung für nominal_text
auto_detect_categories <- function(data, config) {
  nominal_text_vars <- config$variablen %>% 
    filter(data_type == "nominal_text") %>% 
    pull(variable_name)
  
  category_info <- list()
  
  if (length(nominal_text_vars) > 0) {
    cat("Erkenne Kategorien für nominal_text Variablen:", length(nominal_text_vars), "Variablen\n")
    
    for (var in nominal_text_vars) {
      if (var %in% names(data)) {
        # Eindeutige Kategorien finden (ohne NA)
        unique_cats <- data[[var]] %>% 
          na.omit() %>% 
          unique() %>% 
          sort()
        
        category_info[[var]] <- list(
          categories = unique_cats,
          n_categories = length(unique_cats)
        )
        
        # Kurze Versionen für Charts erstellen
        data[[paste0(var, "_short")]] <- str_trunc(data[[var]], 50, ellipsis = "...")
        
        cat("  -", var, ":", length(unique_cats), "Kategorien erkannt\n")
      }
    }
  }
  
  return(list(data = data, category_info = category_info))
}

safe_apply_labels <- function(data, var_name, labels) {
  "Sichere Anwendung von Value Labels auf eine Variable"
  
  if (!var_name %in% names(data)) {
    return(data)
  }
  
  tryCatch({
    # Aktuelle Variable
    var_data <- data[[var_name]]
    
    # Eindeutige Werte ermitteln (ohne NA)
    unique_vals <- unique(var_data[!is.na(var_data)])
    
    # Labels entsprechend dem Datentyp anpassen
    if (is.numeric(var_data)) {
      # Numerische Variable - Labels-Keys zu numeric konvertieren
      numeric_keys <- suppressWarnings(as.numeric(names(labels)))
      valid_keys <- !is.na(numeric_keys)
      
      if (any(valid_keys)) {
        final_labels <- labels[valid_keys]
        names(final_labels) <- numeric_keys[valid_keys]
        
        # Nur Labels für tatsächlich vorhandene Werte verwenden
        existing_keys <- names(final_labels)[names(final_labels) %in% unique_vals]
        if (length(existing_keys) > 0) {
          final_labels <- final_labels[existing_keys]
          data[[var_name]] <- labelled::set_value_labels(var_data, final_labels)
        }
      }
    } else {
      # Character/Factor Variable - Labels-Keys als character verwenden
      char_keys <- as.character(names(labels))
      names(labels) <- char_keys
      
      # Nur Labels für tatsächlich vorhandene Werte verwenden
      existing_keys <- char_keys[char_keys %in% as.character(unique_vals)]
      if (length(existing_keys) > 0) {
        final_labels <- labels[existing_keys]
        data[[var_name]] <- labelled::set_value_labels(var_data, final_labels)
      }
    }
    
    return(data)
    
  }, error = function(e) {
    warning(paste("Fehler beim Anwenden der Labels für Variable", var_name, ":", e$message))
    return(data)
  })
}

# Numerische Versionen für ordinale Variablen erstellen
create_numeric_versions <- function(data, config) {
  ordinal_vars <- config$variablen %>% 
    filter(data_type == "ordinal")
  
  if (nrow(ordinal_vars) > 0) {
    cat("Erstelle numerische Versionen für ordinale Variablen:", nrow(ordinal_vars), "Variablen\n")
    
    for (i in 1:nrow(ordinal_vars)) {
      var <- ordinal_vars$variable_name[i]
      
      if (var %in% names(data)) {
        # Numerische Version erstellen - ROBUST für haven_labelled + AO-Pattern
        var_data <- data[[var]]
        
        # Zu character konvertieren (haven_labelled -> character)
        char_values <- as.character(var_data)
        
        # Numerische Extraktion mit Fallback-Strategien
        numeric_values <- rep(NA_real_, length(char_values))
        
        for (j in seq_along(char_values)) {
          val <- char_values[j]
          if (!is.na(val) && val != "") {
            # Strategie 1: AO-Pattern (AO01 -> 1, AO02 -> 2)
            if (grepl("^AO\\d+$", val)) {
              numeric_values[j] <- as.numeric(gsub("^AO0*", "", val))
            }
            # Strategie 2: Zahl am Anfang ("5 (sehr gut)" -> 5)
            else if (grepl("^\\d+", val)) {
              numeric_values[j] <- as.numeric(str_extract(val, "^\\d+"))
            }
            # Strategie 3: Direkte Konvertierung
            else {
              numeric_values[j] <- suppressWarnings(as.numeric(val))
            }
          }
        }
        
        data[[paste0(var, "_num")]] <- numeric_values
        
        # Metadaten kopieren
        attr(data[[paste0(var, "_num")]], "original_variable") <- var
        attr(data[[paste0(var, "_num")]], "variable_type") <- "ordinal_numeric"
        
        successful <- sum(!is.na(numeric_values))
        cat("  ", var, "-> ", var, "_num (", successful, " gültige Werte)\n", sep = "")
      }
    }
  }
  
  return(data)
}

# Sichere Index-Erstellung (angepasste Version)
create_numeric_index_safe <- function(data_subset, index_name = "Index") {
  if(is.null(data_subset) || nrow(data_subset) == 0 || ncol(data_subset) == 0) {
    cat("Fehler:", index_name, "- Keine Daten verfügbar\n")
    return(NULL)
  }
  
  tryCatch({
    cat("Erstelle", index_name, "aus", ncol(data_subset), "Variablen mit", nrow(data_subset), "Fällen\n")
    
    # Konvertiere alle Spalten zu numerisch
    numeric_data <- prepare_numeric_data_safe(data_subset)
    
    if(is.null(numeric_data)) {
      cat("Fehler:", index_name, "- Datenkonvertierung fehlgeschlagen\n")
      return(NULL)
    }
    
    # Prüfe welche Spalten erfolgreich konvertiert wurden
    numeric_success <- sapply(numeric_data, function(x) {
      is.numeric(x) && sum(!is.na(x)) > 0
    })
    
    cat("Erfolgreich konvertierte Spalten:", sum(numeric_success), "von", length(numeric_success), "\n")
    
    if(!any(numeric_success)) {
      cat("Fehler:", index_name, "- Keine Spalten erfolgreich konvertiert\n")
      return(NULL)
    }
    
    # Verwende nur erfolgreich konvertierte Spalten
    numeric_data_clean <- numeric_data[, numeric_success, drop = FALSE]
    
    # Erstelle Index mit manueller rowMeans-Berechnung
    index_values <- rep(NA, nrow(numeric_data_clean))
    
    for(i in 1:nrow(numeric_data_clean)) {
      row_values <- as.numeric(numeric_data_clean[i, ])
      valid_values <- row_values[!is.na(row_values)]
      
      if(length(valid_values) > 0) {
        index_values[i] <- mean(valid_values)
      }
    }
    
    # Bereinige problematische Werte
    index_values[is.nan(index_values) | is.infinite(index_values)] <- NA
    
    valid_count <- sum(!is.na(index_values))
    cat(index_name, "erfolgreich erstellt mit", valid_count, "gültigen Werten von", length(index_values), "\n")
    
    if(valid_count < 10) {
      cat("Warnung:", index_name, "- Sehr wenige gültige Werte (", valid_count, ")\n")
    }
    
    return(index_values)
    
  }, error = function(e) {
    cat("Fehler bei Index-Erstellung", index_name, ":", e$message, "\n")
    return(NULL)
  })
}

# Spezielle Behandlung für binäre Grid-Fragen (1 oder leer)
convert_binary_grid_to_numeric <- function(x) {
  # Leere Werte und NA zu 0, "1" zu 1
  case_when(
    is.na(x) | x == "" | str_trim(x) == "" ~ 0,
    x == "1" | x == 1 ~ 1,
    TRUE ~ 0
  )
}

create_survey_indices <- function(data, config, index_definitions) {
  
  # Früher Ausstieg wenn keine Indices definiert
  if (length(index_definitions) == 0) {
    cat("Keine Survey-Indizes definiert - überspringe Index-Erstellung\n")
    return(list(data = data, config = config))
  }
  
  cat("Erstelle Survey-Indizes...\n")
  
  for (def in index_definitions) {
    name <- def$name
    label <- def$label
    vars_original <- def$vars_original
    vars_sanitized <- make.names(vars_original)
    
    cat("\n", label, ": Suche nach sanitisierten Variablen ", paste(vars_sanitized, collapse = ", "), "\n")
    
    vars_present <- vars_sanitized[vars_sanitized %in% names(data)]
    cat("Gefundene Variablen: ", paste(vars_present, collapse = ", "), "\n")
    
    if (length(vars_present) > 0) {
      # Daten extrahieren und ordinal zu numerisch konvertieren
      subdata <- data[vars_present]
      for (var in names(subdata)) {
        if(!is.null(def$binary) && def$binary == TRUE) {
          # Binäre Behandlung: leer/NA = 0, "1" = 1
          subdata[[var]] <- case_when(
            is.na(subdata[[var]]) | subdata[[var]] == "" | str_trim(as.character(subdata[[var]])) == "" ~ 0,
            subdata[[var]] == "1" | subdata[[var]] == 1 ~ 1,
            TRUE ~ 0
          )
        } else {
          subdata[[var]] <- extract_numeric_from_ordinal(subdata[[var]])
        }
      }
      cat("Konvertierung abgeschlossen\n")
      
      # Index berechnen
      index_vec <- create_numeric_index_safe(subdata, label)
      
      if (!is.null(index_vec)) {
        data[[name]] <- index_vec
        config <- add_index_to_config(config, name, label, vars_present)
        cat("Index ", label, " erfolgreich erstellt und hinzugefügt\n")
      } else {
        cat("Fehler beim Erstellen des Index: ", label, "\n")
      }
    } else {
      cat("Keine gültigen Variablen für ", label, " gefunden\n")
    }
  }
  
  return(list(data = data, config = config))
}


# Korrigierte add_index_to_config Funktion
add_index_to_config <- function(config, index_var_name, index_description, source_vars) {
  cat("Füge Index-Variable zur Konfiguration hinzu:", index_var_name, "\n")
  
  # Aktuelle Spalten der config$variablen ermitteln
  existing_cols <- names(config$variablen)
  cat("Vorhandene Spalten:", paste(existing_cols, collapse = ", "), "\n")
  
  # Neue Zeile für Index-Variable - direkt mit benötigten Werten erstellen
  new_index_row <- data.frame(
    variable_name = index_var_name,
    question_text = paste(index_description, "- erstellt aus:", paste(source_vars, collapse = ", ")),
    data_type = "numeric",
    coding = NA_character_,
    min_value = NA_real_,
    max_value = NA_real_,
    reverse_coding = FALSE,
    use_NA = FALSE,
    stringsAsFactors = FALSE
  )
  
  # Falls weitere Spalten existieren, mit NA füllen
  for (col in existing_cols) {
    if (!col %in% names(new_index_row)) {
      new_index_row[[col]] <- NA
    }
  }
  
  # Sicherstellen, dass die Spaltenreihenfolge stimmt
  new_index_row <- new_index_row[existing_cols]
  
  # Zur Konfiguration hinzufügen
  config$variablen <- rbind(config$variablen, new_index_row)
  
  cat("Index-Variable", index_var_name, "zur Konfiguration hinzugefügt\n")
  return(config)
}

# Hilfsfunktion für Likert-Skalen
convert_likert_to_numeric <- function(x) {
  case_when(
    str_detect(x, "^1.*überhaupt nicht") ~ 1,
    str_detect(x, "^2.*eher nicht") ~ 2,
    str_detect(x, "^3.*teils") ~ 3,
    str_detect(x, "^4.*eher") ~ 4,
    str_detect(x, "^5.*voll und ganz") ~ 5,
    str_detect(x, "Weiß nicht") ~ NA_real_,
    str_detect(x, "^1") ~ 1,  # Fallback für einfachere Kodierungen
    str_detect(x, "^2") ~ 2,
    str_detect(x, "^3") ~ 3,
    str_detect(x, "^4") ~ 4,
    str_detect(x, "^5") ~ 5,
    TRUE ~ NA_real_
  )
}

convert_to_factor_with_labels <- function(data, var_name, preserve_labels = TRUE) {
  "Konvertiert Variable zu Factor und behält Value Labels bei"
  
  var_data <- data[[var_name]]
  
  if (is.factor(var_data)) {
    return(data)  # Bereits Factor
  }
  
  # Value Labels speichern
  original_labels <- NULL
  if (preserve_labels && requireNamespace("labelled", quietly = TRUE)) {
    if (labelled::is.labelled(var_data)) {
      original_labels <- labelled::val_labels(var_data)
    }
  }
  
  # Zu Factor konvertieren
  data[[var_name]] <- as.factor(var_data)
  
  # Labels wieder anwenden falls vorhanden
  if (!is.null(original_labels) && length(original_labels) > 0) {
    # Factor levels mit Labels matchen
    factor_levels <- levels(data[[var_name]])
    
    # Labels den entsprechenden Levels zuordnen
    for (label_value in names(original_labels)) {
      if (label_value %in% factor_levels) {
        # Level umbenennen wenn Label verfügbar
        levels(data[[var_name]])[levels(data[[var_name]]) == label_value] <- 
          original_labels[[label_value]]
      }
    }
  }
  
  return(data)
}

create_survey_object <- function(data, weight_var) {
  "Erstellt Survey-Objekt mit character statt factor Variablen"
  
  if (!weight_var %in% names(data)) {
    return(NULL)
  }
  
  # Konvertiere zu data.frame (nicht tibble) für Survey-Kompatibilität
  survey_data <- as.data.frame(data)
  
  # ALLE FACTORS ZU CHARACTER konvertieren für Survey-Kompatibilität
  factor_vars <- sapply(survey_data, is.factor)
  if (any(factor_vars)) {
    factor_var_names <- names(survey_data)[factor_vars]
    cat("Konvertiere", length(factor_var_names), "Factor-Variablen zu Character für Survey-Objekt\n")
    
    for (var_name in factor_var_names) {
      survey_data[[var_name]] <- as.character(survey_data[[var_name]])
    }
  }
  
  # Survey-Objekt erstellen
  survey_obj <- svydesign(ids = ~1, weights = survey_data[[weight_var]], data = survey_data)
  
  return(survey_obj)
}

# =============================================================================
# VEREINFACHTE DATEN LADEN UND VORBEREITEN
# =============================================================================

load_and_prepare_data <- function(config, index_definitions = list(), custom_var_labels = NULL, custom_val_labels = NULL) {
  cat("\nLade Daten aus:", DATA_FILE, "\n")
  
  check_file_exists(DATA_FILE)
  
  file_ext <- get_file_extension(DATA_FILE)
  
  data <- switch(file_ext,
                 "xlsx" = read_excel(DATA_FILE),
                 "csv" = read.csv(DATA_FILE, stringsAsFactors = FALSE),
                 "rds" = readRDS(DATA_FILE),
                 stop("Nicht unterstütztes Dateiformat. Unterstützt: .xlsx, .csv, .rds")
  )
  
  cat("Daten geladen. Dimensionen:", nrow(data), "Zeilen,", ncol(data), "Spalten\n")
  
  # 1. ALLE VARIABLENNAMEN SANITIZEN (einmalig)
  cat("Sanitize alle Variablennamen...\n")
  names(data) <- make.names(names(data))
  
  # 2. BASIS-DATENAUFBEREITUNG (ohne Custom Variables)
  data <- convert_text_nas(data, config)
  data <- apply_reverse_coding(data, config)
  data <- create_numeric_versions(data, config)
  
  # 3. SURVEY-INDIZES ERSTELLEN (vor Custom Variables!)
  cat("Erstelle Survey-Indizes...\n")
  index_result <- create_survey_indices(data, config, index_definitions)
  data <- index_result$data
  config <- index_result$config
  
  # 4. CUSTOM VARIABLES ERSTELLEN (jetzt können sie auf Indices zugreifen)
  cat("Erstelle Custom-Variablen...\n")
  data <- add_custom_vars(data)
  
  # 5. CONFIG AN ALLE VARIABLEN ANPASSEN (jetzt mit Indices + Custom Variables!)
  cat("Aktualisiere Config für sanitierte Variablennamen...\n")
  config <- update_config_variable_names(config, data)
  
  # 6. GEWICHTUNGSVARIABLE PRÜFEN
  if (WEIGHTS) {
    WEIGHT_VAR_SANITIZED <- make.names(WEIGHT_VAR)
    if (!WEIGHT_VAR_SANITIZED %in% names(data)) {
      warning(paste("Gewichtungsvariable", WEIGHT_VAR, "nicht gefunden. Analysen werden ungewichtet durchgeführt."))
      WEIGHTS <<- FALSE
    } else {
      WEIGHT_VAR <<- WEIGHT_VAR_SANITIZED
    }
  }
  
  # 7. LABELS ANWENDEN (für alle Variablen)
  data <- apply_variable_labels(data, custom_var_labels, custom_val_labels)
  
  # 8. WEITERE AUFBEREITUNG
  category_info <- auto_detect_categories(data, config)
  data <- category_info$data
  
  # 9. METADATEN ENTFERNEN
  data <- data %>% select(-any_of(meta_vars_to_remove))
  
  # 10. VARIABLENTYPEN SETZEN
  data <- prepare_variable_types_minimal(data, config)
  
  cat("Datenaufbereitung abgeschlossen.\n")
  
  return(list(
    data = data,
    category_info = category_info$category_info,
    config = config
  ))
}


# =============================================================================
# DESKRIPTIVE STATISTIKEN
# =============================================================================

create_descriptive_tables <- function(prepared_data) {
  cat("\nErstelle deskriptive Statistiken...\n")
  
  data <- prepared_data$data
  config <- prepared_data$config
  category_info <- prepared_data$category_info
  
  results <- list()
  
  # Für jede Variable entsprechende Tabelle erstellen
  for (i in 1:nrow(config$variablen)) {
    var_name <- config$variablen$variable_name[i]
    var_type <- config$variablen$data_type[i]
    question_text <- config$variablen$question_text[i]
    use_na <- config$variablen$use_NA[i]
    
    # NEUER FIX: Sicherheitsprüfung für use_na
    if (is.na(use_na)) use_na <- FALSE
    
    cat("Verarbeite:", var_name, "(", var_type, ")\n")
    
    # Filter für diese Variable anwenden (falls vorhanden)
    filtered <- apply_row_filter(data, config$variablen[i,], names(data))
    current_data <- filtered$data
    filter_applied <- filtered$filtered
    filter_info <- filtered$filter_info
    
    # Gewichtetes Survey-Objekt für gefilterte Daten erstellen (falls nötig)
    current_survey_obj <- NULL
    if (WEIGHTS && WEIGHT_VAR %in% names(current_data)) {
      current_survey_obj <- create_survey_object(current_data, WEIGHT_VAR)
      if (filter_applied) {
        cat("  Filter angewendet: '", filter_info$filter_string, "' - ", 
            filter_info$filtered_n, " von ", filter_info$original_n, " Fällen (",
            round(filter_info$filtered_n/filter_info$original_n*100, 1), "%)\n", sep = "")
      }
    }
    
    # NEUER FIX: Fehlerbehandlung für jede Variable
    result <- tryCatch({
      if (var_name %in% names(current_data)) {
        switch(var_type,
               "numeric" = create_numeric_table(current_data, var_name, question_text, use_na, current_survey_obj),
               "nominal_coded" = create_nominal_coded_table(current_data, config$variablen[i,], use_na, current_survey_obj),
               "nominal_text" = create_nominal_text_table(current_data, var_name, question_text, use_na, category_info, current_survey_obj),
               "nominal" = create_nominal_text_table(current_data, var_name, question_text, use_na, category_info, current_survey_obj),
               "ordinal" = create_ordinal_table(current_data, config$variablen[i,], use_na, current_survey_obj),
               "dichotom" = create_dichotom_table(current_data, config$variablen[i,], use_na, current_survey_obj),
               "matrix" = create_matrix_table(current_data, config$variablen[i,], use_na, current_survey_obj)
        )
      } else if (config$variablen$data_type[i] == "matrix") {
        # Matrix-Variable behandeln
        create_matrix_table(current_data, config$variablen[i,], use_na, current_survey_obj)
      } else {
        cat("WARNUNG: Variable", var_name, "nicht in Daten gefunden\n")
        NULL
      }
    }, error = function(e) {
      cat("FEHLER bei Variable", var_name, ":", e$message, "\n")
      NULL
    })
    
    # Filter-Info zum Ergebnis hinzufügen (falls Filter angewendet)
    if (!is.null(result) && filter_applied) {
      result$filter_applied <- TRUE
      result$filter_string <- filter_info$filter_string
      result$original_n <- filter_info$original_n
      result$filtered_n <- filter_info$filtered_n
    } else if (!is.null(result)) {
      result$filter_applied <- FALSE
    }
    
    # Nur hinzufügen wenn Ergebnis nicht NULL
    if (!is.null(result)) {
      results[[var_name]] <- result
    }
  }
  
  cat("Deskriptive Statistiken für", length(results), "Variablen erstellt.\n")
  return(results)
}

# Deskriptive Tabelle für numerische Variablen
create_numeric_table <- function(data, var_name, question_text, use_na, survey_obj = NULL) {
  # Daten filtern basierend auf use_na
  if (!use_na) {
    data_filtered <- data[!is.na(data[[var_name]]), ]
  } else {
    data_filtered <- data
  }
  
  # Statistiken berechnen
  if (!is.null(survey_obj) && WEIGHTS) {
    # Gewichtete Statistiken mit Fehlerbehandlung
    if (!use_na) {
      survey_obj_filtered <- subset(survey_obj, !is.na(get(var_name)))
    } else {
      survey_obj_filtered <- survey_obj
    }
    
    stats <- tryCatch({
      list(
        n = nrow(survey_obj_filtered$variables),
        mean = as.numeric(svymean(as.formula(paste("~", var_name)), survey_obj_filtered, na.rm = !use_na)),
        median = as.numeric(svyquantile(as.formula(paste("~", var_name)), survey_obj_filtered, 0.5, na.rm = !use_na)[[1]][1]),
        q1 = as.numeric(svyquantile(as.formula(paste("~", var_name)), survey_obj_filtered, 0.25, na.rm = !use_na)[[1]][1]),
        q3 = as.numeric(svyquantile(as.formula(paste("~", var_name)), survey_obj_filtered, 0.75, na.rm = !use_na)[[1]][1]),
        min = as.numeric(svyquantile(as.formula(paste("~", var_name)), survey_obj_filtered, 0, na.rm = !use_na)[[1]][1]),
        max = as.numeric(svyquantile(as.formula(paste("~", var_name)), survey_obj_filtered, 1, na.rm = !use_na)[[1]][1]),
        sd = as.numeric(sqrt(svyvar(as.formula(paste("~", var_name)), survey_obj_filtered, na.rm = !use_na)))
      )
    }, error = function(e) {
      cat("FALLBACK: Gewichtete Statistiken für", var_name, "fehlgeschlagen:", e$message, "\n")
      cat("Verwende ungewichtete Statistiken als Fallback\n")
      
      # Fallback: Ungewichtete Statistiken
      values <- data_filtered[[var_name]]
      if (!use_na) values <- values[!is.na(values)]
      
      list(
        n = length(values[!is.na(values)]),
        mean = mean(values, na.rm = TRUE),
        median = median(values, na.rm = TRUE),
        q1 = as.numeric(quantile(values, 0.25, na.rm = TRUE)),
        q3 = as.numeric(quantile(values, 0.75, na.rm = TRUE)),
        min = min(values, na.rm = TRUE),
        max = max(values, na.rm = TRUE),
        sd = sd(values, na.rm = TRUE)
      )
    })
  } else {
    # Ungewichtete Statistiken
    values <- data_filtered[[var_name]]
    if (!use_na) values <- values[!is.na(values)]
    
    stats <- list(
      n = length(values[!is.na(values)]),
      mean = mean(values, na.rm = TRUE),
      median = median(values, na.rm = TRUE),
      q1 = as.numeric(quantile(values, 0.25, na.rm = TRUE)),
      q3 = as.numeric(quantile(values, 0.75, na.rm = TRUE)),
      min = min(values, na.rm = TRUE),
      max = max(values, na.rm = TRUE),
      sd = sd(values, na.rm = TRUE)
    )
  }
  
  # Debug: Längen der Statistiken prüfen
  # cat("Debug - Statistik-Längen:\n")
  # cat("n:", length(stats$n), "- Wert:", stats$n, "\n")
  # cat("mean:", length(stats$mean), "- Wert:", stats$mean, "\n")
  # cat("median:", length(stats$median), "- Wert:", stats$median, "\n")
  # cat("q1:", length(stats$q1), "- Wert:", stats$q1, "\n")
  # cat("q3:", length(stats$q3), "- Wert:", stats$q3, "\n")
  # cat("min:", length(stats$min), "- Wert:", stats$min, "\n")
  # cat("max:", length(stats$max), "- Wert:", stats$max, "\n")
  # cat("sd:", length(stats$sd), "- Wert:", stats$sd, "\n")
  
  # Ergebnis-Tabelle erstellen - ERWEITERT um Min/Max
  result_table <- data.frame(
    Kennwert = c("N", "Mittelwert", "Median", "Q1", "Q3", "Minimum", "Maximum", "Standardabweichung"),
    Wert = c(
      stats$n,
      round(stats$mean, DIGITS_ROUND),
      round(stats$median, DIGITS_ROUND),
      round(stats$q1, DIGITS_ROUND),
      round(stats$q3, DIGITS_ROUND),
      round(stats$min, DIGITS_ROUND),
      round(stats$max, DIGITS_ROUND),
      round(stats$sd, DIGITS_ROUND)
    ),
    stringsAsFactors = FALSE
  )
  
  # Fehlende Werte hinzufügen falls use_NA = TRUE
  if (use_na) {
    n_missing <- sum(is.na(data[[var_name]]))
    result_table <- rbind(result_table, 
                          data.frame(Kennwert = "Fehlende Werte", Wert = n_missing))
  }
  
  return(list(
    table = result_table,
    variable = var_name,
    question = question_text,
    type = "numeric",
    weighted = !is.null(survey_obj) && WEIGHTS
  ))
}

# Deskriptive Tabelle für nominal_coded Variablen
create_nominal_coded_table <- function(data, var_config, use_na, survey_obj = NULL) {
  var_name <- var_config$variable_name
  question_text <- var_config$question_text
  coding <- var_config$coding
  
  # Daten filtern
  if (!use_na) {
    data_filtered <- data[!is.na(data[[var_name]]), ]
  } else {
    data_filtered <- data
  }
  
  # Labels mit Priorisierung laden: RDS -> Config -> Code
  labels <- get_value_labels_with_priority(data, var_name, 
                                           list(variablen = data.frame(
                                             variable_name = var_name,
                                             coding = coding,
                                             stringsAsFactors = FALSE
                                           )))
  
  
  # Häufigkeiten berechnen
  if (!is.null(survey_obj) && WEIGHTS) {
    # Gewichtete Häufigkeiten
    if (!use_na) {
      survey_obj_filtered <- subset(survey_obj, !is.na(get(var_name)))
    } else {
      survey_obj_filtered <- survey_obj
    }
    
    freq_table <- svytable(as.formula(paste("~", var_name)), survey_obj_filtered)
    freq_df <- data.frame(
      Code = names(freq_table),
      Haeufigkeit_absolut = as.numeric(freq_table),
      stringsAsFactors = FALSE
    )
  } else {
    # Ungewichtete Häufigkeiten
    freq_table <- table(data_filtered[[var_name]], useNA = if(use_na) "always" else "no")
    freq_df <- data.frame(
      Code = names(freq_table),
      Haeufigkeit_absolut = as.numeric(freq_table),
      stringsAsFactors = FALSE
    )
  }
  
  # Relative Häufigkeiten
  freq_df$Haeufigkeit_relativ <- round(freq_df$Haeufigkeit_absolut / sum(freq_df$Haeufigkeit_absolut) * 100, DIGITS_ROUND)
  
  # Labels hinzufügen - VERBESSERT
  if (!is.null(labels) && length(labels) > 0) {
    freq_df$Label <- NA_character_
    
    # Debug: Zeige welche Labels wir haben
    cat("  Mapping Labels für", nrow(freq_df), "Codes\n")
    cat("  Verfügbare Label-Keys:", paste(names(labels), collapse=", "), "\n")
    
    for (i in seq_len(nrow(freq_df))) {
      code <- as.character(freq_df$Code[i])
      freq_df$Label[i] <- code  # Default: Verwende Code als Label
      
      # Direkte Übereinstimmung: "1" -> "1"
      if (code %in% names(labels)) {
        freq_df$Label[i] <- labels[code]
        next
      }
      
      # Pattern: AO01, AO02, AO03 -> extrahiere Nummer und versuche Match
      if (grepl("^[A-Z]+0*[0-9]+$", code)) {
        # Extrahiere Nummer: AO01 -> 1, AO02 -> 2, A001 -> 1
        num_part <- gsub("^[A-Z]+0*", "", code)
        
        # Versuche verschiedene Formate
        candidates <- c(
          num_part,                           # "1"
          paste0("AO", num_part),            # "AO1"
          paste0("AO0", num_part),           # "AO01"
          paste0("A", num_part),             # "A1"
          sprintf("%02d", as.numeric(num_part))  # "01"
        )
        
        for (candidate in candidates) {
          if (candidate %in% names(labels)) {
            freq_df$Label[i] <- labels[candidate]
            cat("    Mapped:", code, "->", candidate, "->", labels[candidate], "\n")
            break
          }
        }
      }
    }
    
    freq_df <- freq_df[, c("Code", "Label", "Haeufigkeit_absolut", "Haeufigkeit_relativ")]
  }
  
  return(list(
    table = freq_df,
    variable = var_name,
    question = question_text,
    type = "nominal_coded",
    weighted = !is.null(survey_obj) && WEIGHTS
  ))
}

# Deskriptive Tabelle für nominal_text Variablen
create_nominal_text_table <- function(data, var_name, question_text, use_na, category_info, survey_obj = NULL, config = NULL) {
  # Daten filtern
  if (!use_na) {
    data_filtered <- data[!is.na(data[[var_name]]), ]
  } else {
    data_filtered <- data
  }
  
  # Häufigkeiten berechnen
  if (!is.null(survey_obj) && WEIGHTS) {
    # Gewichtete Häufigkeiten
    if (!use_na) {
      survey_obj_filtered <- subset(survey_obj, !is.na(get(var_name)))
    } else {
      survey_obj_filtered <- survey_obj
    }
    
    freq_table <- svytable(as.formula(paste("~", var_name)), survey_obj_filtered)
    freq_df <- data.frame(
      Kategorie = names(freq_table),
      Haeufigkeit_absolut = as.numeric(freq_table),
      stringsAsFactors = FALSE
    )
  } else {
    # Ungewichtete Häufigkeiten
    freq_table <- table(data_filtered[[var_name]], useNA = if(use_na) "always" else "no")
    freq_df <- data.frame(
      Kategorie = names(freq_table),
      Haeufigkeit_absolut = as.numeric(freq_table),
      stringsAsFactors = FALSE
    )
  }
  
  # *** NEUE LOGIC: Labels anwenden (RDS -> Config -> Code) ***
  labels <- get_value_labels_with_priority(data, var_name, config)
  labels_found <- !is.null(labels) && length(labels) > 0
  
  
  # Labels anwenden falls gefunden
  if (labels_found && !is.null(labels)) {
    # Erstelle Label-Spalte
    freq_df$Kategorie_Label <- freq_df$Kategorie  # Default: verwende Code
    
    # Ersetze Codes durch Labels wo verfügbar
    for (code in names(labels)) {
      freq_df$Kategorie_Label[freq_df$Kategorie == code] <- labels[code]
    }
    
    # Verwende Labels für Anzeige, aber behalte Codes für Referenz
    freq_df <- freq_df[, c("Kategorie", "Kategorie_Label", "Haeufigkeit_absolut")]
    names(freq_df)[2] <- "Kategorie_mit_Label"
  }
  
  # Relative Häufigkeiten
  freq_df$Haeufigkeit_relativ <- round(freq_df$Haeufigkeit_absolut / sum(freq_df$Haeufigkeit_absolut) * 100, DIGITS_ROUND)
  
  # Kurze Versionen für bessere Darstellung (nur wenn keine Labels vorhanden)
  if (!"Kategorie_mit_Label" %in% names(freq_df)) {
    freq_df$Kategorie_kurz <- str_trunc(freq_df$Kategorie, 50, ellipsis = "...")
  }
  
  # Nach Häufigkeit sortieren
  freq_df <- freq_df[order(-freq_df$Haeufigkeit_absolut), ]
  
  return(list(
    table = freq_df,
    variable = var_name,
    question = question_text,
    type = "nominal_text",
    weighted = !is.null(survey_obj) && WEIGHTS
  ))
}

# Deskriptive Tabelle für ordinale Variablen
create_ordinal_table <- function(data, var_config, use_na, survey_obj = NULL) {
  var_name <- var_config$variable_name
  question_text <- var_config$question_text
  coding <- var_config$coding
  
  # NEUER FIX: Prüfe ob Variable existiert und gültige Daten hat
  if (!var_name %in% names(data)) {
    cat("WARNUNG: Variable", var_name, "nicht in Daten gefunden\n")
    return(NULL)
  }
  
  # Prüfe ob genügend gültige Daten vorhanden sind
  valid_data <- data[[var_name]][!is.na(data[[var_name]])]
  if (length(valid_data) == 0) {
    cat("WARNUNG: Keine gültigen Daten für Variable", var_name, "\n")
    return(NULL)
  }
  
  tryCatch({
    # Häufigkeitstabelle erstellen (wie nominal_coded)
    freq_result <- create_nominal_coded_table(data, var_config, use_na, survey_obj)
    
    # Zusätzlich numerische Statistiken für ordinale Variable
    var_name_num <- paste0(var_name, "_num")
    if (var_name_num %in% names(data)) {
      numeric_result <- create_numeric_table(data, var_name_num, question_text, use_na, survey_obj)
      
      # Kombiniere beide Ergebnisse
      return(list(
        table_frequencies = freq_result$table,
        table_numeric = numeric_result$table,
        variable = var_name,
        question = question_text,
        type = "ordinal",
        weighted = !is.null(survey_obj) && WEIGHTS
      ))
    }
    
    # Falls numerische Version nicht existiert, nur Häufigkeiten zurückgeben
    freq_result$type <- "ordinal"  # Typ korrigieren
    return(freq_result)
    
  }, error = function(e) {
    cat("FEHLER bei ordinaler Variable", var_name, ":", e$message, "\n")
    return(NULL)
  })
}


# Deskriptive Tabelle für dichotome Variablen
create_dichotom_table <- function(data, var_config, use_na, survey_obj = NULL) {
  # Behandle wie nominal_coded aber mit spezieller Kennzeichnung
  result <- create_nominal_coded_table(data, var_config, use_na, survey_obj)
  result$type <- "dichotom"
  return(result)
}





# =============================================================================
# ZENTRALE LABEL-EXTRAKTION MIT PRIORISIERUNG
# =============================================================================

get_value_labels_with_priority <- function(data, var_name, config = NULL) {
  "Extrahiert Value Labels mit Priorisierung: 1) RDS-Labels 2) Config-Labels 3) Code als Label"
  
  labels <- NULL
  
  # PRIORITÄT 1: Labels aus RDS-Daten
  if (var_name %in% names(data)) {
    # 1a. Attribut "labels" (häufigste Form)
    if (!is.null(attr(data[[var_name]], "labels"))) {
      labels_raw <- attr(data[[var_name]], "labels")
      
      # WICHTIG: Labels können umgekehrt sein! 
      # Check: Sind die Namen (names) die Texte und die Werte die Codes?
      # Beispiel: c("Universität" = "A1") statt c("A1" = "Universität")
      
      if (length(labels_raw) > 0) {
        # Prüfe ob Werte wie Codes aussehen und Namen wie Texte
        values <- as.character(labels_raw)
        names_vals <- names(labels_raw)
        
        # Wenn Werte kurz sind (codes) und Namen lang (labels), dann umkehren
        avg_value_len <- mean(nchar(values))
        avg_name_len <- mean(nchar(names_vals))
        
        if (avg_value_len < avg_name_len && avg_value_len <= 10) {
          # Umkehren: names werden zu values, values werden zu names
          labels <- setNames(names_vals, values)
          cat("  -> Labels aus RDS-Attribut 'labels' gefunden (umgekehrt) für", var_name, "
")
        } else {
          # Normal: verwende wie sie sind
          labels <- labels_raw
          cat("  -> Labels aus RDS-Attribut 'labels' gefunden für", var_name, "
")
        }
      }
    }
    
    # 1b. Haven/Labelled Package
    if ((is.null(labels) || length(labels) == 0) && requireNamespace("labelled", quietly = TRUE)) {
      if (labelled::is.labelled(data[[var_name]])) {
        labels_raw <- labelled::val_labels(data[[var_name]])
        
        if (!is.null(labels_raw) && length(labels_raw) > 0) {
          # Gleiche Logik: Prüfe ob umgekehrt
          values <- as.character(labels_raw)
          names_vals <- names(labels_raw)
          avg_value_len <- mean(nchar(values))
          avg_name_len <- mean(nchar(names_vals))
          
          if (avg_value_len < avg_name_len && avg_value_len <= 10) {
            labels <- setNames(names_vals, values)
            cat("  -> Labels aus RDS (labelled package, umgekehrt) für", var_name, "
")
          } else {
            labels <- labels_raw
            cat("  -> Labels aus RDS (labelled package) für", var_name, "
")
          }
        }
      }
    }
    
    # 1c. Direkte value.labels Attribut-Prüfung
    if ((is.null(labels) || length(labels) == 0) && !is.null(attr(data[[var_name]], "value.labels"))) {
      labels_raw <- attr(data[[var_name]], "value.labels")
      
      if (length(labels_raw) > 0) {
        values <- as.character(labels_raw)
        names_vals <- names(labels_raw)
        avg_value_len <- mean(nchar(values))
        avg_name_len <- mean(nchar(names_vals))
        
        if (avg_value_len < avg_name_len && avg_value_len <= 10) {
          labels <- setNames(names_vals, values)
          cat("  -> Labels aus RDS-Attribut 'value.labels' (umgekehrt) für", var_name, "
")
        } else {
          labels <- labels_raw
          cat("  -> Labels aus RDS-Attribut 'value.labels' für", var_name, "
")
        }
      }
    }
  }
  
  # PRIORITÄT 2: Labels aus Config-Kodierung (nur wenn keine RDS-Labels gefunden)
  if (is.null(labels) || length(labels) == 0) {
    if (!is.null(config)) {
      var_config <- config$variablen[config$variablen$variable_name == var_name, ]
      
      if (nrow(var_config) > 0 && !is.na(var_config$coding[1]) && var_config$coding[1] != "") {
        labels <- parse_coding(var_config$coding[1])
        if (!is.null(labels) && length(labels) > 0) {
          cat("  -> Labels aus Config-Kodierung gefunden für", var_name, "
")
        }
      }
    }
  }
  
  # Debug: Falls Labels gefunden, zeige sie an
  if (!is.null(labels) && length(labels) > 0) {
    cat("    Gefundene Labels:", paste(names(labels), "=", labels, collapse="; "), "
")
  }
  
  return(labels)
}

# =============================================================================
# NEUE HILFSFUNKTION: LABELS FÜR MATRIX-VARIABLEN MIT FALLBACK-STRATEGIE
# =============================================================================

get_matrix_labels <- function(data, matrix_vars, matrix_name = NULL, var_config = NULL, matrix_coding = NA) {
  "Holt Labels für Matrix-Variable mit mehreren Fallback-Strategien (eliminiert Redundanz)"
  
  labels <- NULL
  
  # Strategie 1: Von erstem Matrix-Item (RDS labels)
  if (length(matrix_vars) > 0 && !is.null(matrix_vars[1])) {
    temp_config <- if (!is.null(var_config)) {
      list(variablen = var_config)
    } else if (!is.na(matrix_coding)) {
      list(variablen = data.frame(
        variable_name = matrix_vars[1],
        coding = matrix_coding,
        stringsAsFactors = FALSE
      ))
    } else {
      NULL
    }
    
    labels <- get_value_labels_with_priority(data, matrix_vars[1], temp_config)
  }
  
  # Strategie 2: Von Matrix-Variable direkt (falls sie in Daten existiert)
  if ((is.null(labels) || length(labels) == 0) && !is.null(matrix_name) && matrix_name %in% names(data)) {
    temp_config <- if (!is.null(var_config)) {
      list(variablen = var_config)
    } else {
      NULL
    }
    labels <- get_value_labels_with_priority(data, matrix_name, temp_config)
  }
  
  # Strategie 3: Aus Config-Kodierung parsen
  if ((is.null(labels) || length(labels) == 0) && !is.na(matrix_coding) && matrix_coding != "") {
    labels <- parse_coding(matrix_coding)
  }
  
  # Strategie 4: Aus var_config direkt
  if ((is.null(labels) || length(labels) == 0) && !is.null(var_config) && 
      "coding" %in% names(var_config) && !is.na(var_config$coding) && var_config$coding != "") {
    labels <- parse_coding(var_config$coding)
  }
  
  return(labels)
}

# =============================================================================
# NEUE HILFSFUNKTION: INTELLIGENTES LABEL-MAPPING MIT PATTERN-ERKENNUNG
# =============================================================================

map_response_labels <- function(unique_responses, labels, verbose = TRUE) {
  "Mappt Response-Werte auf Labels mit intelligenter Pattern-Erkennung (AO01, A1, etc.)"
  
  if (is.null(labels) || length(labels) == 0) {
    # Keine Labels vorhanden - verwende rohe Werte
    response_labels <- unique_responses
    names(response_labels) <- unique_responses
    return(response_labels)
  }
  
  # Initialisiere mit rohen Werten als Fallback
  response_labels <- unique_responses
  names(response_labels) <- unique_responses
  
  if (verbose) {
    cat("  Mappe", length(unique_responses), "Responses auf", length(labels), "Labels\n")
    cat("    Label-Keys:", paste(names(labels), collapse=", "), "\n")
  }
  
  mapped_count <- 0
  
  for (response in unique_responses) {
    response_char <- as.character(response)
    mapped <- FALSE
    
    # Pattern 1: Direkte Übereinstimmung
    if (response_char %in% names(labels)) {
      response_labels[response_char] <- labels[response_char]
      mapped_count <- mapped_count + 1
      if (verbose) cat("      ✓ Direkt:", response_char, "->", labels[response_char], "\n")
      mapped <- TRUE
      next
    }
    
    # Pattern 2: AO-Pattern (AO01, AO02, AO1, AO2)
    if (!mapped && grepl("^AO\\d+$", response_char)) {
      numeric_code <- gsub("^AO0*", "", response_char)  # AO01 -> 1
      
      candidates <- c(
        response_char,                           # "AO01"
        paste0("AO", numeric_code),              # "AO1"
        numeric_code,                            # "1"
        sprintf("%02d", as.numeric(numeric_code)) # "01"
      )
      
      for (candidate in candidates) {
        if (candidate %in% names(labels)) {
          response_labels[response_char] <- labels[candidate]
          mapped_count <- mapped_count + 1
          if (verbose) cat("      ✓ AO-Pattern:", response_char, "-> (via", candidate, ") ->", labels[candidate], "\n")
          mapped <- TRUE
          break
        }
      }
      
      if (mapped) next
    }
    
    # Pattern 3: A-Pattern (A1, A2, A10)
    if (!mapped && grepl("^A\\d+$", response_char)) {
      numeric_code <- gsub("^A", "", response_char)  # A1 -> 1
      
      candidates <- c(
        response_char,  # "A1"
        numeric_code    # "1"
      )
      
      for (candidate in candidates) {
        if (candidate %in% names(labels)) {
          response_labels[response_char] <- labels[candidate]
          mapped_count <- mapped_count + 1
          if (verbose) cat("      ✓ A-Pattern:", response_char, "-> (via", candidate, ") ->", labels[candidate], "\n")
          mapped <- TRUE
          break
        }
      }
      
      if (mapped) next
    }
    
    # Pattern 4: Generisches Präfix-Pattern (beliebige Buchstaben + Zahlen)
    if (!mapped && grepl("^[A-Z]+\\d+$", response_char)) {
      numeric_code <- gsub("^[A-Z]+0*", "", response_char)
      
      if (numeric_code %in% names(labels)) {
        response_labels[response_char] <- labels[numeric_code]
        mapped_count <- mapped_count + 1
        if (verbose) cat("      ✓ Generisch:", response_char, "-> (via", numeric_code, ") ->", labels[numeric_code], "\n")
        mapped <- TRUE
        next
      }
    }
    
    if (!mapped && verbose) {
      cat("      ✗ Kein Match:", response_char, "\n")
    }
  }
  
  if (verbose) {
    cat("  ✓ Erfolgreich gemappt:", mapped_count, "von", length(unique_responses), "Responses\n")
  }
  
  return(response_labels)
}


# =============================================================================
# NEUE HILFSFUNKTION: LABEL-EXTRAKTION FÜR MATRIX-ITEMS
# =============================================================================

extract_item_label <- function(data, var_name, matrix_name) {
  "Extrahiert das echte Label einer Matrix-Variable aus verschiedenen Quellen"
  
  # 1. PRIORITÄT: Custom Variable Labels (explizit definiert)
  # SICHER: Prüfe ob custom_var_labels existiert (muss nicht immer definiert sein)
  if (exists("custom_var_labels", inherits = FALSE)) {
    if (var_name %in% names(custom_var_labels)) {
      custom_label <- custom_var_labels[[var_name]]
      if (!is.null(custom_label) && custom_label != "") {
        cat("  Gefundenes Custom Label:", custom_label, "\n")
        return(custom_label)
      }
    }
  }
  
  # 2. PRIORITÄT: Variable Labels aus Labelled Package (kurz)
  if (requireNamespace("labelled", quietly = TRUE)) {
    if (labelled::is.labelled(data[[var_name]])) {
      var_labels <- labelled::var_label(data[[var_name]])
      if (!is.null(var_labels) && var_labels != "") {
        # ACHTUNG: Labels von SPSS können sehr lang sein (komplette Fragen)
        # Kürze auf max 100 Zeichen
        shortened <- substr(var_labels, 1, 100)
        if (shortened != var_name && nchar(var_labels) > 10) {  # Nur wenn sinnvoll
          cat("  Gefundenes Labelled Label (gekürzt):", shortened, "...\n")
          return(shortened)
        }
      }
    }
  }
  
  # 3. PRIORITÄT: Variable Labels aus Attributen (mit intelligenter Längenbegrenzung)
  var_label <- attr(data[[var_name]], "label")
  if (!is.null(var_label) && var_label != "" && var_label != var_name) {
    # ACHTUNG: SPSS Labels können Fragetexte sein, nicht Matrix-Item-Labels
    # ABER: Verwende sie trotzdem mit intelligenter Kürzung
    
    # Prüfe ob Label ein Item-Label ist (nicht die Hauptfrage):
    # - Item-Labels sind typischerweise < 150 Zeichen
    # - Sie enthalten NICHT den kompletten Fragebeschreibung
    
    if (nchar(var_label) < 150) {
      # Kurz genug - verwende wie es ist
      cat("  Gefundenes Variable Label:", var_label, "\n")
      return(var_label)
    } else {
      # Zu lang - kürze auf ersten sinnvollen Satz/Teil
      # Versuche bis zum ersten Punkt oder 100 Zeichen
      first_sentence_end <- gregexpr("\\.", var_label)[[1]][1]
      if (first_sentence_end > 0 && first_sentence_end < 150) {
        shortened <- substr(var_label, 1, first_sentence_end)
      } else {
        shortened <- substr(var_label, 1, 100)
      }
      shortened <- trimws(shortened)
      cat("  Gefundenes Variable Label (gekürzt):", shortened, "...\n")
      return(shortened)
    }
  }
  
  # 4. PRIORITÄT: Intelligente Extraktion aus Variablennamen
  intelligent_label <- create_intelligent_label(var_name, matrix_name)
  if (intelligent_label != var_name) {
    cat("  Erstelltes intelligentes Label:", intelligent_label, "\n")
    return(intelligent_label)
  }
  
  # 5. FALLBACK: Formatierter Variablenname
  fallback_label <- create_fallback_label(var_name, matrix_name)
  cat("  Fallback Label:", fallback_label, "\n")
  return(fallback_label)
}

# =============================================================================
# INTELLIGENTE LABEL-ERSTELLUNG
# =============================================================================

create_intelligent_label <- function(var_name, matrix_name) {
  "Erstellt ein intelligentes Label basierend auf dem Variablennamen"
  
  # Entferne Matrix-Präfix und extrahiere bedeutungsvollen Teil
  clean_name <- var_name
  
  # Verschiedene Patterns versuchen
  patterns <- list(
    # ZS01[001] -> 001
    paste0("^", matrix_name, "\\[(.+)\\]$"),
    # ZS01.001. -> 001  
    paste0("^", matrix_name, "\\.(.+)\\.$"),
    # ZS01_001 -> 001
    paste0("^", matrix_name, "_(.+)$"),
    # ZS01-001 -> 001
    paste0("^", matrix_name, "-(.+)$")
  )
  
  extracted_part <- NULL
  for (pattern in patterns) {
    if (grepl(pattern, var_name)) {
      extracted_part <- gsub(pattern, "\\1", var_name)
      break
    }
  }
  
  if (is.null(extracted_part)) {
    return(var_name)  # Keine Extraktion möglich
  }
  
  # Versuche den extrahierten Teil zu interpretieren
  # Entferne führende Nullen für bessere Lesbarkeit
  clean_part <- gsub("^0+", "", extracted_part)
  if (clean_part == "") clean_part <- extracted_part  # Falls nur Nullen
  
  # Prüfe ob es eine Zahl ist
  if (grepl("^\\d+$", clean_part)) {
    return(paste("Item", clean_part))
  }
  
  # Prüfe auf spezielle Patterns oder Codes
  # Du kannst hier spezifische Übersetzungen für deine Survey-Items hinzufügen
  special_translations <- list(
    "001" = "Erstes Item",
    "002" = "Zweites Item", 
    "003" = "Drittes Item",
    # Füge hier weitere spezifische Übersetzungen hinzu
    "SQ001" = "Subquestion 1",
    "SQ002" = "Subquestion 2"
  )
  
  if (extracted_part %in% names(special_translations)) {
    return(special_translations[[extracted_part]])
  }
  
  # Fallback: Verwende den extrahierten Teil direkt
  return(paste("Item:", extracted_part))
}

create_fallback_label <- function(var_name, matrix_name) {
  "Erstellt ein Fallback-Label falls alle anderen Methoden fehlschlagen"
  
  # Entferne Matrix-Name und verwende den Rest
  short_name <- gsub(paste0("^", matrix_name), "", var_name)
  short_name <- gsub("^[._-]+", "", short_name)  # Entferne führende Trenner
  short_name <- gsub("[._-]+$", "", short_name)  # Entferne nachgestellte Trenner
  
  if (short_name == "" || short_name == var_name) {
    return(var_name)  # Kann nicht verkürzt werden
  }
  
  return(paste("Item", short_name))
}


# =============================================================================
# KREUZTABELLEN UND STATISTISCHE TESTS
# =============================================================================

detect_actual_data_type <- function(data, var_name) {
  "Erkennt den tatsächlichen Datentyp einer Variable aus den Daten"
  
  if (!var_name %in% names(data)) {
    return("unknown")
  }
  
  var_data <- data[[var_name]]
  
  # Entferne NA für Analyse
  var_data_clean <- var_data[!is.na(var_data)]
  
  if (length(var_data_clean) == 0) {
    return("unknown")
  }
  
  # 1. R-Datentyp prüfen
  if (is.numeric(var_data)) {
    # Prüfe ob es diskrete ganzzahlige Werte sind (könnte ordinal sein)
    unique_vals <- unique(var_data_clean)
    if (length(unique_vals) <= 10 && all(unique_vals == round(unique_vals))) {
      return("numeric_discrete")  # Numerisch aber wenige diskrete Werte
    } else {
      return("numeric")  # Kontinuierlich numerisch
    }
  }
  
  # 2. Für Character/Factor: Prüfe ob numerische Konvertierung möglich
  if (is.character(var_data) || is.factor(var_data)) {
    # Versuche numerische Konvertierung
    numeric_test <- suppressWarnings(as.numeric(as.character(var_data_clean)))
    successful_conversion <- sum(!is.na(numeric_test))
    total_values <- length(var_data_clean)
    
    # Falls > 80% der Werte numerisch konvertierbar sind
    if (successful_conversion / total_values > 0.8) {
      unique_numeric <- unique(numeric_test[!is.na(numeric_test)])
      if (length(unique_numeric) <= 10 && all(unique_numeric == round(unique_numeric))) {
        return("numeric_discrete")
      } else {
        return("numeric")
      }
    } else {
      # Nicht numerisch konvertierbar → nominal
      unique_vals <- unique(var_data_clean)
      if (length(unique_vals) <= 2) {
        return("nominal_binary")
      } else {
        return("nominal")
      }
    }
  }
  
  return("unknown")
}

# Erweiterte create_labeled_factor Funktion mit umfassender Label-Suche
create_labeled_factor <- function(data, var_name, config) {
  "Erstellt einen Factor mit Labels aus verschiedenen Quellen (inkl. RDS mit Umkehr)"
  
  # NEUER FIX: Überspringe numerische Variablen
  if (is.numeric(data[[var_name]])) {
    return(data[[var_name]])  # Gib numerische Variable unverändert zurück
  }
  
  # Originale Werte
  original_values <- data[[var_name]]
  
  # NEUE PRIORISIERUNG: Nutze get_value_labels_with_priority
  labels <- get_value_labels_with_priority(data, var_name, config)
  labels_found <- !is.null(labels) && length(labels) > 0
  
  # Labels anwenden falls gefunden
  if (labels_found) {
    cat("  ✓ Labels gefunden für", var_name, ":", length(labels), "Labels\n")
    
    # Erstelle gelabelte Werte
    labeled_values <- as.character(original_values)
    mapped_count <- 0
    
    for (code in names(labels)) {
      label <- labels[code]
      
      # Direkte Übereinstimmung
      matches <- labeled_values == code
      if (any(matches, na.rm = TRUE)) {
        labeled_values[matches & !is.na(matches)] <- label
        mapped_count <- mapped_count + sum(matches, na.rm = TRUE)
      }
      
      # AO-Pattern: AO01 -> auch "1" mappen
      if (grepl("^AO\\d+$", code)) {
        numeric_code <- as.character(as.numeric(gsub("^AO0*", "", code)))
        matches_num <- labeled_values == numeric_code
        if (any(matches_num, na.rm = TRUE)) {
          labeled_values[matches_num & !is.na(matches_num)] <- label
          mapped_count <- mapped_count + sum(matches_num, na.rm = TRUE)
        }
      }
      
      # A-Pattern: A1 -> auch "1" mappen
      if (grepl("^A\\d+$", code)) {
        numeric_code <- gsub("^A", "", code)
        matches_num <- labeled_values == numeric_code
        if (any(matches_num, na.rm = TRUE)) {
          labeled_values[matches_num & !is.na(matches_num)] <- label
          mapped_count <- mapped_count + sum(matches_num, na.rm = TRUE)
        }
      }
    }
    
    cat("    Gemappt:", mapped_count, "Werte\n")
    
    # Erstelle Factor mit Labels
    return(as.factor(labeled_values))
  } else {
    cat("  ⚠  Keine Labels für", var_name, "- verwende rohe Werte\n")
    return(as.factor(original_values))
  }
}


# Neue Funktion: Prüfe ob Variable eine Matrix ist (basierend auf bestehender Logik)
is_matrix_variable <- function(var_name, data, config) {
  "Prüft ob eine Variable eine Matrix-Variable ist"
  
  # 1. Erst in Config schauen
  var_config <- config$variablen[config$variablen$variable_name == var_name, ]
  if (nrow(var_config) > 0 && var_config$data_type[1] == "matrix") {
    return(TRUE)
  }
  
  # 2. Dann in Daten nach Matrix-Items suchen (nutzt bestehende Logik)
  matrix_patterns <- c(
    paste0("^", var_name, "\\[.+\\]$"),            
    paste0("^", var_name, "\\..+\\.$"),            
    paste0("^", var_name, "_(SQ[0-9]+|[0-9]+)$"), 
    paste0("^", var_name, "-.+$")                  
  )
  
  matrix_vars <- c()
  for (pattern in matrix_patterns) {
    found_vars <- names(data)[grepl(pattern, names(data))]
    matrix_vars <- c(matrix_vars, found_vars)
  }
  
  # Filter out [other] variables (bestehende Logik)
  matrix_vars <- matrix_vars[!grepl("other", matrix_vars, ignore.case = TRUE)]
  matrix_vars <- unique(matrix_vars)
  
  return(length(matrix_vars) >= 2)  # Mindestens 2 Items für Matrix
}

# Neue Funktion: Matrix-Kreuztabelle erstellen
create_matrix_crosstab <- function(data, matrix_var, group_var, survey_obj = NULL, config = NULL) {
  "Erstellt Kreuztabelle für Matrix-Variable vs Gruppenvariable"
  
  cat("Erstelle Matrix-Kreuztabelle:", matrix_var, "x", group_var, "\n")
  
  # Finde Matrix-Items (bestehende Logik...)
  possible_matrix_names <- c(
    matrix_var,
    gsub("\\.", "", matrix_var),
    gsub("\\.$", "", matrix_var),
    gsub("_", "", matrix_var)
  )
  
  matrix_vars <- c()
  actual_matrix_name <- matrix_var
  
  for (test_name in possible_matrix_names) {
    matrix_patterns <- c(
      paste0("^", test_name, "\\[.+\\]$"),
      paste0("^", test_name, "\\..+\\.$"),
      paste0("^", test_name, "_.+$"),
      paste0("^", test_name, "-.+$")
    )
    
    test_matrix_vars <- c()
    for (pattern in matrix_patterns) {
      found_vars <- names(data)[grepl(pattern, names(data))]
      test_matrix_vars <- c(test_matrix_vars, found_vars)
    }
    
    test_matrix_vars <- test_matrix_vars[!grepl("other", test_matrix_vars, ignore.case = TRUE)]
    test_matrix_vars <- unique(test_matrix_vars)
    
    if (length(test_matrix_vars) >= 2) {
      matrix_vars <- test_matrix_vars
      actual_matrix_name <- test_name
      cat("Matrix-Items gefunden mit Basis-Name:", actual_matrix_name, "\n")
      break
    }
  }
  
  matrix_vars <- sort(matrix_vars)
  
  if (length(matrix_vars) == 0) {
    cat("WARNUNG: Keine Matrix-Items gefunden für", matrix_var, "\n")
    return(NULL)
  }
  
  cat("Gefundene Matrix-Items:", length(matrix_vars), "\n")
  
  # Vollständige Fälle für Matrix + Gruppe
  all_vars_needed <- c(matrix_vars, group_var)
  complete_cases <- complete.cases(data[, all_vars_needed])
  complete_data <- data[complete_cases, ]
  
  if (nrow(complete_data) < 5) {
    cat("WARNUNG: Zu wenige vollständige Fälle für Matrix-Kreuztabelle\n")
    return(NULL)
  }
  
  cat("Vollständige Fälle:", nrow(complete_data), "\n")
  
  # Gruppe-Variable mit Labels erstellen
  group_display_var <- paste0(group_var, "_labeled")
  complete_data[[group_display_var]] <- create_labeled_factor(complete_data, group_var, config)
  
  # Eindeutige Gruppen ermitteln
  unique_groups <- levels(complete_data[[group_display_var]])
  cat("Gruppen:", paste(unique_groups, collapse = ", "), "\n")
  
  # Matrix-Konfiguration für Kodierung finden
  matrix_config <- config$variablen[config$variablen$variable_name == matrix_var, ]
  matrix_coding <- if(nrow(matrix_config) > 0) matrix_config$coding[1] else NA
  
  # 1. KATEGORIALE TABELLE (nur absolute Werte)
  categorical_table <- create_matrix_categorical_crosstab(
    complete_data, matrix_vars, group_display_var, unique_groups, matrix_coding, survey_obj, actual_matrix_name
  )
  
  # 2. NUMERISCHE TABELLE (falls Kodierung vorhanden)
  numeric_table <- NULL
  if (!is.na(matrix_coding) && matrix_coding != "") {
    
    # *** KOPIERE DIE LOGIK AUS create_matrix_table() ***
    
    # PRÜFE OB KODIERUNG VORHANDEN IST (ordinal behandeln) ODER DICHOTOM ERKANNT
    has_coding <- !is.na(matrix_coding) && matrix_coding != ""
    
    # NEUE LOGIK: Erkenne ordinale Matrix basierend auf Kodierung (ANALOG zu create_matrix_table)
    is_ordinal_matrix <- FALSE
    if (has_coding) {
      labels <- parse_coding(matrix_coding)  # <-- KORRIGIERT: Verwende matrix_coding statt var_config$coding
      if (!is.null(labels) && length(labels) > 2) {
        # Prüfe ob Labels numerische Codes haben (ordinal)
        numeric_codes <- suppressWarnings(as.numeric(names(labels)))
        if (!any(is.na(numeric_codes)) && length(unique(numeric_codes)) > 2) {
          is_ordinal_matrix <- TRUE
          cat("Ordinale Matrix erkannt basierend auf numerischen Codes in Kodierung\n")
        }
      }
    }
    
    # Erkenne dichotome Matrix (ANALOG zu create_matrix_table)
    is_dichotomous_matrix <- FALSE
    if (!is.null(labels) && length(labels) <= 3) {  # Max 3 Kategorien für dichotom
      label_keys <- names(labels)
      
      # Pattern 1: Y/N in Kodierung
      if (any(c("Y", "N") %in% label_keys) || any(c("1", "0") %in% label_keys)) {
        is_dichotomous_matrix <- TRUE
        cat("Dichotome Matrix erkannt (Y/N Pattern in Kodierung)\n")
      }
    }
    
    # *** ERWEITERTE BEDINGUNG: Erstelle numerische Tabelle für ordinale UND binäre Matrices ***
    if (is_ordinal_matrix || is_dichotomous_matrix) {
      numeric_table <- create_matrix_numeric_crosstab(
        complete_data, matrix_vars, group_display_var, unique_groups, matrix_coding, matrix_config, actual_matrix_name
      )
    } else {
      cat("Matrix ist weder ordinal noch binär - keine numerische Tabelle erstellt\n")
    }
    
  } else {
    # *** FALLBACK: Wenn keine Kodierung in Config, analysiere Datenwerte direkt ***
    cat("Keine Kodierung in Config gefunden, analysiere Datenwerte...\n")
    
    # Sammle alle Datenwerte aus Matrix-Items
    all_data_values <- c()
    for (var in matrix_vars) {
      var_values <- complete_data[[var]][!is.na(complete_data[[var]])]
      all_data_values <- c(all_data_values, var_values)
    }
    
    # Eindeutige Kategorien ermitteln
    unique_responses <- unique(all_data_values)
    unique_responses <- unique_responses[!is.na(unique_responses) & unique_responses != ""]
    
    cat("Gefundene Datenwerte:", paste(head(unique_responses, 10), collapse = ", "), "\n")
    
    # *** VERBESSERTE ORDINAL-ERKENNUNG ***
    
    # Prüfe ob Werte das Format "Zahl (Text)" haben (ordinal)
    ordinal_pattern <- "^\\d+(\\s*\\(.*\\))?$"
    ordinal_matches <- str_detect(unique_responses, ordinal_pattern)
    
    # Entferne "Weiß nicht" und ähnliche aus der Ordinal-Prüfung
    non_ordinal_patterns <- c("Weiß nicht", "Keine Angabe", "N/A", "k.A.", "Missing")
    ordinal_responses <- unique_responses[!unique_responses %in% non_ordinal_patterns]
    
    cat("Gefilterte ordinale Responses:", paste(ordinal_responses, collapse = ", "), "\n")
    
    # Prüfe ob MINDESTENS 3 ordinale Werte vorhanden sind (nicht alle)
    ordinal_count <- sum(str_detect(ordinal_responses, ordinal_pattern))
    total_count <- length(ordinal_responses)
    
    cat("Ordinale Pattern gefunden:", ordinal_count, "von", total_count, "Werten\n")
    
    # *** NEUE BEDINGUNG: Mindestens 3 ordinale Werte (statt alle) ***
    if (ordinal_count >= 3) {
      # Extrahiere numerische Codes nur von ordinalen Werten
      ordinal_values <- ordinal_responses[str_detect(ordinal_responses, ordinal_pattern)]
      numeric_codes <- str_extract(ordinal_values, "^\\d+")
      numeric_codes <- as.numeric(numeric_codes)
      
      cat("Extrahierte numerische Codes:", paste(numeric_codes, collapse = ", "), "\n")
      
      if (!any(is.na(numeric_codes)) && length(unique(numeric_codes)) >= 3) {
        cat("Matrix erkannt als ordinal (durch Datenwert-Analyse)\n")
        
        numeric_table <- create_matrix_numeric_crosstab(
          complete_data, matrix_vars, group_display_var, unique_groups, matrix_coding, matrix_config, actual_matrix_name
        )
      } else {
        cat("Numerische Codes nicht eindeutig ordinal\n")
      }
    } else if (all(unique_responses %in% c("", "1")) || all(unique_responses %in% c("1"))) {
      # Dichotome Matrix: Nur "1" und leere Werte
      cat("Matrix erkannt als dichotom (1/leer Pattern in Datenwerten)\n")
      
      numeric_table <- create_matrix_numeric_crosstab(
        complete_data, matrix_vars, group_display_var, unique_groups, matrix_coding, matrix_config, actual_matrix_name
      )
    } else {
      cat("Matrix-Datenwerte sind weder ordinal noch dichotom - keine numerische Tabelle\n")
      cat("Alle Werte:", paste(unique_responses, collapse = ", "), "\n")
    }
  }
  
  # Rückgabe-Struktur bleibt unverändert
  result <- list(
    categorical = categorical_table,
    numeric = numeric_table,  # Explizit als separate Komponente
    n_total = nrow(complete_data),
    var1_name = matrix_var,
    var2_name = group_var,
    var1_type = "matrix",
    var2_type = detect_actual_data_type(complete_data, group_var),
    matrix_items = matrix_vars,
    groups = unique_groups
  )
  
  return(result)
}
# Hilfsfunktion: Kategoriale Matrix-Kreuztabelle
create_matrix_categorical_crosstab <- function(data, matrix_vars, group_var, unique_groups, matrix_coding, survey_obj, actual_matrix_name) {
  "Erstellt kategoriale Kreuztabelle für Matrix-Items - NUR ABSOLUTE WERTE"
  
  # Alle Antwortkategorien sammeln (bestehende Logik)
  all_responses <- c()
  for (var in matrix_vars) {
    var_responses <- unique(data[[var]][!is.na(data[[var]])])
    all_responses <- c(all_responses, var_responses)
  }
  unique_responses <- unique(all_responses)
  unique_responses <- unique_responses[!is.na(unique_responses) & unique_responses != ""]
  unique_responses <- sort_response_categories(unique_responses)
  
  # Labels aus Kodierung erstellen (bestehende Logik)
  response_labels <- unique_responses
  names(response_labels) <- unique_responses
  
  # *** VEREINFACHTE LABEL-EXTRAKTION MIT NEUER HILFSFUNKTION ***
  labels <- get_matrix_labels(data, matrix_vars, actual_matrix_name, NULL, matrix_coding)
  
  if (!is.null(labels) && length(labels) > 0) {
    cat("  Labels für Matrix-Kreuztabelle gefunden:", length(labels), "Labels\n")
    # Verwende zentralisierte Mapping-Funktion
    response_labels <- map_response_labels(unique_responses, labels, verbose = FALSE)
  }
  
  # *** GEÄNDERT: Survey-Objekt für Gewichtung verwenden ***
  if (!is.null(survey_obj) && WEIGHTS) {
    # Gewichtete Analyse: Survey-Objekt mit aktuellen Daten neu erstellen
    survey_obj_current <- create_survey_object(data, WEIGHT_VAR)
    cat("Verwende gewichtete Matrix-Kreuztabelle\n")
  }
  
  # Tabelle für jedes Matrix-Item erstellen
  result_rows <- list()
  
  for (var in matrix_vars) {
    item_label <- extract_item_label(data, var, actual_matrix_name)
    result_row <- data.frame(Item = item_label, stringsAsFactors = FALSE)
    
    # Für jede Gruppe die Häufigkeiten berechnen
    for (group in unique_groups) {
      group_data <- data[data[[group_var]] == group & !is.na(data[[group_var]]), ]
      
      if (nrow(group_data) > 0) {
        # *** GEÄNDERT: Gewichtete vs. ungewichtete Häufigkeiten ***
        if (!is.null(survey_obj) && WEIGHTS) {
          # Gewichtete Häufigkeiten
          group_survey <- subset(survey_obj_current, get(group_var) == group & !is.na(get(group_var)))
          
          if (nrow(group_survey$variables) > 0) {
            freq_table <- svytable(as.formula(paste("~", var)), group_survey)
            freq_df <- data.frame(
              response = names(freq_table),
              count = as.numeric(freq_table),
              stringsAsFactors = FALSE
            )
          } else {
            freq_df <- data.frame(response = character(), count = numeric(), stringsAsFactors = FALSE)
          }
        } else {
          # Ungewichtete Häufigkeiten
          item_values <- group_data[[var]]
          freq_table <- table(item_values, useNA = "no")
          freq_df <- data.frame(
            response = names(freq_table),
            count = as.numeric(freq_table),
            stringsAsFactors = FALSE
          )
        }
        
        # *** GEÄNDERT: NUR ABSOLUTE WERTE, KEINE PROZENTE ***
        for (response in unique_responses) {
          count <- if(as.character(response) %in% freq_df$response) {
            freq_df$count[freq_df$response == as.character(response)]
          } else {
            0
          }
          
          # Spaltenname mit Label
          response_label <- response_labels[as.character(response)]
          clean_response <- make_clean_colname(response_label) 
          clean_response <- make_clean_colname(response_label)
          col_name <- paste0(group, "_", clean_response)
          
          # *** NUR ABSOLUTE WERTE (keine Prozente mehr) ***
          result_row[[col_name]] <- count
        }
        
        # Total für diese Gruppe
        result_row[[paste0(group, "_Total")]] <- sum(freq_df$count)
      }
    }
    
    result_rows[[var]] <- result_row
  }
  
  # Alle Zeilen zusammenfügen
  if (length(result_rows) > 0) {
    result_table <- do.call(rbind, result_rows)
    rownames(result_table) <- NULL
    return(result_table)
  }
  
  return(NULL)
}

# Hilfsfunktion: Numerische Matrix-Kreuztabelle
create_matrix_numeric_crosstab <- function(data, matrix_vars, group_var, unique_groups, matrix_coding, matrix_config, actual_matrix_name) {
  "Erstellt numerische Kreuztabelle für Matrix-Items - MIT DATENWERT-UNTERSTÜTZUNG"
  
  # *** NEUE LOGIK: Unterscheide zwischen Config-Kodierung und Datenwert-Extraktion ***
  
  # Falls keine Config-Kodierung vorhanden, analysiere Datenwerte
  if (is.na(matrix_coding) || matrix_coding == "") {
    cat("Keine Config-Kodierung - verwende Datenwert-Extraktion\n")
    use_data_extraction <- TRUE
  } else {
    cat("Config-Kodierung vorhanden - verwende Standard-Extraktion\n")
    use_data_extraction <- FALSE
  }
  
  # *** REST DER FUNKTION BLEIBT UNVERÄNDERT BIS ZUR NUMERISCHEN KONVERTIERUNG ***
  
  result_rows <- list()
  
  for (var in matrix_vars) {
    item_label <- extract_item_label(data, var, actual_matrix_name)
    
    # *** ERWEITERTE NUMERISCHE KONVERTIERUNG ***
    var_data <- data[[var]]
    
    if (use_data_extraction) {
      # NEUE METHODE: Extrahiere Zahlen direkt aus "5 (stimme voll und ganz zu)" Format
      cat("  Datenwert-Extraktion für", var, "\n")
      
      numeric_values <- rep(NA, length(var_data))
      
      for (i in seq_along(var_data)) {
        if (!is.na(var_data[i]) && var_data[i] != "") {
          value <- as.character(var_data[i])
          
          # Extrahiere Zahl am Anfang
          if (str_detect(value, "^\\d+")) {
            extracted_number <- as.numeric(str_extract(value, "^\\d+"))
            if (!is.na(extracted_number)) {
              numeric_values[i] <- extracted_number
            }
          }
        }
      }
      
      cat("    Erfolgreich extrahiert:", sum(!is.na(numeric_values)), "von", length(numeric_values), "Werten\n")
      
    } else {
      # BESTEHENDE METHODE: Verwende extract_numeric_from_matrix_coding
      numeric_values <- extract_numeric_from_matrix_coding(
        var_data, 
        matrix_coding, 
        if(nrow(matrix_config) > 0) matrix_config$min_value[1] else NA, 
        if(nrow(matrix_config) > 0) matrix_config$max_value[1] else NA
      )
    }
    
    # *** REST DER STATISTIK-BERECHNUNG BLEIBT UNVERÄNDERT ***
    
    # Zeile für dieses Item
    result_row <- data.frame(Item = item_label, stringsAsFactors = FALSE)
    
    # Für jede Gruppe Statistiken berechnen
    for (group in unique_groups) {
      group_indices <- data[[group_var]] == group & !is.na(data[[group_var]])
      group_numeric <- numeric_values[group_indices]
      group_numeric <- group_numeric[!is.na(group_numeric)]
      
      if (length(group_numeric) > 0) {
        # Ungewichtete Statistiken (Survey-Gewichtung für Matrix-Items ist komplex)
        group_mean <- mean(group_numeric, na.rm = TRUE)
        group_median <- median(group_numeric, na.rm = TRUE)
        group_sd <- sd(group_numeric, na.rm = TRUE)
        group_n <- length(group_numeric)
        
        result_row[[paste0(group, "_Mean")]] <- round(group_mean, DIGITS_ROUND)
        result_row[[paste0(group, "_Median")]] <- round(group_median, DIGITS_ROUND)
        result_row[[paste0(group, "_SD")]] <- round(group_sd, DIGITS_ROUND)
        result_row[[paste0(group, "_N")]] <- group_n
      } else {
        result_row[[paste0(group, "_Mean")]] <- NA
        result_row[[paste0(group, "_Median")]] <- NA
        result_row[[paste0(group, "_SD")]] <- NA
        result_row[[paste0(group, "_N")]] <- 0
      }
    }
    
    result_rows[[var]] <- result_row
  }
  
  # Alle Zeilen zusammenfügen
  if (length(result_rows) > 0) {
    result_table <- do.call(rbind, result_rows)
    rownames(result_table) <- NULL
    return(result_table)
  }
  
  return(NULL)
}

create_contingency_table <- function(data, var1, var2, survey_obj = NULL, config = NULL) {
  
  # MATRIX-ERKENNUNG MIT SOFORTIGEM RETURN (KORRIGIERT)
  var1_is_matrix <- is_matrix_variable(var1, data, config)
  var2_is_matrix <- is_matrix_variable(var2, data, config)
  
  if (var1_is_matrix && !var2_is_matrix) {
    cat("Matrix-Kreuztabelle erkannt:", var1, "(Matrix) x", var2, "(Gruppe)\n")
    return(create_matrix_crosstab(data, var1, var2, survey_obj, config))
  } else if (var2_is_matrix && !var1_is_matrix) {
    cat("Matrix-Kreuztabelle erkannt:", var2, "(Matrix) x", var1, "(Gruppe)\n")
    matrix_result <- create_matrix_crosstab(data, var2, var1, survey_obj, config)
    if (!is.null(matrix_result)) {
      matrix_result$var1_name <- var1
      matrix_result$var2_name <- var2
    }
    return(matrix_result)
  } else if (var1_is_matrix && var2_is_matrix) {
    cat("WARNUNG: Beide Variablen sind Matrizen - nicht unterstützt\n")
    return(NULL)
  }
  
  # AB HIER: NUR NOCH NORMALE KREUZTABELLEN (var1 und var2 existieren als einzelne Variablen)
  
  # Prüfe ob normale Variablen existieren
  if (!var1 %in% names(data) || !var2 %in% names(data)) {
    warning(paste("Normale Variablen nicht gefunden:", var1, "oder", var2))
    return(NULL)
  }
  
  # AUTOMATISCHE TYP-ERKENNUNG AUS DEN DATEN
  var1_actual_type <- detect_actual_data_type(data, var1)
  var2_actual_type <- detect_actual_data_type(data, var2)
  
  cat("Automatische Typ-Erkennung:\n")
  cat("  ", var1, "→", var1_actual_type, "\n")
  cat("  ", var2, "→", var2_actual_type, "\n")
  
  # Daten ohne fehlende Werte für beide Variablen - ERST HIER FILTERN
  complete_data <- data[!is.na(data[[var1]]) & !is.na(data[[var2]]), ]
  
  if (nrow(complete_data) == 0) {
    warning(paste("Keine vollständigen Daten für", var1, "x", var2))
    return(NULL)
  }
  
  # Labels aus Konfiguration anwenden - DIREKT IN complete_data
  var1_display <- var1
  var2_display <- var2
  
  if (!is.null(config)) {
    # Erstelle neue Variablen mit Labels direkt als Factor-Levels
    var1_display <- paste0(var1, "_labeled")
    var2_display <- paste0(var2, "_labeled")
    
    cat("Erstelle gelabelte Variablen:\n")
    cat("  ", var1, "→", var1_display, "\n")
    cat("  ", var2, "→", var2_display, "\n")
    
    # DIREKT IN complete_data erstellen - nicht in ursprünglichen data
    complete_data[[var1_display]] <- create_labeled_factor(complete_data, var1, config)
    cat("  var1_display Levels:", paste(levels(complete_data[[var1_display]]), collapse = ", "), "\n")
    
    complete_data[[var2_display]] <- create_labeled_factor(complete_data, var2, config)
    cat("  var2_display Levels:", paste(levels(complete_data[[var2_display]]), collapse = ", "), "\n")
    
    cat("  ✓ Gelabelte Factors erstellt für Kreuztabelle\n")
  }
  
  # ENTSCHEIDUNG BASIEREND AUF TATSÄCHLICHEN DATENTYPEN
  var1_is_numeric <- var1_actual_type %in% c("numeric", "numeric_discrete")
  var2_is_numeric <- var2_actual_type %in% c("numeric", "numeric_discrete")
  
  # NEUER FIX: Factor-zu-Numerisch Konvertierung für "numeric_discrete"
  if (var1_is_numeric && is.factor(complete_data[[var1]])) {
    cat("→ Konvertiere Factor", var1, "zu numerisch (erkannt als", var1_actual_type, ")\n")
    complete_data[[var1]] <- convert_factor_to_numeric_safe(complete_data[[var1]], var1)
  }
  
  if (var2_is_numeric && is.factor(complete_data[[var2]])) {
    cat("→ Konvertiere Factor", var2, "zu numerisch (erkannt als", var2_actual_type, ")\n")
    complete_data[[var2]] <- convert_factor_to_numeric_safe(complete_data[[var2]], var2)
  }
  
  # Konvertiere zu numerisch falls nötig (bestehende Logik)
  if (var1_is_numeric && !is.numeric(complete_data[[var1]])) {
    complete_data[[var1]] <- suppressWarnings(as.numeric(as.character(complete_data[[var1]])))
  }
  if (var2_is_numeric && !is.numeric(complete_data[[var2]])) {
    complete_data[[var2]] <- suppressWarnings(as.numeric(as.character(complete_data[[var2]])))
  }
  
  # Just-in-Time Factor-Konvertierung für kategoriale Analysen
  if (!var1_is_numeric && !is.factor(complete_data[[var1_display]])) {
    complete_data[[var1_display]] <- as.factor(complete_data[[var1_display]])
  }
  if (!var2_is_numeric && !is.factor(complete_data[[var2_display]])) {
    complete_data[[var2_display]] <- as.factor(complete_data[[var2_display]])
  }
  
  # NEUE LOGIK: Beide numerisch → Korrelationsanalyse statt Kreuztabelle
  if (var1_is_numeric && var2_is_numeric) {
    cat("→ Beide Variablen numerisch: Erstelle Korrelationsanalyse statt Kreuztabelle\n")
    
    # Nutze bestehende perform_correlation_test Funktion
    correlation_result <- perform_correlation_test(complete_data, var1, var2, survey_obj)
    
    # Konvertiere in kompatibles Format
    correlation_table <- data.frame(
      Kennwert = c("Test", "Korrelationskoeffizient", "p-Wert", "Ergebnis", "Interpretation"),
      Wert = c(
        correlation_result$test,
        as.character(correlation_result$statistic),
        as.character(correlation_result$p_value),
        correlation_result$result,
        if("interpretation" %in% names(correlation_result)) correlation_result$interpretation else ""
      ),
      stringsAsFactors = FALSE
    )
    
    return(list(
      correlation_table = correlation_table,
      n_total = nrow(complete_data),
      var1_name = var1,
      var2_name = var2,
      type = "correlation"
    ))
  }
  
  # Bestimme ob eine Variable numerisch und die andere kategorisch ist
  if (var1_is_numeric && !var2_is_numeric) {
    # var1 numerisch, var2 gruppierend (mit Labels)
    cat("→ Erstelle Gruppenmittelwerte:", var1, "gruppiert nach", var2, "(mit Labels)\n")
    return(create_group_means_table(complete_data, var1, var2_display, survey_obj))
  } else if (var2_is_numeric && !var1_is_numeric) {
    # var2 numerisch, var1 gruppierend (mit Labels)
    cat("→ Erstelle Gruppenmittelwerte:", var2, "gruppiert nach", var1, "(mit Labels)\n")
    return(create_group_means_table(complete_data, var2, var1_display, survey_obj))
  }
  
  # Standard Kreuztabelle für kategoriale Variablen
  cat("→ Erstelle Standard-Kreuztabelle mit Labels\n")
  
  if (!is.null(survey_obj) && WEIGHTS) {
    # Survey-Objekt für complete_data neu erstellen
    survey_complete <- create_survey_object(complete_data, WEIGHT_VAR)
    
    # Gewichtete Kreuztabelle - FIX: Verwende die Display-Variablen
    crosstab <- svytable(as.formula(paste("~", var1_display, "+", var2_display)), survey_complete)
    
    # Randverteilungen
    margin1 <- svytable(as.formula(paste("~", var1_display)), survey_complete)
    margin2 <- svytable(as.formula(paste("~", var2_display)), survey_complete)
    
  } else {
    # Ungewichtete Kreuztabelle - FIX: Verwende die Display-Variablen
    crosstab <- table(complete_data[[var1_display]], complete_data[[var2_display]])
    margin1 <- table(complete_data[[var1_display]])
    margin2 <- table(complete_data[[var2_display]])
  }
  
  # In Data Frame konvertieren
  crosstab_df <- as.data.frame.matrix(crosstab)
  
  # Randverteilungen hinzufügen
  crosstab_df$Gesamt <- rowSums(crosstab_df)
  crosstab_df <- rbind(crosstab_df, 
                       c(colSums(crosstab_df[, -ncol(crosstab_df)]), sum(crosstab_df$Gesamt)))
  
  # Zeilen- und Spaltennamen setzen
  rownames(crosstab_df)[nrow(crosstab_df)] <- "Gesamt"
  
  # Relative Häufigkeiten berechnen (Spaltenprozente)
  crosstab_rel <- crosstab_df
  for (j in 1:(ncol(crosstab_rel)-1)) {  # Über Spalten iterieren, letzte Spalte ausschließen
    total_col <- crosstab_rel[nrow(crosstab_rel), j]  # Spaltensumme aus der Gesamt-Zeile
    if (total_col > 0) {
      crosstab_rel[-nrow(crosstab_rel), j] <- round(
        crosstab_rel[-nrow(crosstab_rel), j] / total_col * 100, DIGITS_ROUND
      )
    }
  }
  
  # Spaltenprozente für Gesamt-Zeile
  total_col <- crosstab_rel$Gesamt[nrow(crosstab_rel)]
  if (total_col > 0) {
    crosstab_rel[nrow(crosstab_rel), -ncol(crosstab_rel)] <- round(
      crosstab_rel[nrow(crosstab_rel), -ncol(crosstab_rel)] / total_col * 100, DIGITS_ROUND
    )
  }
  crosstab_rel$Gesamt[1:(nrow(crosstab_rel)-1)] <- 100
  crosstab_rel$Gesamt[nrow(crosstab_rel)] <- 100
  
  return(list(
    absolute = crosstab_df,
    relative = crosstab_rel,
    n_total = sum(crosstab_df$Gesamt[nrow(crosstab_df)]),
    var1_name = var1,
    var2_name = var2,
    var1_type = var1_actual_type,
    var2_type = var2_actual_type
  ))
}


# NEUE HILFSFUNKTION: Sichere Factor-zu-Numerisch Konvertierung
convert_factor_to_numeric_safe <- function(factor_var, var_name) {
  "Sichere Konvertierung von Factor zu numerisch mit AO-Pattern Unterstützung"
  
  if (!is.factor(factor_var)) {
    return(factor_var)  # Bereits nicht-Factor
  }
  
  factor_levels <- levels(factor_var)
  cat("    Factor Levels für", var_name, ":", paste(factor_levels, collapse = ", "), "\n")
  
  # *** NEU: AO-Pattern Erkennung in Factor Levels ***
  # Prüfe ob Levels AO-Pattern enthalten
  ao_pattern_detected <- any(grepl("^AO\\d+", factor_levels))
  
  if (ao_pattern_detected) {
    cat("    → AO-Pattern in Factor Levels erkannt\n")
    
    # Konvertiere AO-Pattern zu numerischen Werten
    numeric_result <- rep(NA, length(factor_var))
    
    for (i in seq_along(factor_var)) {
      if (!is.na(factor_var[i])) {
        level_value <- as.character(factor_var[i])
        
        # Extrahiere Nummer aus AO-Pattern
        if (grepl("^AO\\d+", level_value)) {
          ao_number <- gsub("^AO0*", "", level_value)
          numeric_value <- suppressWarnings(as.numeric(ao_number))
          if (!is.na(numeric_value)) {
            numeric_result[i] <- numeric_value
          }
        } else {
          # Fallback für nicht-AO Werte
          numeric_result[i] <- suppressWarnings(as.numeric(level_value))
        }
      }
    }
    
    cat("    → AO-Pattern Konvertierung erfolgreich:", sum(!is.na(numeric_result)), "Werte\n")
    return(numeric_result)
  }
  
  # Bestehende Logik für normale Factor-Levels
  if (all(grepl("^\\d+(\\.\\d+)?$", factor_levels))) {
    result <- as.numeric(as.character(factor_var))
    cat("    → Konvertierung über numerische Levels erfolgreich\n")
    return(result)
  } else {
    result <- as.numeric(factor_var)
    cat("    → Konvertierung über Level-Position erfolgreich\n")
    return(result)
  }
}

# Gruppenmittelwerte-Funktion (ERWEITERT mit Validierung)
create_group_means_table <- function(data, numeric_var, group_var, survey_obj = NULL) {
  
  cat("Erstelle Gruppenmittelwerte für", numeric_var, "gruppiert nach", group_var, "\n")
  
  # NEUE VALIDIERUNG: Prüfe beide Variablen auf Varianz
  # 1. Prüfe numerische Variable
  valid_numeric <- data[[numeric_var]][!is.na(data[[numeric_var]])]
  if (length(valid_numeric) == 0) {
    cat("FEHLER: Numerische Variable", numeric_var, "hat keine gültigen Werte\n")
    return(NULL)
  }
  
  # NEUER FIX: Konvertiere Factor zu numerisch falls nötig
  if (is.factor(data[[numeric_var]])) {
    cat("KONVERTIERE Factor", numeric_var, "zu numerisch für Gruppenmittelwerte\n")
    
    # Versuche intelligente Konvertierung
    factor_levels <- levels(data[[numeric_var]])
    cat("Factor Levels:", paste(factor_levels, collapse = ", "), "\n")
    
    # Strategie 1: Levels sind bereits numerisch (z.B. "1", "2", "3")
    if (all(grepl("^\\d+(\\.\\d+)?$", factor_levels))) {
      data[[numeric_var]] <- as.numeric(as.character(data[[numeric_var]]))
      cat("→ Konvertierung über numerische Levels erfolgreich\n")
    } else {
      # Strategie 2: Verwende Level-Position als numerische Werte
      data[[numeric_var]] <- as.numeric(data[[numeric_var]])
      cat("→ Konvertierung über Level-Position erfolgreich\n")
    }
    
    # Neue Validierung nach Konvertierung
    valid_numeric <- data[[numeric_var]][!is.na(data[[numeric_var]])]
    if (length(valid_numeric) == 0) {
      cat("FEHLER: Konvertierung zu numerisch fehlgeschlagen\n")
      return(NULL)
    }
  } else if (!is.numeric(data[[numeric_var]])) {
    # Variable ist weder Factor noch numerisch → versuche direkte Konvertierung
    cat("KONVERTIERE", class(data[[numeric_var]])[1], "Variable", numeric_var, "zu numerisch\n")
    data[[numeric_var]] <- suppressWarnings(as.numeric(as.character(data[[numeric_var]])))
    
    valid_numeric <- data[[numeric_var]][!is.na(data[[numeric_var]])]
    if (length(valid_numeric) == 0) {
      cat("FEHLER: Variable", numeric_var, "kann nicht zu numerisch konvertiert werden\n")
      return(NULL)
    }
  }
  
  if (length(unique(valid_numeric)) < 2) {
    cat("FEHLER: Numerische Variable", numeric_var, "hat keine Varianz (alle Werte gleich)\n")
    return(NULL)
  }
  
  # 2. Prüfe Gruppenvariable
  valid_groups <- data[[group_var]][!is.na(data[[group_var]])]
  if (length(valid_groups) == 0) {
    cat("FEHLER: Gruppenvariable", group_var, "hat keine gültigen Werte\n")
    return(NULL)
  }
  
  unique_groups <- unique(valid_groups)
  if (length(unique_groups) < 2) {
    cat("FEHLER: Gruppenvariable", group_var, "hat nur eine Gruppe:", unique_groups, "\n")
    return(NULL)
  }
  
  # 3. Prüfe vollständige Fälle
  complete_cases <- !is.na(data[[numeric_var]]) & !is.na(data[[group_var]])
  if (sum(complete_cases) < 5) {
    cat("FEHLER: Zu wenige vollständige Fälle (", sum(complete_cases), ") für Gruppenvergleich\n")
    return(NULL)
  }
  
  # 4. Prüfe ob jede Gruppe mindestens 2 Werte hat
  group_counts <- table(data[[group_var]][complete_cases])
  if (any(group_counts < 2)) {
    small_groups <- names(group_counts)[group_counts < 2]
    cat("WARNUNG: Folgende Gruppen haben weniger als 2 Werte:", paste(small_groups, collapse = ", "), "\n")
  }
  
  cat("Validierung erfolgreich:", length(unique_groups), "Gruppen mit", sum(complete_cases), "vollständigen Fällen\n")
  
  # NEUER FIX: Survey-Objekt mit konvertierten Daten neu erstellen falls nötig
  if (!is.null(survey_obj) && WEIGHTS) {
    cat("Erstelle Survey-Objekt mit konvertierten Daten...\n")
    
    # WICHTIG: Konvertiere zu data.frame vor Survey-Objekt-Erstellung
    data_for_survey <- as.data.frame(data)
    
    # WICHTIG: Survey-Objekt mit aktuellen (konvertierten) Daten erstellen
    survey_obj <- create_survey_object(data_for_survey, WEIGHT_VAR)
    
    # Gefilterte Survey-Objekt für vollständige Fälle
    survey_complete <- subset(survey_obj, !is.na(get(numeric_var)) & !is.na(get(group_var)))
    
    stats_list <- list()
    
    # Gruppierungen ermitteln
    groups <- unique_groups
    groups <- sort(groups)
    
    for (group in groups) {
      group_survey <- subset(survey_complete, get(group_var) == group)
      
      if (nrow(group_survey$variables) > 0) {
        # Sichere Extraktion von svyquantile - konvertiere zu Vektor
        median_result <- svyquantile(as.formula(paste("~", numeric_var)), group_survey, 0.5, na.rm = TRUE)
        q1_result <- svyquantile(as.formula(paste("~", numeric_var)), group_survey, 0.25, na.rm = TRUE)
        q3_result <- svyquantile(as.formula(paste("~", numeric_var)), group_survey, 0.75, na.rm = TRUE)
        min_result <- svyquantile(as.formula(paste("~", numeric_var)), group_survey, 0, na.rm = TRUE)
        max_result <- svyquantile(as.formula(paste("~", numeric_var)), group_survey, 1, na.rm = TRUE)
        
        stats_list[[as.character(group)]] <- list(
          n = nrow(group_survey$variables),
          mean = as.numeric(svymean(as.formula(paste("~", numeric_var)), group_survey, na.rm = TRUE)),
          median = as.numeric(as.vector(median_result[[1]])[1]),
          q1 = as.numeric(as.vector(q1_result[[1]])[1]),
          q3 = as.numeric(as.vector(q3_result[[1]])[1]),
          min = as.numeric(as.vector(min_result[[1]])[1]),
          max = as.numeric(as.vector(max_result[[1]])[1]),
          sd = as.numeric(sqrt(svyvar(as.formula(paste("~", numeric_var)), group_survey, na.rm = TRUE)))
        )
      }
    }
    
  } else {
    # Ungewichtete Gruppenmittelwerte
    stats_list <- list()
    groups <- sort(unique_groups)
    
    for (group in groups) {
      group_data <- data[data[[group_var]] == group & !is.na(data[[group_var]]), numeric_var]
      group_data <- group_data[!is.na(group_data)]
      
      if (length(group_data) > 0) {
        stats_list[[as.character(group)]] <- list(
          n = length(group_data),
          mean = mean(group_data, na.rm = TRUE),
          median = median(group_data, na.rm = TRUE),
          q1 = as.numeric(quantile(group_data, 0.25, na.rm = TRUE)),
          q3 = as.numeric(quantile(group_data, 0.75, na.rm = TRUE)),
          min = min(group_data, na.rm = TRUE),
          max = max(group_data, na.rm = TRUE),
          sd = sd(group_data, na.rm = TRUE)
        )
      }
    }
  }
  
  # Ergebnis-Tabelle erstellen
  group_stats_df <- data.frame(
    Gruppe = character(),
    N = numeric(),
    Mittelwert = numeric(),
    Median = numeric(),
    Q1 = numeric(),
    Q3 = numeric(),
    Min = numeric(),
    Max = numeric(),
    Standardabweichung = numeric(),
    stringsAsFactors = FALSE
  )
  
  for (group_name in names(stats_list)) {
    stats <- stats_list[[group_name]]
    group_stats_df <- rbind(group_stats_df, data.frame(
      Gruppe = group_name,
      N = stats$n,
      Mittelwert = round(stats$mean, DIGITS_ROUND),
      Median = round(stats$median, DIGITS_ROUND),
      Q1 = round(stats$q1, DIGITS_ROUND),
      Q3 = round(stats$q3, DIGITS_ROUND),
      Min = round(stats$min, DIGITS_ROUND),
      Max = round(stats$max, DIGITS_ROUND),
      Standardabweichung = round(stats$sd, DIGITS_ROUND),
      stringsAsFactors = FALSE
    ))
  }
  
  # Gesamtstatistiken hinzufügen
  if (!is.null(survey_obj) && WEIGHTS) {
    # Sichere Extraktion von svyquantile - konvertiere zu Vektor
    total_median <- svyquantile(as.formula(paste("~", numeric_var)), survey_complete, 0.5, na.rm = TRUE)
    total_q1 <- svyquantile(as.formula(paste("~", numeric_var)), survey_complete, 0.25, na.rm = TRUE)
    total_q3 <- svyquantile(as.formula(paste("~", numeric_var)), survey_complete, 0.75, na.rm = TRUE)
    total_min <- svyquantile(as.formula(paste("~", numeric_var)), survey_complete, 0, na.rm = TRUE)
    total_max <- svyquantile(as.formula(paste("~", numeric_var)), survey_complete, 1, na.rm = TRUE)
    
    total_stats <- list(
      n = nrow(survey_complete$variables),
      mean = as.numeric(svymean(as.formula(paste("~", numeric_var)), survey_complete, na.rm = TRUE)),
      median = as.numeric(as.vector(total_median[[1]])[1]),
      q1 = as.numeric(as.vector(total_q1[[1]])[1]),
      q3 = as.numeric(as.vector(total_q3[[1]])[1]),
      min = as.numeric(as.vector(total_min[[1]])[1]),
      max = as.numeric(as.vector(total_max[[1]])[1]),
      sd = as.numeric(sqrt(svyvar(as.formula(paste("~", numeric_var)), survey_complete, na.rm = TRUE)))
    )
  } else {
    all_values <- data[[numeric_var]][!is.na(data[[numeric_var]])]
    total_stats <- list(
      n = length(all_values),
      mean = mean(all_values, na.rm = TRUE),
      median = median(all_values, na.rm = TRUE),
      q1 = as.numeric(quantile(all_values, 0.25, na.rm = TRUE)),
      q3 = as.numeric(quantile(all_values, 0.75, na.rm = TRUE)),
      min = min(all_values, na.rm = TRUE),
      max = max(all_values, na.rm = TRUE),
      sd = sd(all_values, na.rm = TRUE)
    )
  }
  
  group_stats_df <- rbind(group_stats_df, data.frame(
    Gruppe = "Gesamt",
    N = total_stats$n,
    Mittelwert = round(total_stats$mean, DIGITS_ROUND),
    Median = round(total_stats$median, DIGITS_ROUND),
    Q1 = round(total_stats$q1, DIGITS_ROUND),
    Q3 = round(total_stats$q3, DIGITS_ROUND),
    Min = round(total_stats$min, DIGITS_ROUND),
    Max = round(total_stats$max, DIGITS_ROUND),
    Standardabweichung = round(total_stats$sd, DIGITS_ROUND),
    stringsAsFactors = FALSE
  ))
  
  return(list(
    group_means = group_stats_df,
    n_total = sum(group_stats_df$N[group_stats_df$Gruppe != "Gesamt"]),
    var1_name = numeric_var,
    var2_name = group_var,
    type = "group_means"
  ))
}

# Angepasste create_crosstabs Funktion 
create_crosstabs <- function(prepared_data) {
  cat("\nErstelle Kreuztabellen und statistische Tests...\n")
  
  data <- prepared_data$data
  config <- prepared_data$config
  
  # Prüfen ob Kreuztabellen konfiguriert sind
  if (nrow(config$kreuztabellen) == 0) {
    cat("Keine Kreuztabellen konfiguriert.\n")
    return(list())
  }
  
  results <- list()
  
  # Gewichtetes Survey-Objekt erstellen falls gewünscht
  survey_obj <- NULL
  if (WEIGHTS && WEIGHT_VAR %in% names(data)) {
    survey_obj <- create_survey_object(data, WEIGHT_VAR)
  }
  
  # Für jede konfigurierte Kreuztabelle
  for (i in 1:nrow(config$kreuztabellen)) {
    analysis_name <- config$kreuztabellen$analysis_name[i]
    var1 <- config$kreuztabellen$variable_1[i]
    var2 <- config$kreuztabellen$variable_2[i]
    test_type <- config$kreuztabellen$statistical_test[i]
    
    cat("💫 Verarbeite Kreuztabelle:", analysis_name, "(", var1, "x", var2, ")\n")
    
    # Filter für diese Kreuztabelle anwenden (falls vorhanden)
    filtered <- apply_row_filter(data, config$kreuztabellen[i,], names(data))
    current_data <- filtered$data
    filter_applied <- filtered$filtered
    filter_info <- filtered$filter_info
    
    # Gewichtetes Survey-Objekt für gefilterte Daten erstellen (falls nötig)
    current_survey_obj <- NULL
    if (WEIGHTS && WEIGHT_VAR %in% names(current_data)) {
      current_survey_obj <- create_survey_object(current_data, WEIGHT_VAR)
      if (filter_applied) {
        cat("  Filter angewendet: '", filter_info$filter_string, "' - ", 
            filter_info$filtered_n, " von ", filter_info$original_n, " Fällen (",
            round(filter_info$filtered_n/filter_info$original_n*100, 1), "%)\n", sep = "")
      }
    } else if (!filter_applied) {
      # Kein Filter angewendet: globales Survey-Objekt verwenden (falls vorhanden)
      current_survey_obj <- survey_obj
    }
    
    # MATRIX-ERKENNUNG VOR EXISTENZPRÜFUNG
    var1_is_matrix <- is_matrix_variable(var1, current_data, config)
    var2_is_matrix <- is_matrix_variable(var2, current_data, config)
    
    # Prüfen ob Variablen existieren (mit Matrix-Ausnahme)
    if (!var1_is_matrix && !var1 %in% names(current_data)) {
      cat("WARNUNG: Variable", var1, "nicht gefunden für", analysis_name, "\n")
      next
    }
    if (!var2_is_matrix && !var2 %in% names(current_data)) {
      cat("WARNUNG: Variable", var2, "nicht gefunden für", analysis_name, "\n")
      next
    }
    
    # Spezialfall: Beide normal aber eine nicht gefunden
    if (!var1_is_matrix && !var2_is_matrix && (!var1 %in% names(current_data) || !var2 %in% names(current_data))) {
      cat("WARNUNG: Variable(n) nicht gefunden für", analysis_name, "\n")
      next
    }
    
    # Debug-Ausgabe
    if (var1_is_matrix) {
      cat("→", var1, "als Matrix erkannt\n")
    }
    if (var2_is_matrix) {
      cat("→", var2, "als Matrix erkannt\n")
    }
    
    # Kreuztabelle erstellen (config ist jetzt optional)
    crosstab_result <- create_contingency_table(current_data, var1, var2, current_survey_obj, config)
    
    # Statistischen Test durchführen (für normale und Matrix-Kreuztabellen)
    test_result <- NULL
    if (!is.null(crosstab_result)) {
      if (!"matrix_items" %in% names(crosstab_result)) {
        # Normale Kreuztabelle
        test_result <- perform_statistical_test(current_data, var1, var2, test_type, current_survey_obj, config)
      } else {
        # Matrix-Kreuztabelle: Statistische Tests für jedes Matrix-Item
        matrix_items <- crosstab_result$matrix_items
        group_var <- if(var1_is_matrix) var2 else var1
        
        # Sammle Testergebnisse für alle Matrix-Items
        matrix_test_results <- list()
        
        for (item in matrix_items) {
          # Für jedes Matrix-Item den entsprechenden Test durchführen
          item_test_result <- perform_statistical_test(current_data, item, group_var, test_type, current_survey_obj, config)
          matrix_test_results[[item]] <- item_test_result
        }
        
        # Erstelle zusammengefasstes Testergebnis
        test_result <- list(
          test = paste("Matrix-Kreuztabelle -", test_type, "Test"),
          result = paste("Tests für", length(matrix_items), "Matrix-Items durchgeführt"),
          p_value = NA,
          statistic = NA,
          item_tests = matrix_test_results
        )
      }
    }
    
    # Ergebnisse kombinieren
    results[[analysis_name]] <- list(
      analysis_name = analysis_name,
      variable_1 = var1,
      variable_2 = var2,
      crosstab = crosstab_result,
      statistical_test = test_result,
      weighted = !is.null(current_survey_obj) && WEIGHTS,
      filter_applied = filter_applied,
      filter_string = if(filter_applied) filter_info$filter_string else NA_character_,
      original_n = if(filter_applied) filter_info$original_n else nrow(data),
      filtered_n = if(filter_applied) filter_info$filtered_n else nrow(current_data)
    )
  }
  
  cat("Kreuztabellen für", length(results), "Analysen erstellt.\n")
  return(results)
}

# Statistische Tests durchführen
perform_statistical_test <- function(data, var1, var2, test_type, survey_obj = NULL, config) {
  
  # Daten ohne fehlende Werte
  complete_data <- data[!is.na(data[[var1]]) & !is.na(data[[var2]]), ]
  
  if (nrow(complete_data) < 5) {
    return(list(
      test = test_type,
      result = "Zu wenige Daten für Test",
      p_value = NA,
      statistic = NA
    ))
  }
  
  # NEUE VALIDIERUNG: Factor-Operationen vermeiden
  if (!is.null(survey_obj) && WEIGHTS) {
    # Prüfe ob Variablen für Survey-Operationen geeignet sind
    if ((is.factor(complete_data[[var1]]) && test_type %in% c("correlation", "t_test", "anova")) ||
        (is.factor(complete_data[[var2]]) && test_type %in% c("correlation", "t_test", "anova"))) {
      
      # Temporär zu character konvertieren für Survey-Operationen
      temp_data <- as.data.frame(complete_data)
      if (is.factor(temp_data[[var1]]) && test_type != "chi_square") {
        temp_data[[var1]] <- as.character(temp_data[[var1]])
      }
      if (is.factor(temp_data[[var2]]) && test_type != "chi_square") {
        temp_data[[var2]] <- as.character(temp_data[[var2]])
      }
      
      # Survey-Objekt mit bereinigten Daten erstellen
      survey_obj <- create_survey_object(temp_data, WEIGHT_VAR)
    }
  }
  
  # Variable types bestimmen
  var1_config <- config$variablen[config$variablen$variable_name == var1, ]
  var2_config <- config$variablen[config$variablen$variable_name == var2, ]
  
  var1_type <- if(nrow(var1_config) > 0) var1_config$data_type else "unknown"
  var2_type <- if(nrow(var2_config) > 0) var2_config$data_type else "unknown"
  
  result <- tryCatch({
    switch(test_type,
           "chi_square" = perform_chi_square_test(complete_data, var1, var2, survey_obj),
           "t_test" = perform_t_test_safe(complete_data, var1, var2, var1_type, var2_type, survey_obj),
           "anova" = perform_anova_test_safe(complete_data, var1, var2, var1_type, var2_type, survey_obj),
           "correlation" = perform_correlation_test(complete_data, var1, var2, survey_obj),
           "mann_whitney" = perform_mann_whitney_test(complete_data, var1, var2, var1_type, var2_type),
           list(test = test_type, result = "Test nicht implementiert", p_value = NA, statistic = NA)
    )
  }, error = function(e) {
    list(
      test = test_type,
      result = paste("Fehler:", e$message),
      p_value = NA,
      statistic = NA
    )
  })
  
  return(result)
}
# Chi-Quadrat Test
perform_chi_square_test <- function(data, var1, var2, survey_obj = NULL) {
  if (!is.null(survey_obj) && WEIGHTS) {
    # Gewichteter Chi-Quadrat Test mit survey package
    survey_complete <- subset(survey_obj, !is.na(get(var1)) & !is.na(get(var2)))
    test_result <- svychisq(as.formula(paste("~", var1, "+", var2)), survey_complete)
    
    return(list(
      test = "Chi-Quadrat (gewichtet)",
      statistic = round(test_result$statistic, DIGITS_ROUND),
      p_value = round(test_result$p.value, 4),
      df = test_result$parameter,
      result = if(test_result$p.value < ALPHA_LEVEL) "Signifikant" else "Nicht signifikant"
    ))
  } else {
    # Standard Chi-Quadrat Test
    contingency_table <- table(data[[var1]], data[[var2]])
    test_result <- chisq.test(contingency_table)
    
    return(list(
      test = "Chi-Quadrat",
      statistic = round(test_result$statistic, DIGITS_ROUND),
      p_value = round(test_result$p.value, 4),
      df = test_result$parameter,
      result = if(test_result$p.value < ALPHA_LEVEL) "Signifikant" else "Nicht signifikant"
    ))
  }
}

# T-Test
perform_t_test_safe <- function(data, var1, var2, var1_type, var2_type, survey_obj = NULL) {
  # Bestimme welche Variable numerisch und welche kategorial ist
  var1_actual_type <- detect_actual_data_type(data, var1)
  var2_actual_type <- detect_actual_data_type(data, var2)
  
  if (var1_actual_type %in% c("numeric", "numeric_discrete") && 
      var2_actual_type %in% c("nominal", "nominal_binary")) {
    numeric_var <- var1
    group_var <- var2
  } else if (var2_actual_type %in% c("numeric", "numeric_discrete") && 
             var1_actual_type %in% c("nominal", "nominal_binary")) {
    numeric_var <- var2
    group_var <- var1
  } else {
    return(list(test = "t-Test", result = "Ungeeignete Variablentypen für t-Test", p_value = NA, statistic = NA))
  }
  
  # Sicherstellen dass numerische Variable wirklich numerisch ist
  if (!is.numeric(data[[numeric_var]])) {
    data[[numeric_var]] <- suppressWarnings(as.numeric(as.character(data[[numeric_var]])))
    
    if (all(is.na(data[[numeric_var]]))) {
      return(list(test = "t-Test", result = "Numerische Variable kann nicht konvertiert werden", p_value = NA, statistic = NA))
    }
  }
  
  # Prüfe ob Gruppenvariable genau 2 Gruppen hat
  groups <- unique(data[[group_var]][!is.na(data[[group_var]])])
  if (length(groups) != 2) {
    return(list(test = "t-Test", result = "Gruppenvariable muss genau 2 Ausprägungen haben", p_value = NA, statistic = NA))
  }
  
  # Rest wie vorher...
  if (!is.null(survey_obj) && WEIGHTS) {
    # Gewichteter T-Test mit bereinigten Daten
    survey_complete <- subset(survey_obj, !is.na(get(numeric_var)) & !is.na(get(group_var)))
    
    tryCatch({
      test_result <- svyttest(as.formula(paste(numeric_var, "~", group_var)), survey_complete)
      
      return(list(
        test = "t-Test (gewichtet)",
        statistic = round(test_result$statistic, DIGITS_ROUND),
        p_value = round(test_result$p.value, 4),
        df = round(test_result$parameter, 1),
        result = if(test_result$p.value < ALPHA_LEVEL) "Signifikant" else "Nicht signifikant"
      ))
    }, error = function(e) {
      return(list(test = "t-Test", result = paste("Gewichteter Test fehlgeschlagen:", e$message), p_value = NA, statistic = NA))
    })
  } else {
    # Standard t-Test
    group1_data <- data[data[[group_var]] == groups[1], numeric_var]
    group2_data <- data[data[[group_var]] == groups[2], numeric_var]
    
    group1_data <- group1_data[!is.na(group1_data)]
    group2_data <- group2_data[!is.na(group2_data)]
    
    if (length(group1_data) < 2 || length(group2_data) < 2) {
      return(list(test = "t-Test", result = "Zu wenige Daten in mindestens einer Gruppe", p_value = NA, statistic = NA))
    }
    
    test_result <- t.test(group1_data, group2_data)
    
    return(list(
      test = "t-Test",
      statistic = round(test_result$statistic, DIGITS_ROUND),
      p_value = round(test_result$p.value, 4),
      df = round(test_result$parameter, 1),
      result = if(test_result$p.value < ALPHA_LEVEL) "Signifikant" else "Nicht signifikant"
    ))
  }
}

# ANOVA Test
perform_anova_test_safe <- function(data, var1, var2, var1_type, var2_type, survey_obj = NULL) {
  # Verwende automatische Typ-Erkennung statt Config
  var1_actual_type <- detect_actual_data_type(data, var1)
  var2_actual_type <- detect_actual_data_type(data, var2)
  
  # Bestimme Variablen basierend auf tatsächlichen Typen
  if (var1_actual_type %in% c("numeric", "numeric_discrete") && 
      !var2_actual_type %in% c("numeric", "numeric_discrete")) {
    numeric_var <- var1
    group_var <- var2
  } else if (var2_actual_type %in% c("numeric", "numeric_discrete") && 
             !var1_actual_type %in% c("numeric", "numeric_discrete")) {
    numeric_var <- var2
    group_var <- var1
  } else {
    return(list(test = "ANOVA", result = "Ungeeignete Variablentypen für ANOVA", p_value = NA, statistic = NA))
  }
  
  # Sicherstellen dass numerische Variable wirklich numerisch ist
  if (!is.numeric(data[[numeric_var]])) {
    data[[numeric_var]] <- suppressWarnings(as.numeric(as.character(data[[numeric_var]])))
    
    if (all(is.na(data[[numeric_var]]))) {
      return(list(test = "ANOVA", result = "Numerische Variable kann nicht konvertiert werden", p_value = NA, statistic = NA))
    }
  }
  
  # Vollständige Fälle
  complete_data <- data[!is.na(data[[numeric_var]]) & !is.na(data[[group_var]]), ]
  
  if (nrow(complete_data) < 5) {
    return(list(test = "ANOVA", result = "Zu wenige Daten für ANOVA", p_value = NA, statistic = NA))
  }
  
  # UNGEWICHTETE ANOVA (vermeidet Survey-Factor-Probleme)
  tryCatch({
    formula_str <- paste(numeric_var, "~", group_var)
    
    # Gruppenvariable als Factor falls nötig, aber nur für ANOVA
    if (!is.factor(complete_data[[group_var]])) {
      complete_data[[group_var]] <- as.factor(as.character(complete_data[[group_var]]))
    }
    
    test_result <- aov(as.formula(formula_str), data = complete_data)
    summary_result <- summary(test_result)
    
    # Extrahiere Ergebnisse
    f_value <- NA
    p_value <- NA
    df_str <- "NA"
    
    if (length(summary_result) > 0 && is.list(summary_result[[1]])) {
      anova_table <- summary_result[[1]]
      
      if ("F value" %in% names(anova_table) && length(anova_table$`F value`) > 0) {
        f_value <- anova_table$`F value`[1]
      }
      
      if ("Pr(>F)" %in% names(anova_table) && length(anova_table$`Pr(>F)`) > 0) {
        p_value <- anova_table$`Pr(>F)`[1]
      }
      
      if ("Df" %in% names(anova_table) && length(anova_table$Df) >= 2) {
        df_str <- paste(anova_table$Df[1], anova_table$Df[2], sep = ", ")
      }
    }
    
    return(list(
      test = "ANOVA (ungewichtet)",
      statistic = if(!is.na(f_value)) round(f_value, DIGITS_ROUND) else NA,
      p_value = if(!is.na(p_value)) round(p_value, 4) else NA,
      df = df_str,
      result = if(!is.na(p_value) && p_value < ALPHA_LEVEL) "Signifikant" else "Nicht signifikant"
    ))
    
  }, error = function(e) {
    return(list(
      test = "ANOVA",
      result = paste("ANOVA fehlgeschlagen:", e$message),
      p_value = NA,
      statistic = NA
    ))
  })
}

# Korrelationstest
perform_correlation_test <- function(data, var1, var2, survey_obj = NULL) {
  
  # Automatische Erkennung der Variablentypen aus den Daten
  var1_is_numeric <- is.numeric(data[[var1]])
  var2_is_numeric <- is.numeric(data[[var2]])
  
  cat("Korrelationsanalyse:", var1, "numerisch:", var1_is_numeric, "|", var2, "numerisch:", var2_is_numeric, "\n")
  
  # Vollständige Fälle
  complete_data <- data[!is.na(data[[var1]]) & !is.na(data[[var2]]), ]
  
  if (nrow(complete_data) < 5) {
    return(list(
      test = "Korrelation", 
      result = "Zu wenige Daten für Korrelationsanalyse", 
      p_value = NA, 
      statistic = NA
    ))
  }
  
  # 1. BEIDE NUMERISCH → Pearson-Korrelation
  if (var1_is_numeric && var2_is_numeric) {
    return(perform_pearson_correlation(complete_data, var1, var2, survey_obj))
  }
  
  # 2. BEIDE NOMINAL → Cramér's V (basierend auf Chi²)
  if (!var1_is_numeric && !var2_is_numeric) {
    return(perform_cramers_v(complete_data, var1, var2, survey_obj))
  }
  
  # 3. EINE NUMERISCH, EINE NOMIN AL → Eta²(Korrelationsverhältnis)
  if ((var1_is_numeric && !var2_is_numeric) || (!var1_is_numeric && var2_is_numeric)) {
    numeric_var <- if(var1_is_numeric) var1 else var2
    nominal_var <- if(var1_is_numeric) var2 else var1
    return(perform_eta_squared(complete_data, numeric_var, nominal_var, survey_obj))
  }
  
  # Fallback
  return(list(
    test = "Korrelation", 
    result = "Ungeeignete Variablentypen", 
    p_value = NA, 
    statistic = NA
  ))
}

# 1. Pearson-Korrelation für numerische Variablen
perform_pearson_correlation <- function(data, var1, var2, survey_obj = NULL) {
  
  if (!is.null(survey_obj) && WEIGHTS) {
    # Gewichtete Korrelation - VERBESSERTE FEHLERBEHANDLUNG
    survey_complete <- subset(survey_obj, !is.na(get(var1)) & !is.na(get(var2)))
    
    tryCatch({
      # NEUE VALIDIERUNG: Prüfe Survey-Daten vor svycor
      if (nrow(survey_complete$variables) < 5) {
        stop("Zu wenige vollständige Fälle für gewichtete Korrelation")
      }
      
      # Prüfe ob beide Variablen numerisch sind
      var1_data <- survey_complete$variables[[var1]]
      var2_data <- survey_complete$variables[[var2]]
      
      if (!is.numeric(var1_data) || !is.numeric(var2_data)) {
        stop("Beide Variablen müssen numerisch für Korrelation sein")
      }
      
      # Prüfe auf Varianz
      if (var(var1_data, na.rm = TRUE) == 0 || var(var2_data, na.rm = TRUE) == 0) {
        stop("Eine Variable hat keine Varianz")
      }
      
      # Versuche svycor
      corr_result <- svycor(as.formula(paste("~", var1, "+", var2)), survey_complete)
      
      # P-Wert approximieren
      n <- nrow(survey_complete$variables)
      r <- corr_result[1,2]
      
      # Validiere Korrelationskoeffizient
      if (is.na(r) || !is.finite(r)) {
        stop("Ungültiger Korrelationskoeffizient")
      }
      
      t_stat <- r * sqrt((n-2)/(1-r^2))
      p_value <- 2 * (1 - pt(abs(t_stat), df = n-2))
      
      return(list(
        test = "Pearson-Korrelation (gewichtet)",
        statistic = round(r, DIGITS_ROUND),
        p_value = round(p_value, 4),
        result = if(p_value < ALPHA_LEVEL) "Signifikant" else "Nicht signifikant",
        interpretation = interpret_correlation(r)
      ))
    }, error = function(e) {
      cat("FALLBACK: Gewichtete Korrelation fehlgeschlagen:", e$message, "\n")
      cat("Verwende ungewichtete Korrelation als Fallback\n")
      
      # FALLBACK: Ungewichtete Korrelation
      complete_data <- data[!is.na(data[[var1]]) & !is.na(data[[var2]]), ]
      
      if (nrow(complete_data) < 5) {
        return(list(test = "Pearson-Korrelation", result = "Zu wenige Daten", p_value = NA, statistic = NA))
      }
      
      test_result <- cor.test(complete_data[[var1]], complete_data[[var2]])
      
      return(list(
        test = "Pearson-Korrelation (ungewichtet - Fallback)",
        statistic = round(test_result$estimate, DIGITS_ROUND),
        p_value = round(test_result$p.value, 4),
        result = if(test_result$p.value < ALPHA_LEVEL) "Signifikant" else "Nicht signifikant",
        interpretation = interpret_correlation(test_result$estimate)
      ))
    })
    
  } else {
    # Standard Korrelation
    test_result <- cor.test(data[[var1]], data[[var2]])
    
    return(list(
      test = "Pearson-Korrelation",
      statistic = round(test_result$estimate, DIGITS_ROUND),
      p_value = round(test_result$p.value, 4),
      result = if(test_result$p.value < ALPHA_LEVEL) "Signifikant" else "Nicht signifikant",
      interpretation = interpret_correlation(test_result$estimate)
    ))
  }
}


# 2. Cramér's V für nominale Variablen
perform_cramers_v <- function(data, var1, var2, survey_obj = NULL) {
  
  tryCatch({
    if (!is.null(survey_obj) && WEIGHTS) {
      # Gewichtetes Cramér's V
      survey_complete <- subset(survey_obj, !is.na(get(var1)) & !is.na(get(var2)))
      chi2_result <- svychisq(as.formula(paste("~", var1, "+", var2)), survey_complete)
      
      # Cramér's V berechnen
      chi2_stat <- chi2_result$statistic
      n <- sum(svytable(as.formula(paste("~", var1, "+", var2)), survey_complete))
      
    } else {
      # Ungewichtetes Cramér's V
      contingency_table <- table(data[[var1]], data[[var2]])
      chi2_result <- chisq.test(contingency_table)
      chi2_stat <- chi2_result$statistic
      n <- sum(contingency_table)
    }
    
    # Cramér's V = sqrt(Chi²/ (n * (min(r,c) - 1)))
    min_dim <- min(length(unique(data[[var1]])), length(unique(data[[var2]]))) - 1
    cramers_v <- sqrt(chi2_stat / (n * min_dim))
    
    return(list(
      test = "Cramér's V (nominaler Zusammenhang)",
      statistic = round(cramers_v, DIGITS_ROUND),
      p_value = round(chi2_result$p.value, 4),
      result = if(chi2_result$p.value < ALPHA_LEVEL) "Signifikant" else "Nicht signifikant",
      interpretation = interpret_cramers_v(cramers_v)
    ))
    
  }, error = function(e) {
    return(list(
      test = "Cramér's V", 
      result = paste("Fehler bei Cramér's V:", e$message), 
      p_value = NA, 
      statistic = NA
    ))
  })
}

# 3. Eta²für numerisch öÃ¢â‚¬â€ nominal
perform_eta_squared <- function(data, numeric_var, nominal_var, survey_obj = NULL) {
  
  tryCatch({
    if (!is.null(survey_obj) && WEIGHTS) {
      # Gewichtetes Eta²
      survey_complete <- subset(survey_obj, !is.na(get(numeric_var)) & !is.na(get(nominal_var)))
      
      # ANOVA für Eta²
      anova_model <- svyglm(as.formula(paste(numeric_var, "~", nominal_var)), survey_complete)
      anova_result <- anova(anova_model)
      
      # Eta²= SS_between / SS_total
      ss_between <- anova_result$`Sum Sq`[1]
      ss_total <- sum(anova_result$`Sum Sq`, na.rm = TRUE)
      eta_squared <- ss_between / ss_total
      
      p_value <- anova_result$`Pr(>F)`[1]
      
    } else {
      # Ungewichtetes Eta²
      # ANOVA durchführen
      anova_result <- aov(as.formula(paste(numeric_var, "~", nominal_var)), data = data)
      anova_summary <- summary(anova_result)
      
      # Eta²= SS_between / SS_total
      ss_between <- anova_summary[[1]]$`Sum Sq`[1]
      ss_total <- sum(anova_summary[[1]]$`Sum Sq`)
      eta_squared <- ss_between / ss_total
      
      p_value <- anova_summary[[1]]$`Pr(>F)`[1]
    }
    
    return(list(
      test = "Eta²(Korrelationsverhältnis)",
      statistic = round(eta_squared, DIGITS_ROUND),
      p_value = round(p_value, 4),
      result = if(p_value < ALPHA_LEVEL) "Signifikant" else "Nicht signifikant",
      interpretation = interpret_eta_squared(eta_squared)
    ))
    
  }, error = function(e) {
    return(list(
      test = "Eta²", 
      result = paste("Fehler bei Eta²:", e$message), 
      p_value = NA, 
      statistic = NA
    ))
  })
}

# Interpretationshilfen
interpret_correlation <- function(r) {
  r_abs <- abs(r)
  if (r_abs < 0.1) return("Sehr schwacher Zusammenhang")
  if (r_abs < 0.3) return("Schwacher Zusammenhang")
  if (r_abs < 0.5) return("Mittlerer Zusammenhang")
  if (r_abs < 0.7) return("Starker Zusammenhang")
  return("Sehr starker Zusammenhang")
}

interpret_cramers_v <- function(v) {
  if (v < 0.1) return("Sehr schwacher Zusammenhang")
  if (v < 0.2) return("Schwacher Zusammenhang")
  if (v < 0.4) return("Mittlerer Zusammenhang")
  if (v < 0.6) return("Starker Zusammenhang")
  return("Sehr starker Zusammenhang")
}

interpret_eta_squared <- function(eta2) {
  if (eta2 < 0.01) return("Sehr schwacher Effekt")
  if (eta2 < 0.06) return("Schwacher Effekt")
  if (eta2 < 0.14) return("Mittlerer Effekt")
  return("Starker Effekt")
}

# Mann-Whitney-U Test
perform_mann_whitney_test <- function(data, var1, var2, var1_type, var2_type) {
  # Bestimme welche Variable numerisch/ordinal und welche kategorial ist
  if (var1_type %in% c("numeric", "ordinal") && var2_type %in% c("nominal_coded", "nominal_text", "dichotom")) {
    numeric_var <- var1
    group_var <- var2
  } else if (var2_type %in% c("numeric", "ordinal") && var1_type %in% c("nominal_coded", "nominal_text", "dichotom")) {
    numeric_var <- var2
    group_var <- var1
  } else {
    return(list(test = "Mann-Whitney-U", result = "Ungeeignete Variablentypen", p_value = NA, statistic = NA))
  }
  
  # Prüfe ob Gruppenvariable genau 2 Gruppen hat
  groups <- unique(data[[group_var]][!is.na(data[[group_var]])])
  if (length(groups) != 2) {
    return(list(test = "Mann-Whitney-U", result = "Gruppenvariable muss genau 2 Ausprägungen haben", p_value = NA, statistic = NA))
  }
  
  group1_data <- data[data[[group_var]] == groups[1], numeric_var]
  group2_data <- data[data[[group_var]] == groups[2], numeric_var]
  
  # Entferne NA Werte
  group1_data <- group1_data[!is.na(group1_data)]
  group2_data <- group2_data[!is.na(group2_data)]
  
  test_result <- wilcox.test(group1_data, group2_data)
  
  return(list(
    test = "Mann-Whitney-U",
    statistic = round(test_result$statistic, DIGITS_ROUND),
    p_value = round(test_result$p.value, 4),
    result = if(test_result$p.value < ALPHA_LEVEL) "Signifikant" else "Nicht signifikant"
  ))
}

# =============================================================================
# REGRESSIONSANALYSEN
# =============================================================================

run_regressions <- function(prepared_data) {
  cat("\nFühre Regressionsanalysen durch...\n")
  
  data <- prepared_data$data
  config <- prepared_data$config
  
  # Prüfen ob Regressionen konfiguriert sind
  if (nrow(config$regressionen) == 0) {
    cat("Keine Regressionen konfiguriert.\n")
    return(list())
  }
  
  results <- list()
  
  # Gewichtetes Survey-Objekt erstellen falls gewünscht
  survey_obj <- NULL
  if (WEIGHTS && WEIGHT_VAR %in% names(data)) {
    survey_obj <- create_survey_object(data, WEIGHT_VAR)
  }
  
  # Für jede konfigurierte Regression
  for (i in 1:nrow(config$regressionen)) {
    regression_name <- config$regressionen$regression_name[i]
    dependent_var <- config$regressionen$dependent_var[i]
    independent_vars <- str_split(config$regressionen$independent_vars[i], ";")[[1]]
    regression_type <- config$regressionen$regression_type[i]
    
    cat("💫 Verarbeite Regression:", regression_name, "\n")
    cat("  AV:", dependent_var, "\n")
    cat("  UV:", paste(independent_vars, collapse = ", "), "\n")
    
    # Filter für diese Regression anwenden (falls vorhanden)
    filtered <- apply_row_filter(data, config$regressionen[i,], names(data))
    current_data <- filtered$data
    filter_applied <- filtered$filtered
    filter_info <- filtered$filter_info
    
    # Gewichtetes Survey-Objekt für gefilterte Daten erstellen (falls nötig)
    current_survey_obj <- NULL
    if (WEIGHTS && WEIGHT_VAR %in% names(current_data)) {
      current_survey_obj <- create_survey_object(current_data, WEIGHT_VAR)
      if (filter_applied) {
        cat("  Filter angewendet: '", filter_info$filter_string, "' - ", 
            filter_info$filtered_n, " von ", filter_info$original_n, " Fällen (",
            round(filter_info$filtered_n/filter_info$original_n*100, 1), "%)\n", sep = "")
      }
    } else if (!filter_applied) {
      # Kein Filter angewendet: globales Survey-Objekt verwenden (falls vorhanden)
      current_survey_obj <- survey_obj
    }
    
    # Prüfen ob alle Variablen existieren
    all_vars <- c(dependent_var, independent_vars)
    missing_vars <- all_vars[!all_vars %in% names(current_data)]
    
    if (length(missing_vars) > 0) {
      cat("WARNUNG: Variable(n) nicht gefunden:", paste(missing_vars, collapse = ", "), "\n")
      next
    }
    
    # Regression durchführen
    regression_result <- perform_regression(current_data, dependent_var, independent_vars, 
                                            regression_type, current_survey_obj, config, regression_name)
    
    if (!is.null(regression_result)) {
      # Filter-Info zum Ergebnis hinzufügen
      regression_result$filter_applied <- filter_applied
      regression_result$filter_string <- if(filter_applied) filter_info$filter_string else NA_character_
      regression_result$original_n <- if(filter_applied) filter_info$original_n else nrow(data)
      regression_result$filtered_n <- if(filter_applied) filter_info$filtered_n else nrow(current_data)
      
      results[[regression_name]] <- regression_result
    }
  }
  
  cat("Regressionsanalysen für", length(results), "Modelle erstellt.\n")
  return(results)
}

should_be_factor_for_regression <- function(data, var_name, config) {
  # Prüfe Config-Typ
  config_row <- config$variablen[config$variablen$variable_name == var_name, ]
  if (nrow(config_row) > 0) {
    var_type <- config_row$data_type[1]
    return(var_type %in% c("nominal_coded", "ordinal", "dichotom"))
  }
  
  # Fallback: Auto-Erkennung
  if (is.numeric(data[[var_name]])) {
    return(FALSE)
  } else {
    # Character/bereits Factor → sollte Factor sein
    return(TRUE)
  }
}

# Regression durchführen
# =============================================================================
# BUGFIX: perform_regression Funktion - processed_vars Definition verschieben
# =============================================================================

perform_regression <- function(data, dependent_var, independent_vars, regression_type, survey_obj, config, regression_name) {
  
  cat("Starte Regression:", regression_name, "\n")
  cat("AV:", dependent_var, "| UV:", paste(independent_vars, collapse = ", "), "\n")
  
  # 1. VARIABLE EXISTENZ PRÜFEN
  all_vars <- c(dependent_var, independent_vars)
  missing_vars <- all_vars[!all_vars %in% names(data)]
  
  if (length(missing_vars) > 0) {
    cat("FEHLER: Variablen nicht in Daten gefunden:", paste(missing_vars, collapse = ", "), "\n")
    return(NULL)
  }
  
  # 2. DATENTYP PRÜFUNG
  cat("Prüfe Datentypen:\n")
  for (var in all_vars) {
    var_class <- class(data[[var]])[1]
    n_valid <- sum(!is.na(data[[var]]))
    cat("  ", var, ":", var_class, "| Gültige Werte:", n_valid, "\n")
    
    if (n_valid == 0) {
      cat("FEHLER: Variable", var, "hat keine gültigen Werte\n")
      return(NULL)
    }
  }
  
  # 3. PROCESSED_VARS DEFINIEREN (VERSCHOBEN VOR DATENEXTRAKTION)
  processed_vars <- c()
  for (var_string in independent_vars) {
    var_string <- str_trim(var_string)
    # Interaktionsterme bleiben unverändert (mit *)
    processed_vars <- c(processed_vars, var_string)
  }
  
  # 4. VOLLSTÄNDIGE FÄLLE ERMITTELN (GEÄNDERT FÜR INTERAKTIONSTERME)
  tryCatch({
    # NEUE LOGIK: Extrahiere alle Variablen aus Interaktionstermen
    all_individual_vars <- c(dependent_var)
    
    for (var_string in independent_vars) {
      var_string <- str_trim(var_string)
      if (grepl("\\*", var_string)) {
        # Interaktionsterm: Extrahiere beide Variablen
        interaction_vars <- str_split(var_string, "\\*")[[1]]
        interaction_vars <- str_trim(interaction_vars)
        all_individual_vars <- c(all_individual_vars, interaction_vars)
        cat("  Interaktionsterm erkannt:", var_string, "→ Variablen:", paste(interaction_vars, collapse = ", "), "\n")
      } else {
        # Normale Variable
        all_individual_vars <- c(all_individual_vars, var_string)
      }
    }
    
    # Eindeutige Variablen für complete.cases
    unique_vars <- unique(all_individual_vars)
    cat("Alle Variablen für complete.cases:", paste(unique_vars, collapse = ", "), "\n")
    
    # Sichere Extraktion der Daten
    data_subset <- data[, unique_vars, drop = FALSE]
    complete_cases <- complete.cases(data_subset)
    complete_data <- data[complete_cases, ]  # Behalte alle ursprünglichen Spalten
    
    cat("Vollständige Fälle:", sum(complete_cases), "von", nrow(data), "\n")
    
    if (nrow(complete_data) < 10) {
      cat("WARNUNG: Zu wenige vollständige Fälle (", nrow(complete_data), ") für Regression\n")
      return(NULL)
    }
    
    # Just-in-Time Factor-Konvertierung für Regression
    for (var_string in processed_vars) {
      if (grepl("\\*", var_string)) {
        # Interaktionsterm: Beide Variablen prüfen
        interaction_vars <- str_split(var_string, "\\*")[[1]]
        interaction_vars <- str_trim(interaction_vars)
        
        for (var in interaction_vars) {
          if (var %in% names(complete_data) && should_be_factor_for_regression(complete_data, var, config)) {
            complete_data <- convert_to_factor_with_labels(complete_data, var)
          }
        }
      } else {
        # Normale Variable
        if (var_string %in% names(complete_data) && should_be_factor_for_regression(complete_data, var_string, config)) {
          complete_data <- convert_to_factor_with_labels(complete_data, var_string)
        }
      }
    }
    
  }, error = function(e) {
    cat("FEHLER bei Datenextraktion:", e$message, "\n")
    return(NULL)
  })
  
  # 5. MULTILEVEL CHECK (vor Formel-Erstellung)
  if (regression_type == "multilevel") {
    return(perform_multilevel_regression(complete_data, dependent_var, independent_vars, survey_obj, regression_name))
  }
  
  # 6. FORMEL ERSTELLEN UND VALIDIEREN (GEÄNDERT FÜR INTERAKTIONSTERME)
  formula_str <- paste(dependent_var, "~", paste(processed_vars, collapse = " + "))
  cat("Formel:", formula_str, "\n")
  
  formula_obj <- tryCatch({
    formula_obj <- as.formula(formula_str)
    
    # NEUE VALIDIERUNG: Datentypen für Regression prüfen
    cat("Validiere Variablentypen für Regression:\n")
    for (var_string in processed_vars) {
      if (grepl("\\*", var_string)) {
        # Interaktionsterm: Prüfe beide Variablen
        interaction_vars <- str_split(var_string, "\\*")[[1]]
        interaction_vars <- str_trim(interaction_vars)
        
        for (var in interaction_vars) {
          if (var %in% names(complete_data)) {
            var_class <- class(complete_data[[var]])[1]
            cat("  ", var, "(in Interaktion):", var_class)
            
            # Character zu Factor konvertieren (für kategoriale Variablen)
            if (var_class == "character") {
              complete_data[[var]] <- as.factor(complete_data[[var]])
              cat(" → konvertiert zu factor")
            }
            cat("\n")
          }
        }
      } else {
        # Normale Variable
        if (var_string %in% names(complete_data)) {
          var_class <- class(complete_data[[var_string]])[1]
          cat("  ", var_string, ":", var_class)
          
          # Character zu Factor konvertieren (für kategoriale Variablen)
          if (var_class == "character") {
            complete_data[[var_string]] <- as.factor(complete_data[[var_string]])
            cat(" → konvertiert zu factor")
          }
          cat("\n")
        }
      }
    }
    
    # Test ob Formel mit Daten funktioniert
    model_frame_test <- model.frame(formula_obj, data = complete_data, na.action = na.pass)
    cat("Formel-Test erfolgreich. Model frame Dimensionen:", dim(model_frame_test), "\n")
    
    formula_obj
  }, error = function(e) {
    cat("FEHLER bei Formel-Erstellung:", e$message, "\n")
    return(NULL)
  })
  
  if (is.null(formula_obj)) {
    return(NULL)
  }
  
  # 7. REGRESSION DURCHFÜHREN
  result <- tryCatch({
    switch(regression_type,
           "linear" = perform_linear_regression(complete_data, formula_obj, survey_obj),
           "logistic" = perform_logistic_regression(complete_data, formula_obj, survey_obj),
           "ordinal" = perform_ordinal_regression(complete_data, formula_obj, survey_obj),
           list(error = paste("Regressionstyp", regression_type, "nicht implementiert"))
    )
  }, error = function(e) {
    cat("DETAILLIERTER FEHLER bei Regression:\n")
    cat("  Typ:", regression_type, "\n")
    cat("  Formel:", formula_str, "\n")
    cat("  Daten-Dim:", dim(complete_data), "\n")
    cat("  Fehler:", e$message, "\n")
    
    list(error = paste("Fehler bei Regression:", e$message))
  })
  
  if ("error" %in% names(result)) {
    cat("FEHLER:", result$error, "\n")
    return(NULL)
  }
  
  # Metadaten hinzufügen
  result$regression_name <- regression_name
  result$dependent_var <- dependent_var
  result$independent_vars <- independent_vars
  result$regression_type <- regression_type
  result$n_complete <- nrow(complete_data)
  result$weighted <- !is.null(survey_obj) && WEIGHTS
  
  cat("✓ Regression", regression_name, "erfolgreich abgeschlossen\n")
  return(result)
}

# Lineare Regression
perform_linear_regression <- function(data, formula_obj, survey_obj = NULL) {
  
  if (!is.null(survey_obj) && WEIGHTS) {
    # Gewichtete lineare Regression
    survey_complete <- subset(survey_obj, complete.cases(survey_obj$variables[, all.vars(formula_obj)]))
    model <- svyglm(formula_obj, survey_complete, family = gaussian())
    
    # Modell-Zusammenfassung
    model_summary <- summary(model)
    
    # R-squared approximation für gewichtete Regression
    fitted_values <- fitted(model)
    observed_values <- survey_complete$variables[, all.vars(formula_obj)[1]]
    
    # Bessere R² Berechnung für gewichtete Regression
    ss_res <- sum((observed_values - fitted_values)^2)
    ss_tot <- sum((observed_values - mean(observed_values))^2)
    r_squared <- 1 - (ss_res / ss_tot)
    
    # Für gewichtete Regression: Pseudo-F-Test
    n <- nrow(survey_complete$variables)
    p <- length(all.vars(formula_obj)) - 1  # Anzahl Prädiktoren
    f_stat <- (r_squared / p) / ((1 - r_squared) / (n - p - 1))
    f_p_value <- pf(f_stat, p, n - p - 1, lower.tail = FALSE)
    
    # Modell-Güte für gewichtete Regression
    model_fit <- data.frame(
      Kennwert = c("R²", "Adjustiertes R²", "F-Statistik", "p-Wert (Modell)", "N"),
      Wert = c(
        round(r_squared, DIGITS_ROUND),
        round(1 - (1 - r_squared) * (n - 1) / (n - p - 1), DIGITS_ROUND),  # Adj. R²
        round(f_stat, DIGITS_ROUND),
        round(f_p_value, 4),
        n
      ),
      stringsAsFactors = FALSE
    )
    
  } else {
    # Standard lineare Regression
    model <- lm(formula_obj, data = data)
    model_summary <- summary(model)
    r_squared <- model_summary$r.squared
    
    # F-Statistik korrekt extrahieren
    f_stat <- model_summary$fstatistic
    if (!is.null(f_stat) && length(f_stat) >= 3) {
      f_value <- f_stat[1]
      f_p_value <- pf(f_stat[1], f_stat[2], f_stat[3], lower.tail = FALSE)
    } else {
      f_value <- NA
      f_p_value <- NA
    }
    
    # Modell-Güte für ungewichtete Regression
    model_fit <- data.frame(
      Kennwert = c("R²", "Adjustiertes R²", "F-Statistik", "p-Wert (Modell)", "N"),
      Wert = c(
        round(r_squared, DIGITS_ROUND),
        round(model_summary$adj.r.squared, DIGITS_ROUND),
        if(!is.na(f_value)) round(f_value, DIGITS_ROUND) else "NA",
        if(!is.na(f_p_value)) round(f_p_value, 4) else "NA",
        nrow(data)
      ),
      stringsAsFactors = FALSE
    )
  }
  
  # Koeffizienten-Tabelle erstellen (unverändert)
  coef_table <- data.frame(
    Variable = rownames(model_summary$coefficients),
    Koeffizient = sapply(model_summary$coefficients[, "Estimate"], smart_round_coefficient),
    Std_Fehler = round(model_summary$coefficients[, "Std. Error"], DIGITS_ROUND),
    t_Wert = round(model_summary$coefficients[, "t value"], DIGITS_ROUND),
    p_Wert = round(model_summary$coefficients[, "Pr(>|t|)"], 4),
    Signifikanz = ifelse(model_summary$coefficients[, "Pr(>|t|)"] < 0.001, "***",
                         ifelse(model_summary$coefficients[, "Pr(>|t|)"] < 0.01, "**",
                                ifelse(model_summary$coefficients[, "Pr(>|t|)"] < 0.05, "*",
                                       ifelse(model_summary$coefficients[, "Pr(>|t|)"] < 0.1, ".", "")))),
    stringsAsFactors = FALSE
  )
  
  return(list(
    model = model,
    coefficients = coef_table,
    model_fit = model_fit,
    type = "linear"
  ))
}

# Logistische Regression
perform_logistic_regression <- function(data, formula_obj, survey_obj = NULL) {
  
  if (!is.null(survey_obj) && WEIGHTS) {
    # Gewichtete logistische Regression
    survey_complete <- subset(survey_obj, complete.cases(survey_obj$variables[, all.vars(formula_obj)]))
    model <- svyglm(formula_obj, survey_complete, family = binomial())
    n <- nrow(survey_complete$variables)
  } else {
    # Standard logistische Regression
    model <- glm(formula_obj, data = data, family = binomial())
    n <- nrow(data)
  }
  
  model_summary <- summary(model)
  
  # Koeffizienten-Tabelle
  coef_table <- data.frame(
    Variable = rownames(model_summary$coefficients),
    Koeffizient = sapply(model_summary$coefficients[, "Estimate"], smart_round_coefficient),
    Std_Fehler = round(model_summary$coefficients[, "Std. Error"], DIGITS_ROUND),
    z_Wert = round(model_summary$coefficients[, "z value"], DIGITS_ROUND),
    p_Wert = round(model_summary$coefficients[, "Pr(>|z|)"], 4),
    Odds_Ratio = round(exp(model_summary$coefficients[, "Estimate"]), DIGITS_ROUND),
    Signifikanz = ifelse(model_summary$coefficients[, "Pr(>|z|)"] < 0.001, "***",
                         ifelse(model_summary$coefficients[, "Pr(>|z|)"] < 0.01, "**",
                                ifelse(model_summary$coefficients[, "Pr(>|z|)"] < 0.05, "*",
                                       ifelse(model_summary$coefficients[, "Pr(>|z|)"] < 0.1, ".", "")))),
    stringsAsFactors = FALSE
  )
  
  # Pseudo R² und weitere Statistiken
  null_deviance <- model$null.deviance
  residual_deviance <- model$deviance
  pseudo_r2_mcfadden <- 1 - (residual_deviance / null_deviance)
  
  # Cox & Snell R²
  pseudo_r2_cox_snell <- 1 - exp((residual_deviance - null_deviance) / n)
  
  # Nagelkerke R²
  pseudo_r2_nagelkerke <- pseudo_r2_cox_snell / (1 - exp(-null_deviance / n))
  
  # Chi²-Test für Modell
  chi2_stat <- null_deviance - residual_deviance
  df <- model$df.null - model$df.residual
  chi2_p_value <- pchisq(chi2_stat, df, lower.tail = FALSE)
  
  model_fit <- data.frame(
    Kennwert = c("Pseudo R² (McFadden)", "Pseudo R² (Cox & Snell)", "Pseudo R² (Nagelkerke)", 
                 "AIC", "Chi²-Statistik", "p-Wert (Modell)", "N"),
    Wert = c(
      round(pseudo_r2_mcfadden, DIGITS_ROUND),
      round(pseudo_r2_cox_snell, DIGITS_ROUND),
      round(pseudo_r2_nagelkerke, DIGITS_ROUND),
      round(AIC(model), DIGITS_ROUND),
      round(chi2_stat, DIGITS_ROUND),
      round(chi2_p_value, 4),
      n
    ),
    stringsAsFactors = FALSE
  )
  
  return(list(
    model = model,
    coefficients = coef_table,
    model_fit = model_fit,
    type = "logistic"
  ))
}

# Ordinale Regression (korrigiert)
perform_ordinal_regression <- function(data, formula_obj, survey_obj = NULL) {
  # Für ordinale Regression würde man normalerweise MASS::polr verwenden
  # Da das package nicht immer verfügbar ist, verwenden wir hier eine vereinfachte lineare Regression
  cat("HINWEIS: Ordinale Regression als lineare Regression durchgeführt.\n")
  
  # Verwende die korrigierte lineare Regression
  linear_result <- perform_linear_regression(data, formula_obj, survey_obj)
  
  # Ändere nur den Typ
  linear_result$type <- "ordinal (als linear)"
  
  return(linear_result)
}


# Ordinale Regression (vereinfacht)
perform_ordinal_regression <- function(data, formula_obj, survey_obj = NULL) {
  # Für ordinale Regression würde man normalerweise MASS::polr verwenden
  # Da das package nicht immer verfügbar ist, verwenden wir hier eine vereinfachte lineare Regression
  cat("HINWEIS: Ordinale Regression als lineare Regression durchgeführt.\n")
  
  # Verwende die korrigierte lineare Regression
  linear_result <- perform_linear_regression(data, formula_obj, survey_obj)
  
  # Ändere nur den Typ
  linear_result$type <- "ordinal (als linear)"
  
  return(linear_result)
}

perform_multilevel_regression <- function(data, dependent_var, independent_vars, survey_obj = NULL, regression_name) {
  
  cat("Führe Mehrebenenmodell durch:", regression_name, "\n")
  
  
  # AUTOMATISCHE CLUSTERING-VARIABLE ERKENNUNG
  cluster_var <- detect_cluster_variable(data, independent_vars)
  
  if (is.null(cluster_var)) {
    return(list(error = "Keine Clustering-Variable erkannt (z.B. Hochschul-ID)"))
  }
  
  cat("Clustering-Variable erkannt:", cluster_var, "\n")
  
  # LEVEL-1 UND LEVEL-2 VARIABLEN TRENNEN
  level_vars <- separate_multilevel_variables(data, independent_vars, cluster_var)
  
  # FORMEL ERSTELLEN
  formula_result <- create_multilevel_formula(dependent_var, level_vars$level1_vars, level_vars$level2_vars, cluster_var)
  
  if (is.null(formula_result$formula)) {
    return(list(error = paste("Fehler bei Formel-Erstellung:", formula_result$error)))
  }
  
  cat("Mehrebenen-Formel:", as.character(formula_result$formula), "\n")
  cat("Level-1 Variablen:", paste(level_vars$level1_vars, collapse = ", "), "\n")
  cat("Level-2 Variablen:", paste(level_vars$level2_vars, collapse = ", "), "\n")
  
  # MODELL SCHÄTZEN
  tryCatch({
    
    # Prüfe Variablentyp der AV
    if (is.factor(data[[dependent_var]]) || length(unique(data[[dependent_var]])) <= 2) {
      # Logistisches Mehrebenenmodell
      cat("Schätze logistisches Mehrebenenmodell...\n")
      model <- glmer(formula_result$formula, data = data, family = binomial())
      model_type <- "logistic_multilevel"
    } else {
      # Lineares Mehrebenenmodell
      cat("Schätze lineares Mehrebenenmodell...\n")
      model <- lmer(formula_result$formula, data = data)
      model_type <- "linear_multilevel"
    }
    
    # ERGEBNISSE EXTRAHIEREN
    return(extract_multilevel_results(model, model_type, level_vars, cluster_var, nrow(data)))
    
  }, error = function(e) {
    cat("FEHLER beim Schätzen des Mehrebenenmodells:", e$message, "\n")
    return(list(error = paste("Modell-Schätzung fehlgeschlagen:", e$message)))
  })
}

detect_cluster_variable <- function(data, independent_vars) {
  "Erkennt automatisch die Clustering-Variable (Hochschul-ID)"
  
  # Bekannte Hochschul-ID Patterns
  cluster_patterns <- c(
    "attribute_2", "hochschul_id", "hs_id", "uni_id", 
    "institution_id", "school_id", "cluster_id"
  )
  
  # 1. Direkte Suche nach bekannten Patterns
  for (pattern in cluster_patterns) {
    if (pattern %in% names(data)) {
      # Prüfe ob es wirklich eine Clustering-Variable ist (zwischen 5-100 Cluster)
      n_clusters <- length(unique(data[[pattern]][!is.na(data[[pattern]])]))
      if (n_clusters >= 5 && n_clusters <= 100) {
        cat("Clustering-Variable gefunden:", pattern, "(", n_clusters, "Cluster)\n")
        return(pattern)
      }
    }
  }
  
  # 2. Suche in independent_vars nach potentiellen Cluster-Variablen
  for (var in independent_vars) {
    if (var %in% names(data)) {
      n_unique <- length(unique(data[[var]][!is.na(data[[var]])]))
      total_n <- nrow(data)
      
      # Heuristik: 5-100 Gruppen, jede Gruppe hat mind. 5 Personen
      if (n_unique >= 5 && n_unique <= 100 && (total_n / n_unique) >= 5) {
        cat("Potentielle Clustering-Variable:", var, "(", n_unique, "Cluster)\n")
        return(var)
      }
    }
  }
  
  return(NULL)
}

separate_multilevel_variables <- function(data, independent_vars, cluster_var) {
  "Trennt Variablen in Level-1 (Individual) und Level-2 (Cluster)"
  
  level1_vars <- c()
  level2_vars <- c()
  
  for (var in independent_vars) {
    if (var == cluster_var) {
      next  # Cluster-Variable überspringen
    }
    
    if (!var %in% names(data)) {
      next  # Variable nicht in Daten
    }
    
    # Prüfe Varianz innerhalb von Clustern
    within_cluster_variance <- check_within_cluster_variance(data, var, cluster_var)
    
    if (within_cluster_variance > 0.1) {
      # Variable variiert innerhalb Cluster → Level-1
      level1_vars <- c(level1_vars, var)
      cat("Level-1:", var, "(variiert innerhalb Cluster)\n")
    } else {
      # Variable konstant innerhalb Cluster → Level-2
      level2_vars <- c(level2_vars, var)
      cat("Level-2:", var, "(konstant innerhalb Cluster)\n")
    }
  }
  
  return(list(
    level1_vars = level1_vars,
    level2_vars = level2_vars
  ))
}

check_within_cluster_variance <- function(data, var, cluster_var) {
  "Prüft ob Variable innerhalb Cluster variiert"
  
  # Für numerische Variablen: Varianz innerhalb Cluster
  if (is.numeric(data[[var]])) {
    cluster_variances <- aggregate(data[[var]], 
                                   by = list(data[[cluster_var]]), 
                                   FUN = function(x) var(x, na.rm = TRUE))
    
    mean_within_var <- mean(cluster_variances$x, na.rm = TRUE)
    total_var <- var(data[[var]], na.rm = TRUE)
    
    return(mean_within_var / total_var)
  }
  
  # Für kategoriale Variablen: Anzahl verschiedener Kategorien pro Cluster
  cluster_categories <- aggregate(data[[var]], 
                                  by = list(data[[cluster_var]]), 
                                  FUN = function(x) length(unique(x[!is.na(x)])))
  
  mean_categories <- mean(cluster_categories$x, na.rm = TRUE)
  total_categories <- length(unique(data[[var]][!is.na(data[[var]])]))
  
  return(mean_categories / total_categories)
}

create_multilevel_formula <- function(dependent_var, level1_vars, level2_vars, cluster_var) {
  "Erstellt Mehrebenen-Formel mit Random Intercept und optional Random Slopes"
  
  if (length(level1_vars) == 0 && length(level2_vars) == 0) {
    return(list(formula = NULL, error = "Keine Prädiktoren verfügbar"))
  }
  
  # Fixed Effects
  fixed_effects <- c(level1_vars, level2_vars)
  
  # Basis-Formel mit Random Intercept
  if (length(fixed_effects) > 0) {
    fixed_part <- paste(fixed_effects, collapse = " + ")
    formula_str <- paste(dependent_var, "~", fixed_part, "+ (1 |", cluster_var, ")")
  } else {
    # Nur Random Intercept
    formula_str <- paste(dependent_var, "~ 1 + (1 |", cluster_var, ")")
  }
  
  # Optional: Random Slope für erste Level-1 Variable falls vorhanden
  if (length(level1_vars) >= 1) {
    # Erweiterte Formel mit Random Slope (nur für erste Variable um Konvergenz zu verbessern)
    formula_str_extended <- paste(dependent_var, "~", fixed_part, "+ (1 +", level1_vars[1], "|", cluster_var, ")")
    
    return(list(
      formula = as.formula(formula_str_extended),
      formula_simple = as.formula(formula_str),
      fixed_effects = fixed_effects,
      random_slope_var = level1_vars[1]
    ))
  }
  
  return(list(
    formula = as.formula(formula_str),
    formula_simple = as.formula(formula_str),
    fixed_effects = fixed_effects,
    random_slope_var = NULL
  ))
}

extract_multilevel_results <- function(model, model_type, level_vars, cluster_var, n_total) {
  "Extrahiert Ergebnisse aus Mehrebenenmodell"
  
  # Model Summary
  model_summary <- summary(model)
  
  # FIXED EFFECTS TABELLE
  fixed_coef <- fixef(model)
  fixed_se <- sqrt(diag(vcov(model)))
  
  if (model_type == "linear_multilevel") {
    # T-Tests für lineare Modelle
    t_values <- fixed_coef / fixed_se
    df_est <- nrow(model@frame) - length(fixed_coef)  # Approximation
    p_values <- 2 * (1 - pt(abs(t_values), df = df_est))
    
    coefficients_table <- data.frame(
      Variable = names(fixed_coef),
      Koeffizient = sapply(fixed_coef, smart_round_coefficient),
      Std_Fehler = round(fixed_se, DIGITS_ROUND),
      t_Wert = round(t_values, DIGITS_ROUND),
      p_Wert = round(p_values, 4),
      Signifikanz = ifelse(p_values < 0.001, "***",
                           ifelse(p_values < 0.01, "**",
                                  ifelse(p_values < 0.05, "*",
                                         ifelse(p_values < 0.1, ".", "")))),
      stringsAsFactors = FALSE
    )
  } else {
    # Z-Tests für logistische Modelle
    z_values <- fixed_coef / fixed_se
    p_values <- 2 * (1 - pnorm(abs(z_values)))
    
    coefficients_table <- data.frame(
      Variable = names(fixed_coef),
      Koeffizient = sapply(fixed_coef, smart_round_coefficient),
      Std_Fehler = round(fixed_se, DIGITS_ROUND),
      z_Wert = round(z_values, DIGITS_ROUND),
      p_Wert = round(p_values, 4),
      Odds_Ratio = if(model_type == "logistic_multilevel") round(exp(fixed_coef), DIGITS_ROUND) else NA,
      Signifikanz = ifelse(p_values < 0.001, "***",
                           ifelse(p_values < 0.01, "**",
                                  ifelse(p_values < 0.05, "*",
                                         ifelse(p_values < 0.1, ".", "")))),
      stringsAsFactors = FALSE
    )
  }
  
  # RANDOM EFFECTS TABELLE
  random_effects <- as.data.frame(VarCorr(model))
  
  random_table <- data.frame(
    Komponente = paste(random_effects$grp, random_effects$var1, sep = " - "),
    Varianz = round(random_effects$vcov, DIGITS_ROUND),
    Std_Abweichung = round(random_effects$sdcor, DIGITS_ROUND),
    stringsAsFactors = FALSE
  )
  
  # MODEL FIT STATISTIKEN
  n_clusters <- length(unique(model@frame[[cluster_var]]))
  avg_cluster_size <- round(n_total / n_clusters, 1)
  
  # ICC berechnen
  var_components <- as.data.frame(VarCorr(model))
  if (model_type == "linear_multilevel") {
    between_var <- var_components$vcov[var_components$grp == cluster_var][1]
    within_var <- attr(VarCorr(model), "sc")^2  # Residual variance
    icc <- between_var / (between_var + within_var)
    
    model_fit <- data.frame(
      Kennwert = c("AIC", "BIC", "Log-Likelihood", "ICC", "Anzahl Cluster", "mittlere Cluster-Größe", "N"),
      Wert = c(
        round(AIC(model), 1),
        round(BIC(model), 1),
        round(as.numeric(logLik(model)), 1),
        round(icc, 3),
        n_clusters,
        avg_cluster_size,
        n_total
      ),
      stringsAsFactors = FALSE
    )
  } else {
    # Für logistische Modelle
    model_fit <- data.frame(
      Kennwert = c("AIC", "BIC", "Log-Likelihood", "Anzahl Cluster", "mittlere Cluster-Größe", "N"),
      Wert = c(
        round(AIC(model), 1),
        round(BIC(model), 1),
        round(as.numeric(logLik(model)), 1),
        n_clusters,
        avg_cluster_size,
        n_total
      ),
      stringsAsFactors = FALSE
    )
  }
  
  return(list(
    model = model,
    coefficients = coefficients_table,
    random_effects = random_table,
    model_fit = model_fit,
    type = model_type,
    cluster_variable = cluster_var,
    level1_variables = level_vars$level1_vars,
    level2_variables = level_vars$level2_vars,
    n_clusters = n_clusters
  ))
}


smart_round_coefficient <- function(x, digits = 2) {
  "Intelligente Rundung für Koeffizienten - zeigt mehr Stellen für sehr kleine Werte"
  
  if (is.na(x) || x == 0) return(x)
  
  abs_x <- abs(x)
  
  if (abs_x >= 0.01) {
    # Normale Werte: Standard-Rundung
    return(round(x, digits))
  } else if (abs_x >= 0.001) {
    # Kleine Werte: 3 Dezimalstellen
    return(round(x, 3))
  } else if (abs_x >= 0.0001) {
    # Sehr kleine Werte: 4 Dezimalstellen  
    return(round(x, 4))
  } else {
    # Extrem kleine Werte: Wissenschaftliche Notation
    return(formatC(x, format = "e", digits = 2))
  }
}


# =============================================================================
# TEXTANTWORTEN VERARBEITUNG
# =============================================================================


# DEBUG-FUNKTION: Finde alle GP05 und AS03 verwandten Variablen
debug_missing_variables <- function(data) {
  cat("=== DEBUG: Variablensuche ===\n")
  
  all_vars <- names(data)
  
  # 1. Suche alle GP05 ähnlichen Variablen
  gp05_vars <- all_vars[grepl("GP05", all_vars, ignore.case = TRUE)]
  cat("\nAlle GP05 verwandten Variablen:\n")
  if (length(gp05_vars) > 0) {
    for (var in gp05_vars) {
      cat("  -", var, "\n")
    }
  } else {
    cat("  Keine GP05 Variablen gefunden!\n")
  }
  
  # 2. Suche alle GP Variablen überhaupt
  gp_vars <- all_vars[grepl("^GP", all_vars, ignore.case = TRUE)]
  cat("\nAlle GP Variablen (erste 20):\n")
  for (var in head(gp_vars, 20)) {
    cat("  -", var, "\n")
  }
  
  # 3. Suche alle AS03 ähnlichen Variablen
  as03_vars <- all_vars[grepl("AS03", all_vars, ignore.case = TRUE)]
  cat("\nAlle AS03 verwandten Variablen:\n")
  if (length(as03_vars) > 0) {
    for (var in as03_vars) {
      cat("  -", var, "\n")
    }
  } else {
    cat("  Keine AS03 Variablen gefunden!\n")
  }
  
  # 4. Suche alle "other" Variablen
  other_vars <- all_vars[grepl("other", all_vars, ignore.case = TRUE)]
  cat("\nAlle 'other' verwandten Variablen:\n")
  if (length(other_vars) > 0) {
    for (var in head(other_vars, 20)) {
      cat("  -", var, "\n")
    }
  } else {
    cat("  Keine 'other' Variablen gefunden!\n")
  }
  
  # 5. Prüfe konkrete Variablennamen
  test_vars <- c("GP05", "GP03", "AS03", "AS03.other.", "AS03_other", "AS03[other]", "GP05_text", "GP05.text.")
  cat("\nDirekte Tests für spezifische Variablennamen:\n")
  for (test_var in test_vars) {
    exists <- test_var %in% all_vars
    cat("  -", test_var, ":", if(exists) "✓ EXISTIERT" else "✗ FEHLT", "\n")
  }
  
  # 6. Zeige Gesamtanzahl und Beispiele
  cat("\nGesamtanzahl Variablen:", length(all_vars), "\n")
  cat("Erste 30 Variablennamen:\n")
  for (var in head(all_vars, 30)) {
    cat("  -", var, "\n")
  }
}



# HILFSFUNKTION: Einzelvariable mit bewährter Logik finden
find_single_variable <- function(target_var, data_vars) {
  # Verwende die bewährte update_variable_list Logik für eine einzelne Variable
  result <- update_variable_list(c(target_var), data_vars)
  return(result[1])
}


process_text_responses <- function(prepared_data, custom_val_labels = NULL) {
  cat("\n💫 Verarbeite offene Textantworten...\n")
  
  data <- prepared_data$data
  config <- prepared_data$config
  
  if (nrow(config$textantworten) == 0) {
    cat("Keine Textantworten konfiguriert.\n")
    return(list())
  }
  
  results <- list()
  
  for (i in 1:nrow(config$textantworten)) {
    analysis_name <- config$textantworten$analysis_name[i]
    # VERWENDE BEREITS AKTUALISIERTE CONFIG-NAMEN (nicht die originalen!)
    text_var <- config$textantworten$text_variable[i]  
    sort_var <- config$textantworten$sort_variable[i]
    min_length <- config$textantworten$min_length[i]
    include_empty <- config$textantworten$include_empty[i]
    
    cat("\n--- Verarbeite:", analysis_name, "---\n")
    
    # Filter für diese Textanalyse anwenden (falls vorhanden)
    filtered <- apply_row_filter(data, config$textantworten[i,], names(data))
    current_data <- filtered$data
    filter_applied <- filtered$filtered
    filter_info <- filtered$filter_info
    
    if (filter_applied) {
      cat("  Filter angewendet: '", filter_info$filter_string, "' - ", 
          filter_info$filtered_n, " von ", filter_info$original_n, " Fällen (",
          round(filter_info$filtered_n/filter_info$original_n*100, 1), "%)\n", sep = "")
    }
    
    cat("Suche Text-Variable:", text_var, "\n")
    
    # DIREKTE PRÜFUNG - Config ist bereits aktualisiert!
    if (!text_var %in% names(current_data)) {
      cat("ÜBERSPRINGE:", analysis_name, "- Text-Variable", text_var, "nicht in Daten gefunden\n")
      cat("Verfügbare ähnliche Variablen:", paste(names(current_data)[grepl(text_var, names(current_data))], collapse = ", "), "\n")
      next
    }
    
    # Sort-Variable prüfen
    if (!is.na(sort_var) && sort_var != "" && !sort_var %in% names(current_data)) {
      cat("WARNUNG: Sort-Variable", sort_var, "nicht gefunden, verwende ohne Sortierung\n")
      sort_var <- NA
    }
    
    cat("✓ Verwende Text-Variable:", text_var, "| Sort-Variable:", sort_var, "\n")
    
    # Extrahiere Textantworten (unverändert)
    text_result <- extract_text_responses_simple(current_data, text_var, sort_var, min_length, include_empty)
    
    if (!is.null(text_result)) {
      # Filter-Info zum Ergebnis hinzufügen
      results[[analysis_name]] <- list(
        analysis_name = analysis_name,
        text_variable = text_var,
        sort_variable = sort_var,
        min_length = min_length,
        include_empty = include_empty,
        responses = text_result$responses,
        summary = text_result$summary,
        filter_applied = filter_applied,
        filter_string = if(filter_applied) filter_info$filter_string else NA_character_,
        original_n = if(filter_applied) filter_info$original_n else nrow(data),
        filtered_n = if(filter_applied) filter_info$filtered_n else nrow(current_data)
      )
      cat("✓ Analyse", analysis_name, "erfolgreich abgeschlossen\n")
    }
  }
  
  cat("\nTextantworten für", length(results), "Analysen verarbeitet.\n")
  return(results)
}

# Textantworten extrahieren
extract_text_responses_simple <- function(data, text_var, sort_var, min_length, include_empty, custom_val_labels = NULL) {
  
  cat("  Verwende Text-Variable:", text_var, "\n")
  
  # Basis-Daten vorbereiten (Variable ist bereits gefunden)
  text_data <- data.frame(
    ID = 1:nrow(data),
    Text = as.character(data[[text_var]]),
    stringsAsFactors = FALSE
  )
  
  # Sort-Variable hinzufügen falls vorhanden
  if (!is.na(sort_var) && sort_var != "" && sort_var %in% names(data)) {
    text_data$Sort_Kategorie <- as.character(data[[sort_var]])
    
    # Labels für Sort-Variable mit Priorisierung: RDS -> custom_val_labels -> Code
    labels <- get_value_labels_with_priority(data, sort_var, NULL)
    
    # Falls keine RDS-Labels, versuche custom_val_labels
    if ((is.null(labels) || length(labels) == 0) && !is.null(custom_val_labels) && sort_var %in% names(custom_val_labels)) {
      labels <- custom_val_labels[[sort_var]]
    }
    
    if (!is.null(labels) && length(labels) > 0) {
      text_data$Sort_Kategorie_Label <- NA_character_
      
      # Für jeden Code das passende Label finden (wie in create_nominal_coded_table)
      for (i in seq_len(nrow(text_data))) {
        code <- as.character(text_data$Sort_Kategorie[i])
        text_data$Sort_Kategorie_Label[i] <- code  # Default: Verwende Code als Label
        
        # Direkte Übereinstimmung: "1" -> "1"
        if (code %in% names(labels)) {
          text_data$Sort_Kategorie_Label[i] <- labels[code]
          next
        }
        
        # Pattern: AO01, AO02, AO03 -> extrahiere Nummer und versuche Match
        if (grepl("^[A-Z]+0*[0-9]+$", code)) {
          # Extrahiere Nummer: AO01 -> 1, AO02 -> 2, A001 -> 1
          num_part <- gsub("^[A-Z]+0*", "", code)
          
          # Versuche verschiedene Formate
          candidates <- c(
            num_part,                           # "1"
            paste0("AO", num_part),            # "AO1"
            paste0("AO0", num_part),           # "AO01"
            paste0("A", num_part),             # "A1"
            sprintf("%02d", as.numeric(num_part))  # "01"
          )
          
          for (candidate in candidates) {
            if (candidate %in% names(labels)) {
              text_data$Sort_Kategorie_Label[i] <- labels[candidate]
              break
            }
          }
        }
      }
    } else {
      text_data$Sort_Kategorie_Label <- text_data$Sort_Kategorie
    }
  } else {
    text_data$Sort_Kategorie <- "Alle"
    text_data$Sort_Kategorie_Label <- "Alle"
  }
  
  # Text bereinigen und filtern
  text_data$Text_bereinigt <- str_trim(text_data$Text)
  text_data$Text_Laenge <- nchar(text_data$Text_bereinigt)
  
  # Filtern nach Mindestlänge
  if (!include_empty) {
    text_data <- text_data[!is.na(text_data$Text_bereinigt) & 
                             text_data$Text_bereinigt != "" & 
                             text_data$Text_Laenge >= min_length, ]
  }
  
  if (nrow(text_data) == 0) {
    cat("  Keine Textantworten nach Filterung verfügbar\n")
    return(NULL)
  }
  
  # Antworten nach Sort-Kategorie sortieren
  text_data <- text_data[order(text_data$Sort_Kategorie_Label, -text_data$Text_Laenge), ]
  
  # Zusammenfassung erstellen
  summary_data <- text_data %>%
    group_by(Sort_Kategorie_Label) %>%
    summarise(
      Anzahl_Antworten = n(),
      Durchschnittliche_Laenge = round(mean(Text_Laenge, na.rm = TRUE), 1),
      Min_Laenge = min(Text_Laenge, na.rm = TRUE),
      Max_Laenge = max(Text_Laenge, na.rm = TRUE),
      .groups = 'drop'
    ) %>%
    arrange(desc(Anzahl_Antworten))
  
  # Antworten-Tabelle für Export vorbereiten
  responses_table <- text_data %>%
    select(Sort_Kategorie_Label, Text_bereinigt, Text_Laenge) %>%
    rename(
      Kategorie = Sort_Kategorie_Label,
      Textantwort = Text_bereinigt,
      Zeichen = Text_Laenge
    )
  
  cat("  Erfolgreich", nrow(responses_table), "Textantworten extrahiert\n")
  
  return(list(
    responses = responses_table,
    summary = summary_data,
    total_responses = nrow(text_data)
  ))
}


# =============================================================================
# VARIABLEN-ÜBERSICHT EXPORT ERGÄNZUNG
# =============================================================================

# Neue Funktion: Variablen-Übersicht erstellen
create_variable_overview <- function(data, config, descriptive_results, crosstab_results, regression_results, text_results = NULL, custom_var_labels = NULL) {
  cat("Erstelle Variablen-Übersicht...\n")
  
  # Alle Variablen im Datensatz sammeln
  all_vars <- names(data)
  
  # Config-Variablen extrahieren
  config_vars <- if(nrow(config$variablen) > 0) config$variablen$variable_name else character(0)
  
  # Genutzte Variablen sammeln
  used_in_descriptive <- names(descriptive_results)
  
  used_in_crosstabs <- character(0)
  if(length(crosstab_results) > 0) {
    for(result in crosstab_results) {
      used_in_crosstabs <- c(used_in_crosstabs, result$variable_1, result$variable_2)
    }
    used_in_crosstabs <- unique(used_in_crosstabs)
  }
  
  used_in_regressions <- character(0)
  if(length(regression_results) > 0) {
    for(result in regression_results) {
      # Abhängige Variable
      used_in_regressions <- c(used_in_regressions, result$dependent_var)
      # Unabhängige Variablen (mit Interaktionstermen)
      for(var_string in result$independent_vars) {
        if(grepl("\\*", var_string)) {
          # Interaktionsterm: Beide Variablen extrahieren
          interaction_vars <- str_split(var_string, "\\*")[[1]]
          interaction_vars <- str_trim(interaction_vars)
          used_in_regressions <- c(used_in_regressions, interaction_vars)
        } else {
          used_in_regressions <- c(used_in_regressions, str_trim(var_string))
        }
      }
    }
    used_in_regressions <- unique(used_in_regressions)
  }
  
  # Matrix-Variablen aus Config sammeln
  matrix_vars_used <- character(0)
  matrix_config <- config$variablen[config$variablen$data_type == "matrix", ]
  if(nrow(matrix_config) > 0) {
    for(i in 1:nrow(matrix_config)) {
      matrix_name <- matrix_config$variable_name[i]
      # Finde alle Matrix-Items mit verschiedenen Trennern (gleiche Logik wie in create_matrix_table)
      matrix_patterns <- c(
        paste0("^", matrix_name, "\\[.+\\]$"),     # Original: ZS01[001]
        paste0("^", matrix_name, "\\..+\\.$"),     # Sanitized: ZS01.001.
        paste0("^", matrix_name, "_.+$"),          # Underscore: ZS01_001
        paste0("^", matrix_name, "-.+$")           # Dash: ZS01-001
      )
      
      found_matrix_vars <- c()
      for (pattern in matrix_patterns) {
        found_vars <- names(data)[grepl(pattern, names(data))]
        found_matrix_vars <- c(found_matrix_vars, found_vars)
      }
      
      # FILTER OUT [other] variables
      found_matrix_vars <- found_matrix_vars[!grepl("other", found_matrix_vars, ignore.case = TRUE)]
      found_matrix_vars <- unique(found_matrix_vars)
      
      if(length(found_matrix_vars) > 0) {
        matrix_vars_used <- c(matrix_vars_used, found_matrix_vars)
        # Matrix-Variable selbst auch als verwendet markieren
        if(matrix_name %in% used_in_descriptive) {
          matrix_vars_used <- c(matrix_vars_used, matrix_name)
        }
      }
    }
    matrix_vars_used <- unique(matrix_vars_used)
  }
  
  # Textantworten-Variablen sammeln
  used_in_textantworten <- character(0)
  if(!is.null(text_results) && length(text_results) > 0) {
    for(result in text_results) {
      used_in_textantworten <- c(used_in_textantworten, result$text_variable)
      if(!is.na(result$sort_variable) && result$sort_variable != "") {
        used_in_textantworten <- c(used_in_textantworten, result$sort_variable)
      }
    }
    used_in_textantworten <- unique(used_in_textantworten)
  }
  
  # Übersichtstabelle erstellen
  overview <- data.frame(
    Variable = all_vars,
    Variable_Label = sapply(all_vars, function(var) {
      # Erst Custom Labels prüfen
      if(!is.null(custom_var_labels) && var %in% names(custom_var_labels)) {
        return(custom_var_labels[[var]])
      }
      # Dann Attribut-Labels aus Daten
      var_label <- attr(data[[var]], "label")
      if(!is.null(var_label) && var_label != "" && var_label != var) {
        return(var_label)
      }
      # Labelled-Package Labels
      if(requireNamespace("labelled", quietly = TRUE) && labelled::is.labelled(data[[var]])) {
        labelled_label <- labelled::var_label(data[[var]])
        if(!is.null(labelled_label) && labelled_label != "") {
          return(labelled_label)
        }
      }
      return("")  # Kein Label gefunden
    }),
    stringsAsFactors = FALSE
  )
  
  # Matrix-Info hinzufügen
  overview$Matrix_Info <- sapply(all_vars, function(var) {
    # Prüfe ob Variable eine Matrix-Hauptvariable ist
    matrix_config_row <- config$variablen[config$variablen$variable_name == var & config$variablen$data_type == "matrix", ]
    if(nrow(matrix_config_row) > 0) {
      return("Matrix-Hauptvariable")
    }
    
    # Prüfe ob Variable ein Matrix-Item ist
    if(var %in% matrix_vars_used) {
      # Finde die zugehörige Matrix-Hauptvariable
      for(i in 1:nrow(matrix_config)) {
        matrix_name <- matrix_config$variable_name[i]
        matrix_patterns <- c(
          paste0("^", matrix_name, "\\[.+\\]$"),
          paste0("^", matrix_name, "\\..+\\.$"),
          paste0("^", matrix_name, "_.+$"),
          paste0("^", matrix_name, "-.+$")
        )
        
        for(pattern in matrix_patterns) {
          if(grepl(pattern, var)) {
            return(paste0("Matrix-Item von ", matrix_name))
          }
        }
      }
      return("Matrix-Item")
    }
    
    return("")
  })
  
  # Data Type bestimmen
  overview$Data_Type <- sapply(all_vars, function(var) {
    # Erst Config prüfen
    config_row <- config$variablen[config$variablen$variable_name == var, ]
    if(nrow(config_row) > 0) {
      return(paste0(config_row$data_type[1], " (Config)"))
    } else {
      # Automatisch ermitteln
      return(paste0(detect_actual_data_type(data, var), " (Auto)"))
    }
  })
  
  # Factor Status
  overview$Factor <- sapply(all_vars, function(var) {
    if(is.factor(data[[var]])) "Ja" else "Nein"
  })
  
  # Nutzung in Analysen
  overview$In_Deskriptiven_Tabellen <- ifelse(all_vars %in% used_in_descriptive | all_vars %in% matrix_vars_used, "Ja", "Nein")  
  overview$In_Kreuztabellen <- ifelse(all_vars %in% used_in_crosstabs, "Ja", "Nein")
  overview$In_Regressionen <- ifelse(all_vars %in% used_in_regressions, "Ja", "Nein")
  overview$In_Textantworten <- ifelse(all_vars %in% used_in_textantworten, "Ja", "Nein")
  
  
  # Nach Alphabet sortieren
  overview <- overview[order(overview$Variable), ]
  
  cat("Variablen-Übersicht erstellt für", nrow(overview), "Variablen\n")
  
  return(overview)
}

# Variablen-Übersicht exportieren
export_variable_overview <- function(wb, variable_overview, header_style, table_style, title_style) {
  addWorksheet(wb, "Variablen_Übersicht")
  
  current_row <- 1
  
  # Titel
  writeData(wb, "Variablen_Übersicht", "Variablen-Übersicht", startRow = current_row)
  addStyle(wb, "Variablen_Übersicht", title_style, rows = current_row, cols = 1)
  current_row <- current_row + 2
  
  # Erklärung
  writeData(wb, "Variablen_Übersicht", 
            "Übersicht aller Variablen im Datensatz mit Informationen zu Datentyp, Factor-Status und Verwendung in Analysen.", 
            startRow = current_row)
  current_row <- current_row + 2
  
  # Tabelle schreiben
  writeData(wb, "Variablen_Übersicht", variable_overview, startRow = current_row, colNames = TRUE)
  addStyle(wb, "Variablen_Übersicht", header_style, rows = current_row, cols = 1:ncol(variable_overview))
  addStyle(wb, "Variablen_Übersicht", table_style, 
           rows = (current_row + 1):(current_row + nrow(variable_overview)), 
           cols = 1:ncol(variable_overview), gridExpand = TRUE)
  
  # Spaltenbreiten anpassen
  setColWidths(wb, "Variablen_Übersicht", cols = 1, widths = 25)  # Variable
  setColWidths(wb, "Variablen_Übersicht", cols = 2, widths = 40)  # Variable_Label
  setColWidths(wb, "Variablen_Übersicht", cols = 3, widths = 20)  # Data_Type
  setColWidths(wb, "Variablen_Übersicht", cols = 4, widths = 10)  # Factor
  setColWidths(wb, "Variablen_Übersicht", cols = 5:8, widths = 15) # Analysen (jetzt 4 Spalten)
  setColWidths(wb, "Variablen_Übersicht", cols = 9, widths = 25)  # Matrix_Info (jetzt Spalte 9)
  
}


# =============================================================================
# EXCEL EXPORT
# =============================================================================

export_results <- function(descriptive_results, crosstab_results, regression_results, text_results = NULL, variable_overview) {
  cat("\nExportiere Ergebnisse nach Excel...\n")
  
  # Workbook erstellen
  wb <- createWorkbook()
  
  # Stylesheet definieren
  header_style <- createStyle(
    fontName = "Arial",
    fontSize = 12,
    fontColour = "white",
    fgFill = "#4472C4",
    halign = "center",
    valign = "center",
    textDecoration = "bold",
    border = "TopBottomLeftRight",
    borderColour = "white"
  )
  
  table_style <- createStyle(
    fontName = "Arial",
    fontSize = 11,
    border = "TopBottomLeftRight",
    borderColour = "#D9D9D9"
  )
  
  title_style <- createStyle(
    fontName = "Arial",
    fontSize = 14,
    fontColour = "#4472C4",
    textDecoration = "bold"
  )
  
  # Sheet 1: Deskriptive Statistiken
  if (length(descriptive_results) > 0) {
    export_descriptive_statistics(wb, descriptive_results, header_style, table_style, title_style)
  }
  
  # Sheet 2: Kreuztabellen
  if (length(crosstab_results) > 0) {
    export_crosstabs(wb, crosstab_results, header_style, table_style, title_style)
  }
  
  # Sheet 3: Statistische Tests
  if (length(crosstab_results) > 0) {
    export_statistical_tests(wb, crosstab_results, header_style, table_style, title_style)
  }
  
  # Sheet 4: Regressionsanalysen
  if (length(regression_results) > 0) {
    export_regressions(wb, regression_results, header_style, table_style, title_style)
  }
  
  # Sheet 5: Textantworten
  if (!is.null(text_results) && length(text_results) > 0) {
    export_text_responses(wb, text_results, header_style, table_style, title_style)
  }
  
  # Sheet 6: Variablen-Übersicht (NEU)
  export_variable_overview(wb, variable_overview, header_style, table_style, title_style)
  
  # Excel-Datei speichern
  saveWorkbook(wb, OUTPUT_FILE, overwrite = TRUE)
  cat("Ergebnisse erfolgreich exportiert nach:", OUTPUT_FILE, "\n")
}

# Deskriptive Statistiken exportieren
export_descriptive_statistics <- function(wb, descriptive_results, header_style, table_style, title_style) {
  addWorksheet(wb, "Deskriptive_Statistiken")
  
  current_row <- 1
  
  # Titel
  writeData(wb, "Deskriptive_Statistiken", "Deskriptive Statistiken", startRow = current_row)
  addStyle(wb, "Deskriptive_Statistiken", title_style, rows = current_row, cols = 1)
  current_row <- current_row + 2
  
  for (var_name in names(descriptive_results)) {
    result <- descriptive_results[[var_name]]
    
    # Variable Überschrift
    writeData(wb, "Deskriptive_Statistiken", 
              paste("Variable:", var_name, "-", result$question), 
              startRow = current_row)
    addStyle(wb, "Deskriptive_Statistiken", title_style, rows = current_row, cols = 1)
    current_row <- current_row + 1
    
    # Gewichtung info
    if (result$weighted) {
      writeData(wb, "Deskriptive_Statistiken", "Gewichtete Ergebnisse", startRow = current_row)
      current_row <- current_row + 1
    }
    
    # Filter info (falls angewendet)
    if (!is.null(result$filter_applied) && result$filter_applied) {
      filter_text <- paste0("Filter angewendet: ", result$filter_string, " (", 
                            result$filtered_n, " von ", result$original_n, " Fällen, ",
                            round(result$filtered_n/result$original_n*100, 1), "%)")
      writeData(wb, "Deskriptive_Statistiken", filter_text, startRow = current_row)
      current_row <- current_row + 1
    }
    
    # Tabelle(n) schreiben basierend auf Typ
    if (result$type == "ordinal" && "table_frequencies" %in% names(result)) {
      # ORDINALE VARIABLEN - Häufigkeiten + Numerische Kennwerte
      writeData(wb, "Deskriptive_Statistiken", "Häufigkeiten:", startRow = current_row)
      current_row <- current_row + 1
      writeData(wb, "Deskriptive_Statistiken", result$table_frequencies, startRow = current_row, colNames = TRUE)
      addStyle(wb, "Deskriptive_Statistiken", header_style, rows = current_row, cols = 1:ncol(result$table_frequencies))
      addStyle(wb, "Deskriptive_Statistiken", table_style, 
               rows = (current_row + 1):(current_row + nrow(result$table_frequencies)), 
               cols = 1:ncol(result$table_frequencies), gridExpand = TRUE)
      current_row <- current_row + nrow(result$table_frequencies) + 2
      
      # Numerische Kennwerte
      writeData(wb, "Deskriptive_Statistiken", "Numerische Kennwerte:", startRow = current_row)
      current_row <- current_row + 1
      writeData(wb, "Deskriptive_Statistiken", result$table_numeric, startRow = current_row, colNames = TRUE)
      addStyle(wb, "Deskriptive_Statistiken", header_style, rows = current_row, cols = 1:ncol(result$table_numeric))
      addStyle(wb, "Deskriptive_Statistiken", table_style, 
               rows = (current_row + 1):(current_row + nrow(result$table_numeric)), 
               cols = 1:ncol(result$table_numeric), gridExpand = TRUE)
      current_row <- current_row + nrow(result$table_numeric) + 3
      
    } else if (result$type %in% c("matrix_ordinal", "matrix_dichotomous", "matrix_numeric") && "table_categorical" %in% names(result)) {
      # MATRIX ORDINAL - Kategoriale Häufigkeiten + Numerische Kennwerte
      writeData(wb, "Deskriptive_Statistiken", "Kategoriale Häufigkeiten:", startRow = current_row)
      current_row <- current_row + 1
      writeData(wb, "Deskriptive_Statistiken", result$table_categorical, startRow = current_row, colNames = TRUE)
      addStyle(wb, "Deskriptive_Statistiken", header_style, rows = current_row, cols = 1:ncol(result$table_categorical))
      addStyle(wb, "Deskriptive_Statistiken", table_style, 
               rows = (current_row + 1):(current_row + nrow(result$table_categorical)), 
               cols = 1:ncol(result$table_categorical), gridExpand = TRUE)
      current_row <- current_row + nrow(result$table_categorical) + 2
      
      # Numerische Kennwerte für Matrix-Items
      if ("table_numeric" %in% names(result)) {
        writeData(wb, "Deskriptive_Statistiken", "Numerische Kennwerte (ordinale Skala):", startRow = current_row)
        current_row <- current_row + 1
        writeData(wb, "Deskriptive_Statistiken", result$table_numeric, startRow = current_row, colNames = TRUE)
        addStyle(wb, "Deskriptive_Statistiken", header_style, rows = current_row, cols = 1:ncol(result$table_numeric))
        addStyle(wb, "Deskriptive_Statistiken", table_style, 
                 rows = (current_row + 1):(current_row + nrow(result$table_numeric)), 
                 cols = 1:ncol(result$table_numeric), gridExpand = TRUE)
        current_row <- current_row + nrow(result$table_numeric) + 3
      }
      
    } else {
      # ALLE ANDEREN TYPEN - Standard Tabelle (numeric, nominal_coded, nominal_text, dichotom)
      # Prüfe ob table vorhanden ist
      if (!is.null(result$table) && nrow(result$table) > 0) {
        writeData(wb, "Deskriptive_Statistiken", result$table, startRow = current_row, colNames = TRUE)
        addStyle(wb, "Deskriptive_Statistiken", header_style, rows = current_row, cols = 1:ncol(result$table))
        addStyle(wb, "Deskriptive_Statistiken", table_style, 
                 rows = (current_row + 1):(current_row + nrow(result$table)), 
                 cols = 1:ncol(result$table), gridExpand = TRUE)
        current_row <- current_row + nrow(result$table) + 3
      } else {
        # Fallback: Versuche table_categorical
        if (!is.null(result$table_categorical) && nrow(result$table_categorical) > 0) {
          writeData(wb, "Deskriptive_Statistiken", result$table_categorical, startRow = current_row, colNames = TRUE)
          addStyle(wb, "Deskriptive_Statistiken", header_style, rows = current_row, cols = 1:ncol(result$table_categorical))
          addStyle(wb, "Deskriptive_Statistiken", table_style, 
                   rows = (current_row + 1):(current_row + nrow(result$table_categorical)), 
                   cols = 1:ncol(result$table_categorical), gridExpand = TRUE)
          current_row <- current_row + nrow(result$table_categorical) + 2
          
          # Prüfe auch table_numeric
          if (!is.null(result$table_numeric) && nrow(result$table_numeric) > 0) {
            writeData(wb, "Deskriptive_Statistiken", result$table_numeric, startRow = current_row, colNames = TRUE)
            addStyle(wb, "Deskriptive_Statistiken", header_style, rows = current_row, cols = 1:ncol(result$table_numeric))
            addStyle(wb, "Deskriptive_Statistiken", table_style, 
                     rows = (current_row + 1):(current_row + nrow(result$table_numeric)), 
                     cols = 1:ncol(result$table_numeric), gridExpand = TRUE)
            current_row <- current_row + nrow(result$table_numeric) + 3
          }
        } else {
          cat("WARNUNG: Keine exportierbare Tabelle für Variable", var_name, "(Typ:", result$type, ")\n")
          writeData(wb, "Deskriptive_Statistiken", paste("WARNUNG: Keine Daten für", var_name), startRow = current_row)
          current_row <- current_row + 2
        }
      }
    }
  }
  
  # Spaltenbreite anpassen
  setColWidths(wb, "Deskriptive_Statistiken", cols = 1:10, widths = "20")
}

# Kreuztabellen exportieren
export_crosstabs <- function(wb, crosstab_results, header_style, table_style, title_style) {
  addWorksheet(wb, "Kreuztabellen")
  
  current_row <- 1
  
  # Titel
  writeData(wb, "Kreuztabellen", "Kreuztabellen", startRow = current_row)
  addStyle(wb, "Kreuztabellen", title_style, rows = current_row, cols = 1)
  current_row <- current_row + 2
  
  for (analysis_name in names(crosstab_results)) {
    result <- crosstab_results[[analysis_name]]
    
    # Analyse Überschrift
    writeData(wb, "Kreuztabellen", 
              paste("Analyse:", analysis_name, "(", result$variable_1, "x", result$variable_2, ")"), 
              startRow = current_row)
    addStyle(wb, "Kreuztabellen", title_style, rows = current_row, cols = 1)
    current_row <- current_row + 1
    
    # Gewichtung info
    if (result$weighted) {
      writeData(wb, "Kreuztabellen", "Gewichtete Ergebnisse", startRow = current_row)
      current_row <- current_row + 1
    }
    
    # Filter info (falls angewendet)
    if (!is.null(result$filter_applied) && result$filter_applied) {
      filter_text <- paste0("Filter angewendet: ", result$filter_string, " (", 
                            result$filtered_n, " von ", result$original_n, " Fällen, ",
                            round(result$filtered_n/result$original_n*100, 1), "%)")
      writeData(wb, "Kreuztabellen", filter_text, startRow = current_row)
      current_row <- current_row + 1
    }
    
    
    if (!is.null(result$crosstab)) {
      
      # NEU: Matrix-Kreuztabellen (als ersten else-if Block einfügen)
      if ("matrix_items" %in% names(result$crosstab)) {
        cat("Exportiere Matrix-Kreuztabelle:", analysis_name, "\n")
        
        # Matrix-Info
        writeData(wb, "Kreuztabellen", 
                  paste("Matrix-Variable:", result$crosstab$var1_name, "mit", 
                        length(result$crosstab$matrix_items), "Items"), 
                  startRow = current_row)
        current_row <- current_row + 1
        
        # Kategoriale Tabelle
        if (!is.null(result$crosstab$categorical)) {
          writeData(wb, "Kreuztabellen", "Kategoriale Häufigkeiten:", startRow = current_row)
          current_row <- current_row + 1
          writeData(wb, "Kreuztabellen", result$crosstab$categorical, startRow = current_row, colNames = TRUE)
          addStyle(wb, "Kreuztabellen", header_style, rows = current_row, cols = 1:ncol(result$crosstab$categorical))
          addStyle(wb, "Kreuztabellen", table_style, 
                   rows = (current_row + 1):(current_row + nrow(result$crosstab$categorical)), 
                   cols = 1:ncol(result$crosstab$categorical), gridExpand = TRUE)
          current_row <- current_row + nrow(result$crosstab$categorical) + 2
        }
        
        # KORRIGIERT: Numerische Tabelle exportieren
        if (!is.null(result$crosstab$numeric)) {
          writeData(wb, "Kreuztabellen", "Numerische Kennwerte nach Gruppen:", startRow = current_row)
          current_row <- current_row + 1
          writeData(wb, "Kreuztabellen", result$crosstab$numeric, startRow = current_row, colNames = TRUE)
          addStyle(wb, "Kreuztabellen", header_style, rows = current_row, cols = 1:ncol(result$crosstab$numeric))
          addStyle(wb, "Kreuztabellen", table_style, 
                   rows = (current_row + 1):(current_row + nrow(result$crosstab$numeric)), 
                   cols = 1:ncol(result$crosstab$numeric), gridExpand = TRUE)
          current_row <- current_row + nrow(result$crosstab$numeric) + 3
        } else {
          writeData(wb, "Kreuztabellen", "Keine numerischen Kennwerte (kategoriale Matrix)", startRow = current_row)
          current_row <- current_row + 2
        }
      } else if ("type" %in% names(result$crosstab) && result$crosstab$type == "group_means") {
        # Gruppenmittelwerte exportieren
        writeData(wb, "Kreuztabellen", "Gruppenmittelwerte:", startRow = current_row)
        current_row <- current_row + 1
        writeData(wb, "Kreuztabellen", result$crosstab$group_means, startRow = current_row, colNames = TRUE)
        addStyle(wb, "Kreuztabellen", header_style, rows = current_row, cols = 1:ncol(result$crosstab$group_means))
        addStyle(wb, "Kreuztabellen", table_style, 
                 rows = (current_row + 1):(current_row + nrow(result$crosstab$group_means)), 
                 cols = 1:ncol(result$crosstab$group_means), gridExpand = TRUE)
        current_row <- current_row + nrow(result$crosstab$group_means) + 3
      } else if ("type" %in% names(result$crosstab) && result$crosstab$type == "correlation") {
        # Korrelationsanalyse exportieren
        writeData(wb, "Kreuztabellen", "Korrelationsanalyse:", startRow = current_row)
        current_row <- current_row + 1
        writeData(wb, "Kreuztabellen", result$crosstab$correlation_table, startRow = current_row, colNames = TRUE)
        addStyle(wb, "Kreuztabellen", header_style, rows = current_row, cols = 1:ncol(result$crosstab$correlation_table))
        addStyle(wb, "Kreuztabellen", table_style, 
                 rows = (current_row + 1):(current_row + nrow(result$crosstab$correlation_table)), 
                 cols = 1:ncol(result$crosstab$correlation_table), gridExpand = TRUE)
        current_row <- current_row + nrow(result$crosstab$correlation_table) + 3
      } else {
        # Standard Kreuztabellen exportieren
        # Absolute Häufigkeiten
        writeData(wb, "Kreuztabellen", "Absolute Häufigkeiten:", startRow = current_row)
        current_row <- current_row + 1
        absolute_table <- result$crosstab$absolute
        absolute_table_with_rownames <- cbind(Variable = rownames(absolute_table), absolute_table)
        writeData(wb, "Kreuztabellen", absolute_table_with_rownames, startRow = current_row, colNames = TRUE)
        addStyle(wb, "Kreuztabellen", header_style, rows = current_row, cols = 1:(ncol(result$crosstab$absolute) + 1))
        addStyle(wb, "Kreuztabellen", table_style, 
                 rows = (current_row + 1):(current_row + nrow(result$crosstab$absolute)), 
                 cols = 1:(ncol(result$crosstab$absolute) + 1), gridExpand = TRUE)
        current_row <- current_row + nrow(result$crosstab$absolute) + 2
        
        # Relative Häufigkeiten (Zeilenprozente)
        writeData(wb, "Kreuztabellen", "Relative Häufigkeiten (Zeilenprozente):", startRow = current_row)
        current_row <- current_row + 1
        relative_table <- result$crosstab$relative
        relative_table_with_rownames <- cbind(Variable = rownames(relative_table), relative_table)
        writeData(wb, "Kreuztabellen", relative_table_with_rownames, startRow = current_row, colNames = TRUE)
        addStyle(wb, "Kreuztabellen", header_style, rows = current_row, cols = 1:(ncol(result$crosstab$relative) + 1))
        addStyle(wb, "Kreuztabellen", table_style, 
                 rows = (current_row + 1):(current_row + nrow(result$crosstab$relative)), 
                 cols = 1:(ncol(result$crosstab$relative) + 1), gridExpand = TRUE)
        current_row <- current_row + nrow(result$crosstab$relative) + 3
      }
    } else {
      writeData(wb, "Kreuztabellen", "Keine Daten verfügbar", startRow = current_row)
      current_row <- current_row + 3
    }
  }
  
  # Spaltenbreite anpassen
  setColWidths(wb, "Kreuztabellen", cols = 1:10, widths = "20")
}

# Statistische Tests exportieren
export_statistical_tests <- function(wb, crosstab_results, header_style, table_style, title_style) {
  addWorksheet(wb, "Statistische_Tests")
  
  current_row <- 1
  
  # Titel
  writeData(wb, "Statistische_Tests", "Statistische Tests", startRow = current_row)
  addStyle(wb, "Statistische_Tests", title_style, rows = current_row, cols = 1)
  current_row <- current_row + 2
  
  # Nur fortfahren wenn es Ergebnisse gibt
  if (length(crosstab_results) == 0) {
    writeData(wb, "Statistische_Tests", "Keine statistischen Tests durchgeführt.", startRow = current_row)
    return()
  }
  
  # Übersichtstabelle erstellen
  test_summary <- data.frame(
    Analyse = character(),
    Variable_1 = character(),
    Variable_2 = character(),
    Test = character(),
    Statistik = character(),
    p_Wert = character(),
    Ergebnis = character(),
    stringsAsFactors = FALSE
  )
  
  for (analysis_name in names(crosstab_results)) {
    result <- crosstab_results[[analysis_name]]
    
    if (!is.null(result$statistical_test)) {
      test <- result$statistical_test
      
      # Prüfen ob Matrix-Test mit mehreren Item-Tests
      if (!is.null(test$item_tests)) {
        # Für jeden Item-Test einen Eintrag hinzufügen
        for (item_name in names(test$item_tests)) {
          item_test <- test$item_tests[[item_name]]
          test_summary <- rbind(test_summary, data.frame(
            Analyse = paste0(analysis_name, " [", item_name, "]"),
            Variable_1 = result$variable_1,
            Variable_2 = result$variable_2,
            Test = item_test$test,
            Statistik = if(!is.na(item_test$statistic)) as.character(item_test$statistic) else "-",
            p_Wert = if(!is.na(item_test$p_value)) as.character(item_test$p_value) else "-",
            Ergebnis = item_test$result,
            stringsAsFactors = FALSE
          ))
        }
      } else {
        # Normaler Test
        test_summary <- rbind(test_summary, data.frame(
          Analyse = analysis_name,
          Variable_1 = result$variable_1,
          Variable_2 = result$variable_2,
          Test = test$test,
          Statistik = if(!is.na(test$statistic)) as.character(test$statistic) else "-",
          p_Wert = if(!is.na(test$p_value)) as.character(test$p_value) else "-",
          Ergebnis = test$result,
          stringsAsFactors = FALSE
        ))
      }
    }
  }
  
  # Nur Übersichtstabelle schreiben wenn Daten vorhanden
  if (nrow(test_summary) > 0) {
    writeData(wb, "Statistische_Tests", "Übersicht aller Tests:", startRow = current_row)
    current_row <- current_row + 1
    writeData(wb, "Statistische_Tests", test_summary, startRow = current_row, colNames = TRUE)
    addStyle(wb, "Statistische_Tests", header_style, rows = current_row, cols = 1:ncol(test_summary))
    addStyle(wb, "Statistische_Tests", table_style, 
             rows = (current_row + 1):(current_row + nrow(test_summary)), 
             cols = 1:ncol(test_summary), gridExpand = TRUE)
    current_row <- current_row + nrow(test_summary) + 3
    
    # Detaillierte Testergebnisse
    writeData(wb, "Statistische_Tests", "Detaillierte Testergebnisse:", startRow = current_row)
    current_row <- current_row + 2
    
    for (analysis_name in names(crosstab_results)) {
      result <- crosstab_results[[analysis_name]]
      
      if (!is.null(result$statistical_test)) {
        test <- result$statistical_test
        
        writeData(wb, "Statistische_Tests", paste("Analyse:", analysis_name), startRow = current_row)
        addStyle(wb, "Statistische_Tests", title_style, rows = current_row, cols = 1)
        current_row <- current_row + 1
        
        # Prüfen ob es sich um einen Matrix-Test handelt
        if (!is.null(test$item_tests)) {
          # Matrix-Kreuztabelle mit Tests pro Item
          writeData(wb, "Statistische_Tests", "Matrix-Item Tests:", startRow = current_row)
          current_row <- current_row + 1
          
          # Tabellarische Übersicht der Matrix-Item Tests
          matrix_test_summary <- data.frame(
            Item = names(test$item_tests),
            Test = sapply(test$item_tests, `[[`, "test"),
            Statistik = sapply(test$item_tests, function(x) if(!is.na(x$statistic)) as.character(x$statistic) else "-"),
            p_Wert = sapply(test$item_tests, function(x) if(!is.na(x$p_value)) as.character(x$p_value) else "-"),
            Ergebnis = sapply(test$item_tests, `[[`, "result"),
            stringsAsFactors = FALSE
          )
          
          writeData(wb, "Statistische_Tests", matrix_test_summary, startRow = current_row, colNames = TRUE)
          addStyle(wb, "Statistische_Tests", header_style, rows = current_row, cols = 1:5)
          addStyle(wb, "Statistische_Tests", table_style, 
                   rows = (current_row + 1):(current_row + nrow(matrix_test_summary)), 
                   cols = 1:5, gridExpand = TRUE)
          current_row <- current_row + nrow(matrix_test_summary) + 2
        } else {
          # Normaler Kreuztabellen-Test
          test_details <- data.frame(
            Parameter = character(),
            Wert = character(),
            stringsAsFactors = FALSE
          )
          
          # Basis-Parameter
          test_details <- rbind(test_details, data.frame(Parameter = "Test", Wert = test$test))
          test_details <- rbind(test_details, data.frame(Parameter = "Statistik", Wert = if(!is.na(test$statistic)) as.character(test$statistic) else "-"))
          test_details <- rbind(test_details, data.frame(Parameter = "p-Wert", Wert = if(!is.na(test$p_value)) as.character(test$p_value) else "-"))
          test_details <- rbind(test_details, data.frame(Parameter = "Ergebnis", Wert = test$result))
          
          # Freiheitsgrade nur wenn vorhanden
          if ("df" %in% names(test) && !is.null(test$df) && length(test$df) == 1 && !is.na(test$df)) {
            test_details <- rbind(test_details, data.frame(Parameter = "Freiheitsgrade", Wert = as.character(test$df)))
          }
          
          writeData(wb, "Statistische_Tests", test_details, startRow = current_row, colNames = TRUE)
          addStyle(wb, "Statistische_Tests", header_style, rows = current_row, cols = 1:2)
          addStyle(wb, "Statistische_Tests", table_style, 
                   rows = (current_row + 1):(current_row + nrow(test_details)), 
                   cols = 1:2, gridExpand = TRUE)
          current_row <- current_row + nrow(test_details) + 2
        }
      }
    }
  } else {
    writeData(wb, "Statistische_Tests", "Keine statistischen Tests erfolgreich durchgeführt.", startRow = current_row)
  }
  
  # Spaltenbreite anpassen
  setColWidths(wb, "Statistische_Tests", cols = 1:7, widths = "auto")
}


# Regressionen exportieren
export_regressions_old <- function(wb, regression_results, header_style, table_style, title_style) {
  addWorksheet(wb, "Regressionsanalysen")
  
  current_row <- 1
  
  # Titel
  writeData(wb, "Regressionsanalysen", "Regressionsanalysen", startRow = current_row)
  addStyle(wb, "Regressionsanalysen", title_style, rows = current_row, cols = 1)
  current_row <- current_row + 2
  
  for (reg_name in names(regression_results)) {
    result <- regression_results[[reg_name]]
    
    # Regression Überschrift
    writeData(wb, "Regressionsanalysen", 
              paste("Modell:", reg_name, "(", result$regression_type, ")"), 
              startRow = current_row)
    addStyle(wb, "Regressionsanalysen", title_style, rows = current_row, cols = 1)
    current_row <- current_row + 1
    
    # Modell-Info
    writeData(wb, "Regressionsanalysen", 
              paste("AV:", result$dependent_var, "| UV:", paste(result$independent_vars, collapse = ", ")), 
              startRow = current_row)
    current_row <- current_row + 1
    
    writeData(wb, "Regressionsanalysen", 
              paste("N =", result$n_complete, "| Gewichtet:", result$weighted), 
              startRow = current_row)
    current_row <- current_row + 2
    
    # Koeffizienten
    writeData(wb, "Regressionsanalysen", "Koeffizienten:", startRow = current_row)
    current_row <- current_row + 1
    writeData(wb, "Regressionsanalysen", result$coefficients, startRow = current_row, colNames = TRUE)
    addStyle(wb, "Regressionsanalysen", header_style, rows = current_row, cols = 1:ncol(result$coefficients))
    addStyle(wb, "Regressionsanalysen", table_style, 
             rows = (current_row + 1):(current_row + nrow(result$coefficients)), 
             cols = 1:ncol(result$coefficients), gridExpand = TRUE)
    current_row <- current_row + nrow(result$coefficients) + 2
    
    # Modell-Güte
    writeData(wb, "Regressionsanalysen", "Modell-Güte:", startRow = current_row)
    current_row <- current_row + 1
    writeData(wb, "Regressionsanalysen", result$model_fit, startRow = current_row, colNames = TRUE)
    addStyle(wb, "Regressionsanalysen", header_style, rows = current_row, cols = 1:ncol(result$model_fit))
    addStyle(wb, "Regressionsanalysen", table_style, 
             rows = (current_row + 1):(current_row + nrow(result$model_fit)), 
             cols = 1:ncol(result$model_fit), gridExpand = TRUE)
    current_row <- current_row + nrow(result$model_fit) + 3
  }
  
  # Spaltenbreite anpassen
  setColWidths(wb, "Regressionsanalysen", cols = 1:7, widths = "auto")
}

# Erweitern Sie export_regressions() um Mehrebenenmodelle:
export_regressions <- function(wb, regression_results, header_style, table_style, title_style) {
  addWorksheet(wb, "Regressionsanalysen")
  
  current_row <- 1
  
  # Titel
  writeData(wb, "Regressionsanalysen", "Regressionsanalysen", startRow = current_row)
  addStyle(wb, "Regressionsanalysen", title_style, rows = current_row, cols = 1)
  current_row <- current_row + 2
  
  for (reg_name in names(regression_results)) {
    result <- regression_results[[reg_name]]
    
    # Regression Überschrift
    writeData(wb, "Regressionsanalysen", 
              paste("Modell:", reg_name, "(", result$regression_type, ")"), 
              startRow = current_row)
    addStyle(wb, "Regressionsanalysen", title_style, rows = current_row, cols = 1)
    current_row <- current_row + 1
    
    # Modell-Info
    writeData(wb, "Regressionsanalysen", 
              paste("AV:", result$dependent_var, "| UV:", paste(result$independent_vars, collapse = ", ")), 
              startRow = current_row)
    current_row <- current_row + 1
    
    # MEHREBENEN-SPEZIFISCHE INFO
    if (grepl("multilevel", result$type)) {
      writeData(wb, "Regressionsanalysen", 
                paste("Clustering:", result$cluster_variable, "| Level-1:", paste(result$level1_variables, collapse = ", "), 
                      "| Level-2:", paste(result$level2_variables, collapse = ", ")), 
                startRow = current_row)
      current_row <- current_row + 1
      
      writeData(wb, "Regressionsanalysen", 
                paste("N =", result$model_fit$Wert[result$model_fit$Kennwert == "N"], 
                      "| Cluster =", result$n_clusters, "| Gewichtet:", result$weighted), 
                startRow = current_row)
    } else {
      writeData(wb, "Regressionsanalysen", 
                paste("N =", result$n_complete, "| Gewichtet:", result$weighted), 
                startRow = current_row)
    }
    current_row <- current_row + 2
    
    # Fixed Effects / Koeffizienten
    writeData(wb, "Regressionsanalysen", "Fixed Effects / Koeffizienten:", startRow = current_row)
    current_row <- current_row + 1
    writeData(wb, "Regressionsanalysen", result$coefficients, startRow = current_row, colNames = TRUE)
    addStyle(wb, "Regressionsanalysen", header_style, rows = current_row, cols = 1:ncol(result$coefficients))
    addStyle(wb, "Regressionsanalysen", table_style, 
             rows = (current_row + 1):(current_row + nrow(result$coefficients)), 
             cols = 1:ncol(result$coefficients), gridExpand = TRUE)
    current_row <- current_row + nrow(result$coefficients) + 2
    
    # Random Effects (nur für Mehrebenenmodelle)
    if (grepl("multilevel", result$type) && "random_effects" %in% names(result)) {
      writeData(wb, "Regressionsanalysen", "Random Effects:", startRow = current_row)
      current_row <- current_row + 1
      writeData(wb, "Regressionsanalysen", result$random_effects, startRow = current_row, colNames = TRUE)
      addStyle(wb, "Regressionsanalysen", header_style, rows = current_row, cols = 1:ncol(result$random_effects))
      addStyle(wb, "Regressionsanalysen", table_style, 
               rows = (current_row + 1):(current_row + nrow(result$random_effects)), 
               cols = 1:ncol(result$random_effects), gridExpand = TRUE)
      current_row <- current_row + nrow(result$random_effects) + 2
    }
    
    # Modell-Güte
    writeData(wb, "Regressionsanalysen", "Modell-Güte:", startRow = current_row)
    current_row <- current_row + 1
    writeData(wb, "Regressionsanalysen", result$model_fit, startRow = current_row, colNames = TRUE)
    addStyle(wb, "Regressionsanalysen", header_style, rows = current_row, cols = 1:ncol(result$model_fit))
    addStyle(wb, "Regressionsanalysen", table_style, 
             rows = (current_row + 1):(current_row + nrow(result$model_fit)), 
             cols = 1:ncol(result$model_fit), gridExpand = TRUE)
    current_row <- current_row + nrow(result$model_fit) + 3
  }
  
  # Spaltenbreite anpassen
  setColWidths(wb, "Regressionsanalysen", cols = 1:8, widths = "auto")
}

# Textantworten exportieren
export_text_responses <- function(wb, text_results, header_style, table_style, title_style) {
  addWorksheet(wb, "Textantworten")
  
  current_row <- 1
  
  # Titel
  writeData(wb, "Textantworten", "Offene Textantworten", startRow = current_row)
  addStyle(wb, "Textantworten", title_style, rows = current_row, cols = 1)
  current_row <- current_row + 2
  
  for (analysis_name in names(text_results)) {
    result <- text_results[[analysis_name]]
    
    # Analyse Überschrift
    writeData(wb, "Textantworten", 
              paste("Analyse:", analysis_name), 
              startRow = current_row)
    addStyle(wb, "Textantworten", title_style, rows = current_row, cols = 1)
    current_row <- current_row + 1
    
    # Analyse-Info
    writeData(wb, "Textantworten", 
              paste("Variable:", result$text_variable, "| Sortiert nach:", result$sort_variable, "| Min. Länge:", result$min_length), 
              startRow = current_row)
    current_row <- current_row + 2
    
    # Zusammenfassung
    writeData(wb, "Textantworten", "Zusammenfassung:", startRow = current_row)
    current_row <- current_row + 1
    writeData(wb, "Textantworten", result$summary, startRow = current_row, colNames = TRUE)
    addStyle(wb, "Textantworten", header_style, rows = current_row, cols = 1:ncol(result$summary))
    addStyle(wb, "Textantworten", table_style, 
             rows = (current_row + 1):(current_row + nrow(result$summary)), 
             cols = 1:ncol(result$summary), gridExpand = TRUE)
    current_row <- current_row + nrow(result$summary) + 2
    
    # Alle Textantworten
    writeData(wb, "Textantworten", "Alle Textantworten:", startRow = current_row)
    current_row <- current_row + 1
    writeData(wb, "Textantworten", result$responses, startRow = current_row, colNames = TRUE)
    addStyle(wb, "Textantworten", header_style, rows = current_row, cols = 1:ncol(result$responses))
    addStyle(wb, "Textantworten", table_style, 
             rows = (current_row + 1):(current_row + nrow(result$responses)), 
             cols = 1:ncol(result$responses), gridExpand = TRUE)
    current_row <- current_row + nrow(result$responses) + 3
  }
  
  # Spaltenbreite anpassen
  setColWidths(wb, "Textantworten", cols = 1, widths = 25)  # Kategorie
  setColWidths(wb, "Textantworten", cols = 2, widths = 80)  # Textantwort
  setColWidths(wb, "Textantworten", cols = 3, widths = 10)  # Zeichen
}

# =============================================================================
# NEUE FUNKTION: FINALEN DATENSATZ SPEICHERN
# =============================================================================

save_final_dataset <- function(data, config) {
  if (!SAVE_FINAL_DATASET) {
    cat("Speichern des finalen Datensatzes ist deaktiviert.\n")
    return()
  }
  
  cat("\nSpeichere finalen Datensatz...\n")
  
  # Erstelle Ausgabe-Verzeichnis falls nicht vorhanden
  output_dir <- dirname(FINAL_DATASET_FILE)
  if (!dir.exists(output_dir)) {
    dir.create(output_dir, recursive = TRUE)
    cat("Verzeichnis erstellt:", output_dir, "\n")
  }
  
  # Füge Metadaten als Attribute hinzu
  attr(data, "processing_info") <- list(
    processing_date = Sys.time(),
    original_variables = ncol(data) - length(grep("(_index|_num|_binary|_quote|_avg|_kat)$", names(data))),
    created_variables = length(grep("(_index|_num|_binary|_quote|_avg|_kat)$", names(data))),
    total_variables = ncol(data),
    n_observations = nrow(data),
    config_variables = nrow(config$variablen),
    weights_used = WEIGHTS,
    weight_variable = if(WEIGHTS) WEIGHT_VAR else NA
  )
  
  # Speichere als RDS (behält alle Attribute und Datentypen)
  saveRDS(data, FINAL_DATASET_FILE)
  cat("Finaler Datensatz gespeichert als:", FINAL_DATASET_FILE, "\n")
  cat("✓ Finaler Datensatz erfolgreich gespeichert\n")
}


# =============================================================================
# HAUPTPROGRAMM
# =============================================================================

main <- function() {
  
  # Sicherstellen dass benötigte Variablen existieren
  if (!exists("index_definitions")) {
    index_definitions <- list()
  }
  if (!exists("custom_var_labels")) {
    custom_var_labels <- NULL
  }
  if (!exists("custom_val_labels")) {
    custom_val_labels <- NULL
  }
  
  cat("=============================================================================\n")
  cat("SURVEY DATENAUSWERTUNG - START\n")
  cat("=============================================================================\n")
  
  # 1. Setup
  cat("\n1. SETUP\n")
  cat("---------\n")
  load_packages()
  
  # 2. Konfiguration laden
  cat("\n2. KONFIGURATION LADEN\n")
  cat("-----------------------\n")
  config <- load_config()
  
  # 3. Daten laden und vorbereiten
  cat("\n3. DATEN LADEN UND VORBEREITEN\n")
  cat("-------------------------------\n")
  prepared_data <- load_and_prepare_data(config, index_definitions, custom_var_labels, custom_val_labels)
  
  # 4. Deskriptive Analysen
  cat("\n4. DESKRIPTIVE STATISTIKEN\n")
  cat("---------------------------\n")
  descriptive_results <- create_descriptive_tables(prepared_data)
  
  # 5. Ergebnisse anzeigen (vorläufig)
  cat("\n5. ERGEBNISSE (VORSCHAU)\n")
  cat("-------------------------\n")
  for (var_name in names(descriptive_results)) {
    result <- descriptive_results[[var_name]]
    cat("\nVariable:", var_name, "(", result$type, ")\n")
    cat("Frage:", result$question, "\n")
    cat("Gewichtet:", result$weighted, "\n")
    
    if (result$type %in% c("ordinal") && "table_frequencies" %in% names(result)) {
      cat("Häufigkeiten:\n")
      print(result$table_frequencies)
      cat("Numerische Kennwerte:\n")
      print(result$table_numeric)
      
    } else if (result$type == "matrix") {
      cat("Matrix mit", result$n_items, "Items:\n")
      print(result$table)
      
    } else if (result$type == "matrix_ordinal") {
      cat("Ordinale Matrix mit", result$n_items, "Items:\n")
      cat("Kategoriale Häufigkeiten:\n")
      print(result$table_categorical)
      cat("Numerische Kennwerte:\n")
      print(result$table_numeric)
      
    } else if (result$type == "matrix_dichotomous") {
      cat("Dichotome Matrix mit", result$n_items, "Items:\n")
      cat("Kategoriale Häufigkeiten:\n")
      print(result$table_categorical)
      cat("Numerische Kennwerte:\n")
      print(result$table_numeric)
      
    } else {
      print(result$table)
    }
    cat("\n", rep("-", 50), "\n")
  }
  # 5. Kreuztabellen
  cat("\n5. KREUZTABELLEN\n")
  cat("----------------\n")
  crosstab_results <- create_crosstabs(prepared_data)
  
  # 6. Regressionen
  cat("\n6. REGRESSIONSANALYSEN\n")
  cat("----------------------\n")
  regression_results <- run_regressions(prepared_data)
  
  # 7. Textantworten - NEUE ERGÄNZUNG
  cat("\n7. TEXTANTWORTEN\n")
  cat("----------------\n")
  text_results <- process_text_responses(prepared_data, custom_val_labels)
  
  # 8. Export (Nummer angepasst)
  cat("\n8. EXCEL EXPORT\n")
  cat("---------------\n")
  
  # Variablen-Übersicht erstellen
  cat("Erstelle Variablen-Übersicht...\n")
  variable_overview <- create_variable_overview(
    prepared_data$data, 
    prepared_data$config, 
    descriptive_results, 
    crosstab_results, 
    regression_results,
    text_results,
    custom_var_labels
  )
  
  # Excel Export mit allen Ergebnissen
  export_results(
    descriptive_results, 
    crosstab_results, 
    regression_results, 
    text_results, 
    variable_overview
  )
  
  # *** FINALEN DATENSATZ SPEICHERN ***
  save_final_dataset(prepared_data$data, prepared_data$config)
}