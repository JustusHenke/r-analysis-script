# =============================================================================
# ZENTRALISIERTE LABEL-VERWALTUNG - ALL-IN-ONE
# =============================================================================
# Alle Labels kommen aus dem globalen Codebook (erstellt beim Daten-Laden)
# Config kann Labels überschreiben
# =============================================================================

get_labels <- function(var_name, label_type = "value", config = NULL) {
  "Zentrale Label-Funktion für ALLE Tabellentypen
  
  Args:
    var_name: Name der Variable
    label_type: 'value' (Werte-Labels), 'variable' (Variablen-Label), 'item' (Matrix-Item-Label)
    config: Optional - Config für Überschreibungen
  
  Returns:
    - Für 'value': Named vector mit Code = Label (z.B. c('AO01' = 'Ja', 'AO02' = 'Nein'))
    - Für 'variable': String mit Variablen-Label
    - Für 'item': String mit Matrix-Item-Label (bereinigt, ohne Klammern)
  "
  
  if (!exists("global_codebook", envir = .GlobalEnv)) {
    warning("Globales Codebook nicht gefunden!")
    return(if (label_type == "value") NULL else var_name)
  }
  
  codebook <- get("global_codebook", envir = .GlobalEnv)
  
  if (!var_name %in% codebook$Variable) {
    return(if (label_type == "value") NULL else var_name)
  }
  
  # ============================================================================
  # WERTE-LABELS (für Kategorien in Tabellen)
  # ============================================================================
  if (label_type == "value") {
    # 1. Prüfe Config-Überschreibung
    if (!is.null(config)) {
      var_config <- config$variablen[config$variablen$variable_name == var_name, ]
      if (nrow(var_config) > 0 && !is.na(var_config$coding[1]) && var_config$coding[1] != "") {
        labels <- parse_coding(var_config$coding[1])
        if (!is.null(labels) && length(labels) > 0) {
          return(labels)
        }
      }
    }
    
    # 2. Hole aus Codebook
    wertelabels_str <- codebook$Wertelabels[codebook$Variable == var_name]
    if (!is.na(wertelabels_str) && wertelabels_str != "keine" && wertelabels_str != "") {
      labels <- parse_coding(wertelabels_str)
      return(labels)
    }
    
    return(NULL)
  }
  
  # ============================================================================
  # VARIABLEN-LABEL (für Spaltenüberschriften)
  # ============================================================================
  if (label_type == "variable") {
    var_label <- codebook$Label[codebook$Variable == var_name]
    
    if (!is.na(var_label) && var_label != "" && var_label != var_name) {
      # Kürze wenn zu lang
      if (nchar(var_label) > 150) {
        return(paste0(substr(var_label, 1, 147), "..."))
      }
      return(var_label)
    }
    
    return(var_name)
  }
  
  # ============================================================================
  # MATRIX-ITEM-LABEL (bereinigt, nur Text aus Klammern)
  # ============================================================================
  if (label_type == "item") {
    # 1. Prüfe Config item_labels Überschreibung
    if (!is.null(config) && "item_labels" %in% names(config$variablen)) {
      var_config <- config$variablen[config$variablen$variable_name == var_name, ]
      if (nrow(var_config) > 0) {
        item_labels_str <- var_config$item_labels[1]
        if (!is.na(item_labels_str) && item_labels_str != "") {
          # Parse und suche passenden Code
          item_labels <- parse_coding(item_labels_str)
          # Extrahiere Item-Code aus Variablennamen
          # (z.B. KW04_SQ001 -> SQ001)
          # Hier müsste matrix_name übergeben werden - vereinfachen wir später
        }
      }
    }
    
    # 2. Hole aus Codebook
    var_label <- codebook$Label[codebook$Variable == var_name]
    
    if (!is.na(var_label) && var_label != "" && var_label != var_name) {
      # Extrahiere Text aus eckigen Klammern
      bracket_match <- regexpr("\\[([^\\]]+)\\]", var_label)
      
      if (bracket_match > 0) {
        bracket_start <- bracket_match[1] + 1
        bracket_length <- attr(bracket_match, "match.length") - 2
        item_label <- substr(var_label, bracket_start, bracket_start + bracket_length - 1)
        
        # Prüfe ob nicht generisch
        if (nchar(item_label) > 0 && !grepl("^(Subquestion|Item|SQ)\\s*\\d+$", item_label)) {
          return(item_label)
        }
      }
      
      # Fallback: Verwende volles Label (gekürzt)
      if (nchar(var_label) < 150) {
        return(var_label)
      } else {
        return(paste0(substr(var_label, 1, 147), "..."))
      }
    }
    
    return(var_name)
  }
  
  # Unbekannter label_type
  return(var_name)
}

# Wrapper-Funktionen für Kompatibilität
get_value_labels_with_priority <- function(data, var_name, config = NULL) {
  get_labels(var_name, "value", config)
}

get_variable_label <- function(var_name, label_type = "full") {
  if (label_type == "item") {
    return(get_labels(var_name, "item"))
  } else {
    return(get_labels(var_name, "variable"))
  }
}

extract_item_label <- function(data, var_name, matrix_name, var_config = NULL, debug = FALSE) {
  get_labels(var_name, "item", if (!is.null(var_config)) list(variablen = var_config) else NULL)
}
