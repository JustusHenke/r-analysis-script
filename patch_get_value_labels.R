# Patch für get_value_labels_with_priority
# Diese Funktion ersetzt die alte Version

get_value_labels_with_priority <- function(data, var_name, config = NULL) {
  "Extrahiert Value Labels mit Priorisierung: 1) Globales Codebook 2) Config-Labels"
  
  labels <- NULL
  
  # *** PRIORITÄT 1: Globales Codebook (Standard-Quelle) ***
  if (exists("global_codebook", envir = .GlobalEnv)) {
    codebook <- get("global_codebook", envir = .GlobalEnv)
    
    if (var_name %in% codebook$Variable) {
      wertelabels_str <- codebook$Wertelabels[codebook$Variable == var_name]
      
      if (!is.na(wertelabels_str) && wertelabels_str != "keine" && wertelabels_str != "") {
        # Parse Wertelabels (Format: "AO01 = Label1; AO02 = Label2")
        labels <- parse_coding(wertelabels_str)
        
        if (!is.null(labels) && length(labels) > 0) {
          cat("  ✓ Labels aus globalem Codebook für", var_name, ":", length(labels), "Labels\n")
          return(labels)
        }
      }
    }
  }
  
  # *** PRIORITÄT 2: Labels aus Config-Kodierung (Überschreibung) ***
  if (!is.null(config)) {
    var_config <- config$variablen[config$variablen$variable_name == var_name, ]
    
    if (nrow(var_config) > 0 && !is.na(var_config$coding[1]) && var_config$coding[1] != "") {
      labels <- parse_coding(var_config$coding[1])
      if (!is.null(labels) && length(labels) > 0) {
        cat("  ✓ Labels aus Config-Kodierung für", var_name, ":", length(labels), "Labels\n")
        return(labels)
      }
    }
  }
  
  # Keine Labels gefunden
  if (is.null(labels) || length(labels) == 0) {
    cat("  ⚠ Keine Labels gefunden für", var_name, "\n")
  }
  
  return(labels)
}
