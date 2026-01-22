# Debug-Script um zu prüfen welche Labels für KW03 vorhanden sind

# Lade die Daten
data <- readRDS("data.rds")

# Finde alle KW03 Variablen
kw03_vars <- names(data)[grepl("^KW03", names(data))]

cat("Gefundene KW03 Variablen:\n")
print(kw03_vars)
cat("\n")

# Prüfe Labels für jede Variable
for (var in kw03_vars) {
  cat("========================================\n")
  cat("Variable:", var, "\n")
  cat("----------------------------------------\n")
  
  # Prüfe alle Attribute
  all_attrs <- attributes(data[[var]])
  cat("Alle Attribute:\n")
  print(names(all_attrs))
  cat("\n")
  
  # Label Attribut
  if ("label" %in% names(all_attrs)) {
    cat("Label Attribut:\n")
    print(all_attrs$label)
    cat("\n")
  }
  
  # Labelled Package
  if (requireNamespace("labelled", quietly = TRUE)) {
    if (labelled::is.labelled(data[[var]])) {
      var_label <- labelled::var_label(data[[var]])
      cat("Labelled Package Label:\n")
      print(var_label)
      cat("\n")
    }
  }
  
  # Erste paar Werte
  cat("Erste Werte:\n")
  print(head(data[[var]], 10))
  cat("\n")
}
