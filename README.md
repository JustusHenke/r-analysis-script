# Survey Datenauswertung - README

## Überblick

Dieses R-Skript führt eine automatisierte Auswertung von Umfragedaten durch, basierend auf einer konfigurierbaren Excel-Steuerungsdatei. Es erstellt deskriptive Statistiken, Kreuztabellen, Regressionsanalysen und verarbeitet offene Textantworten.

## Dateien und Struktur

```
├── 5 Analyse_Cockpit.R          # Hauptskript mit Konfiguration
├── __AnalysisFunctions.R        # Analysefunktionen 
├── 5 Analyse-Konfig.xlsx        # Konfigurationsdatei (EXCEL)
├── 00_Codebook.xlsx             # Codebook (optional)
├── 3 Erhebungsdaten/
│   └── survey_data_imputed.rds  # Ihre Umfragedaten
└── 4 Auswertung/
    └── Ergebnisse.xlsx          # Ausgabedatei
```

## Schnellstart

### 1. Daten vorbereiten
- Umfragedaten als `.rds`, `.csv` oder `.xlsx` speichern
- Pfad in `DATA_FILE` in `5 Analyse_Cockpit.R` anpassen

### 2. Konfiguration anpassen
- `5 Analyse-Konfig.xlsx` öffnen und ausfüllen (siehe unten)
- Pfade in `5 Analyse_Cockpit.R` prüfen

### 3. Analyse starten
```r
source("5 Analyse_Cockpit.R")
main()
```

## Konfiguration der Excel-Datei

Die Datei `5 Analyse-Konfig.xlsx` enthält mehrere Arbeitsblätter:

### Sheet 1: "Variablen" (PFLICHT)

Definiert alle zu analysierenden Variablen:

| Spalte | Beschreibung | Beispiel |
|--------|--------------|----------|
| `variable_name` | Exakter Variablenname im Datensatz | `SD01` |
| `question_text` | Beschreibung/Fragentext | "Geschlecht" |
| `data_type` | Datentyp (siehe unten) | `nominal_coded` |
| `coding` | Kodierung für Labels | `1=Weiblich;2=Männlich;3=Divers` |
| `min_value` | Minimum für numerische Vars | `1` |
| `max_value` | Maximum für numerische Vars | `5` |
| `reverse_coding` | Umkodierung (TRUE/FALSE) | `FALSE` |
| `use_NA` | Fehlende Werte einbeziehen | `FALSE` |

#### Datentypen:
- **`numeric`**: Kontinuierliche Zahlen (Alter, Noten)
- **`nominal_coded`**: Kategorien mit Codes (1=Ja, 2=Nein)
- **`nominal_text`**: Kategorien als Text ("Ja", "Nein")
- **`ordinal`**: Rangordnung (1=schlecht bis 5=gut)
- **`dichotom`**: Ja/Nein bzw. 0/1 Variablen
- **`matrix`**: Matrix-Fragen (ZS01[001], ZS01[002], ...)

#### Kodierung-Format:
```
1=Stimme gar nicht zu;2=Stimme eher nicht zu;3=Teils/teils;4=Stimme eher zu;5=Stimme voll zu
```

### Sheet 2: "Kreuztabellen" (OPTIONAL)

Definiert Kreuztabellen zwischen Variablen:

| Spalte | Beschreibung | Beispiel |
|--------|--------------|----------|
| `analysis_name` | Name der Analyse | "Geschlecht_x_Zufriedenheit" |
| `variable_1` | Erste Variable | `SD01` |
| `variable_2` | Zweite Variable | `GP01` |
| `statistical_test` | Test-Typ | `chi_square` |

#### Test-Typen:
- **`chi_square`**: Chi-Quadrat-Test (kategorisch × kategorisch)
- **`t_test`**: t-Test (numerisch × binär kategorisch)
- **`anova`**: ANOVA (numerisch × mehrkategorisch)
- **`correlation`**: Korrelation (numerisch × numerisch)
- **`mann_whitney`**: Mann-Whitney-U (ordinal × binär)

### Sheet 3: "Regressionen" (OPTIONAL)

Definiert Regressionsmodelle:

| Spalte | Beschreibung | Beispiel |
|--------|--------------|----------|
| `regression_name` | Name des Modells | "Zufriedenheit_Modell" |
| `dependent_var` | Abhängige Variable | `zufriedenheit_index` |
| `independent_vars` | Unabhängige Variablen | `SD01;SD04;AS07` |
| `regression_type` | Modell-Typ | `linear` |

#### Regressions-Typen:
- **`linear`**: Lineare Regression
- **`logistic`**: Logistische Regression
- **`ordinal`**: Ordinale Regression
- **`multilevel`**: Mehrebenen-Regression

#### Unabhängige Variablen:
- Mit `;` trennen: `SD01;SD04;AS07`
- Interaktionen mit `*`: `SD01*SD04;AS07`

### Sheet 4: "Textantworten" (OPTIONAL)

Verarbeitet offene Textantworten:

| Spalte | Beschreibung | Beispiel |
|--------|--------------|----------|
| `analysis_name` | Name der Analyse | "Verbesserungsvorschläge" |
| `text_variable` | Variable mit Textantworten | `GP05[other]` |
| `sort_variable` | Sortierung nach Variable | `SD01` |
| `min_length` | Mindest-Zeichenzahl | `3` |
| `include_empty` | Leere Antworten einschließen | `FALSE` |

## Matrix-Variablen

Für Matrix-Fragen (Likert-Skalen mit mehreren Items):

### Konfiguration:
```
variable_name: ZS01
data_type: matrix  
coding: 1=Stimme gar nicht zu;2=Stimme eher nicht zu;3=Teils/teils;4=Stimme eher zu;5=Stimme voll zu
```

### Das Skript erkennt automatisch:
- `ZS01[001]`, `ZS01[002]`, ... (LimeSurvey-Format)
- `ZS01.001.`, `ZS01.002.`, ... (R-sanitized)
- `ZS01_001`, `ZS01_002`, ... (alternative Formate)

### Dichotome Matrix (Checkbox-Grids):
Für Mehrfachauswahlmatrizen (1 = ausgewählt, leer = nicht ausgewählt):
```
data_type: matrix
coding: 1=Ausgewählt
```

## Hauptkonfiguration (5 Analyse_Cockpit.R)

### Datei-Pfade anpassen:
```r
CONFIG_FILE <- "5 Analyse-Konfig.xlsx"
DATA_FILE <- "3 Erhebungsdaten/survey_data_imputed.rds"
OUTPUT_FILE <- "4 Auswertung/Stipendiaten-Analyse-Ergebnisse.xlsx"
```

### Gewichtung:
```r
WEIGHTS <- TRUE          # Gewichtung aktivieren
WEIGHT_VAR <- "weight"   # Name der Gewichtungsvariable
```

### Weitere Einstellungen:
```r
ALPHA_LEVEL <- 0.05              # Signifikanzniveau
DIGITS_ROUND <- 2                # Rundung auf Dezimalstellen
INCLUDE_MISSING_DEFAULT <- FALSE # Fehlende Werte standardmäßig ausschließen
```

## Custom Variables

Das Skript kann automatisch zusätzliche Variablen erstellen. Diese werden in der Funktion `add_custom_vars()` definiert:

```r
add_custom_vars <- function(data) {
  data %>%
    mutate(
      # Neue Variable basierend auf existierenden
      geschlecht_maennl = case_when(
        SA02 == "Männlich" ~ 1,
        SA02 == "Weiblich" ~ 0,
        TRUE ~ NA_real_
      ),
      
      # Kategorisierung numerischer Variablen
      noten_kategorie = cut(
        AS07,
        breaks = c(-Inf, 1.5, 2.5, 3.5, 4.5, Inf),
        labels = c("sehr gut", "gut", "befriedigend", "ausreichend", "mangelhaft")
      )
    )
}
```

## Index-Bildung

Automatische Erstellung von Indizes aus Matrix-Items:

```r
index_definitions <- list(
  generate_index_definition(
    name = "zufriedenheit_index", 
    label = "Zufriedenheits-Index", 
    prefix = "GP01", 
    range = 1:5
  )
)
```

Dies erstellt automatisch einen Index aus `GP01[001]` bis `GP01[005]`.

## Ausgabe

Das Skript erstellt eine Excel-Datei mit folgenden Arbeitsblättern:

1. **Deskriptive_Statistiken**: Häufigkeiten, Mittelwerte, etc.
2. **Kreuztabellen**: Absolute und relative Häufigkeiten
3. **Statistische_Tests**: Testergebnisse (Chi², t-Test, etc.)
4. **Regressionsanalysen**: Koeffizienten und Modellgüte
5. **Textantworten**: Offene Antworten kategorisiert
6. **Variablen_Übersicht**: Überblick über alle Variablen

## Häufige Probleme & Lösungen

### Variablen nicht gefunden
- **Problem**: `WARNUNG: Variable SD01 nicht gefunden`
- **Lösung**: Überprüfen Sie Variablennamen in Ihren Daten. R konvertiert Sonderzeichen automatisch (z.B. `[` → `.`).

### Matrix-Items werden nicht erkannt
- **Problem**: Matrix-Frage zeigt keine Items
- **Lösung**: Prüfen Sie das Format Ihrer Matrix-Variablen. Unterstützt werden:
  - `ZS01[001]`, `ZS01[002]` (Original)
  - `ZS01.001.`, `ZS01.002.` (R-sanitized)

### Fehlende Labels
- **Problem**: Codes statt Labels in Ausgabe
- **Lösung**: Kodierung in Excel-Konfig hinzufügen:
  ```
  1=Stimme gar nicht zu;2=Stimme eher nicht zu;3=Teils/teils;4=Stimme eher zu;5=Stimme voll zu
  ```

### Regression schlägt fehl
- **Problem**: `FEHLER bei Regression: ...`
- **Lösungen**:
  - Prüfen Sie, ob alle Variablen existieren
  - Stellen Sie sicher, dass genügend vollständige Fälle vorhanden sind
  - Verwenden Sie korrekte Syntax für Interaktionen: `var1*var2`

### Gewichtung funktioniert nicht
- **Problem**: Gewichtete Analyse schlägt fehl
- **Lösung**: 
  - Prüfen Sie, ob `WEIGHT_VAR` in Ihren Daten existiert
  - Setzen Sie `WEIGHTS <- FALSE` für ungewichtete Analyse

## Advanced Features

### Mehrebenen-Regression
Für hierarchische Daten (z.B. Studierende in Hochschulen):
```r
regression_type: multilevel
```
Das Skript erkennt automatisch Clustering-Variablen wie `hochschul_id` oder `attribute_2`.

### [other]-Variablen
Für "Sonstiges"-Antworten:
```r
text_variable: AS03[other]
```
Das Skript findet automatisch entsprechende Textvariablen.

### Interaktionsterme
In Regressionen:
```r
independent_vars: geschlecht*bildungsfern;alter
```

## Support

Bei Problemen:
1. Prüfen Sie die R-Konsole auf Fehlermeldungen
2. Überprüfen Sie Ihre Excel-Konfiguration
3. Stellen Sie sicher, dass alle Variablennamen korrekt sind
4. Testen Sie mit einer kleineren Teilmenge Ihrer Daten

## Lizenz & Autor

Survey Analysis Script  
Autor: [Ihr Name]  
Datum: 2025  
Beschreibung: Automatisierte Auswertung von Survey-Daten basierend auf Excel-Konfiguration