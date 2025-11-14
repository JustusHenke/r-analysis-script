# CRUSH.md - R Survey Analysis System

## Project Overview

This is an **R-based automated survey data analysis system** designed for academic/research contexts. It reads survey data (from SPSS, CSV, or RDS files), applies configuration from an Excel file, and generates comprehensive statistical analyses including descriptive statistics, cross-tabulations, regressions, and text response processing.

**Language**: All code comments, variable names, and console output are in **German**. Documentation (README) is also in German.

**Key Characteristics**:
- Declarative configuration via Excel (no coding required for basic analyses)
- Supports weighted analyses
- Handles complex LimeSurvey matrix questions
- Preserves SPSS/haven labels from RDS files
- Generates formatted Excel output with multiple sheets

---

## Project Structure

```
.
├── Analysis-Cockpit.R          # Main entry point - configuration file
├── __AnalysisFunctions.R       # Core analysis functions (3000+ lines)
├── Prepare-SPSS-Data.R         # Helper script to convert SPSS → RDS
├── Analysis-Config.xlsx        # User configuration (Excel-based)
├── README.md                   # Comprehensive German documentation
└── LICENSE                     # MIT License
```

---

## Essential Commands

### Running the Analysis

**Interactive Mode (RStudio)**:
```r
# Set working directory to script location
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

# Run the main analysis
source("Analysis-Cockpit.R")
# The script auto-executes main() if run interactively
```

**Manual Execution**:
```r
source("Analysis-Cockpit.R")
results <- main()
```

### Preparing SPSS Data

```r
# Convert SPSS .sav files to RDS (preserves labels)
source("Prepare-SPSS-Data.R")

# Read SPSS file
data <- read_sav("Rohdaten/DATEINAME.sav")

# Save as RDS
saveRDS(data, "Rohdaten/DATEINAME.rds")
```

### No Test/Build Commands
- This is a script-based project with no formal test suite
- No package management files (DESCRIPTION, NAMESPACE)
- No CI/CD configuration
- Manual testing by running the analysis on real survey data

---

## Configuration Architecture

### Analysis-Cockpit.R (Main Configuration)

This file contains all user-configurable settings:

```r
# File paths
CONFIG_FILE <- "Analysis-Config.xlsx"
DATA_FILE <- "PATHTO/DATAFILE.rds"         # or .csv, .xlsx
OUTPUT_FILE <- "PATHTO/ANALYSISFILE.xlsx"
FINAL_DATASET_FILE <- "PATHTO/ANALYZEDDATAFILE.rds"

# Weighting
WEIGHTS <- FALSE                            # Enable/disable weighting
WEIGHT_VAR <- "weight"                      # Weight variable name

# Analysis parameters
ALPHA_LEVEL <- 0.05                         # Significance level
DIGITS_ROUND <- 2                           # Decimal places
INCLUDE_MISSING_DEFAULT <- FALSE            # Include missing values by default
```

**Custom Variables** (`add_custom_vars` function):
- Define derived variables using dplyr syntax
- Runs AFTER index creation, BEFORE config update
- Example: Create binary gender variable, categorize continuous vars

**Index Definitions** (`index_definitions` list):
- Automatically create composite indices from matrix items
- Uses `generate_index_definition()` helper
- Supports binary matrices (checkbox grids)

**Custom Labels**:
- `custom_var_labels`: Variable descriptions
- `custom_val_labels`: Value labels for categorical variables

### Analysis-Config.xlsx (Excel Configuration)

**Sheet 1: "Variablen" (REQUIRED)**

Defines all variables to analyze:

| Column | Description | Example |
|--------|-------------|---------|
| `variable_name` | Exact variable name in dataset | `SD01` |
| `question_text` | Question label/description | "Geschlecht" |
| `data_type` | Data type (see below) | `nominal_coded` |
| `coding` | Value labels (semicolon-separated) | `1=Weiblich;2=Männlich;3=Divers` |
| `min_value` | Minimum value (numeric vars) | `1` |
| `max_value` | Maximum value (numeric vars) | `5` |
| `reverse_coding` | Reverse code this variable? | `FALSE` |
| `use_NA` | Include missing values in output | `FALSE` |

**Data Types**:
- `numeric`: Continuous numbers (age, scores)
- `nominal_coded`: Categorical with numeric codes (1=Yes, 2=No)
- `nominal_text`: Categorical with text values ("Yes", "No")
- `ordinal`: Ordered categories (1=poor to 5=excellent)
- `dichotom`: Binary yes/no or 0/1 variables
- `matrix`: Matrix questions (ZS01[001], ZS01[002], ...)

**Sheet 2: "Kreuztabellen" (OPTIONAL)**

Cross-tabulation definitions:

| Column | Description | Example |
|--------|-------------|---------|
| `analysis_name` | Name for this analysis | "Geschlecht_x_Zufriedenheit" |
| `variable_1` | First variable | `SD01` |
| `variable_2` | Second variable | `GP01` |
| `statistical_test` | Test type | `chi_square` |

**Test Types**: `chi_square`, `t_test`, `anova`, `correlation`, `mann_whitney`

**Sheet 3: "Regressionen" (OPTIONAL)**

Regression model definitions:

| Column | Description | Example |
|--------|-------------|---------|
| `regression_name` | Model name | "Zufriedenheit_Modell" |
| `dependent_var` | Outcome variable | `zufriedenheit_index` |
| `independent_vars` | Predictors (semicolon-separated) | `SD01;SD04;AS07` |
| `regression_type` | Model type | `linear` |

**Regression Types**: `linear`, `logistic`, `ordinal`, `multilevel`
**Interaction Terms**: Use `*` for interactions: `SD01*SD04;AS07`

**Sheet 4: "Textantworten" (OPTIONAL)**

Open-ended text response processing:

| Column | Description | Example |
|--------|-------------|---------|
| `analysis_name` | Analysis name | "Verbesserungsvorschläge" |
| `text_variable` | Variable with text responses | `GP05[other]` |
| `sort_variable` | Sort by this variable | `SD01` |
| `min_length` | Minimum character length | `3` |
| `include_empty` | Include empty responses | `FALSE` |

---

## Code Organization

### __AnalysisFunctions.R Structure

**Main sections** (3000+ lines):

1. **Logging Functions** (lines 11-54)
   - `setup_logging()`, `log_cat()`, `close_logging()`
   - Dual output to console and log file

2. **Package Loading** (lines 56-91)
   - Auto-installs missing packages: readxl, openxlsx, dplyr, tidyr, stringr, psych, survey, haven, labelled, lme4
   - `load_packages()` function

3. **Configuration Loading** (lines 112-217)
   - `load_config()`: Reads Excel sheets, validates structure
   - `validate_variable_config()`: Checks required columns and valid data types

4. **Matrix Variable Handling** (lines 219-1016)
   - `create_matrix_table()`: Complex logic for matrix questions
   - `extract_numeric_from_matrix_coding()`: Converts coded responses to numeric
   - `parse_coding()`: Parses coding strings with multiple formats
   - Supports: Original format `ZS01[001]`, sanitized `ZS01.001.`, underscore `ZS01_001`
   - Detects dichotomous matrices (checkboxes), ordinal matrices (Likert scales), numeric matrices

5. **Data Loading and Preparation** (lines 1867-1940)
   - `load_and_prepare_data()`: Main pipeline
   - **Order of operations**:
     1. Load data (RDS/CSV/Excel)
     2. Sanitize variable names (`make.names()`)
     3. Convert text NAs to real NAs
     4. Apply reverse coding
     5. Create survey indices (BEFORE custom variables)
     6. Create custom variables (can use indices)
     7. Update config for sanitized names
     8. Apply variable labels
     9. Auto-detect categories
     10. Remove metadata variables
     11. Set variable types

6. **Descriptive Statistics** (lines 1942-2351)
   - `create_descriptive_tables()`: Main dispatcher
   - Type-specific functions:
     - `create_numeric_table()`: Mean, median, quartiles, SD
     - `create_nominal_coded_table()`: Frequencies with labels
     - `create_nominal_text_table()`: Text category frequencies
     - `create_ordinal_table()`: Frequencies + numeric stats
     - `create_dichotom_table()`: Binary frequencies
     - `create_matrix_table()`: Complex matrix analysis

7. **Label Handling** (lines 2357-2596)
   - `get_value_labels_with_priority()`: **Central label extraction** with priority:
     1. RDS labels (from SPSS/haven) - checks 3 attribute types
     2. Config coding strings
     3. Raw codes as fallback
   - **Label reversal detection**: Detects if labels are stored backwards (value=label vs label=value)
   - **NEW: `get_matrix_labels()`**: Streamlined matrix label extraction (eliminates redundancy)
     - Combines RDS, matrix variable, and config fallback strategies
     - Single function replaces 3+ duplicate code blocks
   - **NEW: `map_response_labels()`**: Intelligent label-to-value mapping
     - Handles direct matches, AO-patterns, A-patterns, generic patterns
     - Centralizes mapping logic previously duplicated 3+ times
     - Verbose mode for debugging
   - `extract_item_label()`: Extracts labels for matrix sub-items

8. **Cross-Tabulations** (lines 2598-3200+)
   - `create_matrix_crosstab()`: Matrix × categorical variable
   - `create_labeled_factor()`: Apply labels to factors
   - Statistical test support: chi-square, t-test, ANOVA, correlation

9. **Utility Functions**
   - `sort_response_categories()`: Intelligent sorting of ordinal responses
   - `make_clean_colname()`: Sanitize column names for R
   - `convert_text_nas()`: Convert "N/A", "missing" to real NA
   - `apply_reverse_coding()`: Reverse Likert scales (e.g., 1→5, 5→1)

### Analysis-Cockpit.R Structure

**Simple 3-part structure**:
1. **Configuration Section** (lines 6-28): File paths, weights, parameters
2. **Custom Variables Section** (lines 30-151): User-defined transformations
3. **Execution Section** (lines 153-168): Sources `__AnalysisFunctions.R` and calls `main()`

---

## Naming Conventions and Patterns

### Variable Naming

**Config/Script Level**:
- `UPPER_SNAKE_CASE` for global constants: `CONFIG_FILE`, `WEIGHT_VAR`, `ALPHA_LEVEL`
- `snake_case` for functions: `load_config()`, `create_matrix_table()`
- `snake_case` for local variables: `var_name`, `data_filtered`, `freq_table`

**German Variable Names**:
- Survey data often uses German codes: `SD01` (Soziodemographie), `GP01` (Gesamtprojekt), `AS07` (Allgemein-Studie)
- Output columns in German: `Haeufigkeit_absolut`, `Mittelwert`, `Median`, `Standardabweichung`

### Function Patterns

**Naming Pattern**: `<verb>_<noun>_<modifier>`
- `create_matrix_table()`
- `get_value_labels_with_priority()`
- `extract_numeric_from_matrix_coding()`

**Return Pattern**: Functions return **lists** with standardized structure:
```r
list(
  table = result_dataframe,
  variable = var_name,
  question = question_text,
  type = "numeric" | "nominal_coded" | "matrix" | etc.,
  weighted = TRUE/FALSE
)
```

**Matrix functions** return extended structure:
```r
list(
  table_categorical = ...,
  table_numeric = ...,        # For ordinal/binary matrices
  type = "matrix_dichotomous" | "matrix_ordinal" | "matrix_numeric",
  matrix_items = c(...),
  response_categories = c(...),
  response_labels = c(...)
)
```

### Code Style Conventions

**Indentation**: 2 spaces (standard R)

**Line Length**: Generally respects ~80-100 characters, but some complex lines exceed this

**Comments**:
- Section headers with `# =======` lines (73 characters)
- Function purposes documented with strings: `"Does something"`
- Inline comments explain complex logic
- German text in comments for domain-specific terms

**dplyr Pipelines**:
```r
data %>%
  filter(!is.na(variable)) %>%
  mutate(new_var = old_var * 2) %>%
  select(relevant_cols)
```

**Error Handling**:
- Heavy use of `tryCatch()` for robust execution
- `cat()` for warnings/debug output
- `stop()` for fatal errors (missing required sheets)

---

## Matrix Variables - Critical Pattern

**Matrix questions** are the most complex data type. Understanding this pattern is crucial:

### Format Recognition

The system detects matrix items using multiple patterns:
```r
# Original LimeSurvey format
ZS01[001], ZS01[002], ZS01[003]

# R-sanitized format (brackets replaced)
ZS01.001., ZS01.002., ZS01.003.

# Alternative formats
ZS01_001, ZS01-001
```

**Pattern matching** in `create_matrix_table()`:
```r
matrix_patterns <- c(
  paste0("^", matrix_name, "\\[.+\\]$"),     # Original
  paste0("^", matrix_name, "\\..+\\.$"),     # Sanitized
  paste0("^", matrix_name, "_.+$"),          # Underscore
  paste0("^", matrix_name, "-.+$")           # Dash
)
```

### Matrix Types

The system auto-detects 3 matrix types:

1. **Dichotomous Matrix** (checkbox grids):
   - Data: Only "1" (selected) or empty/NA (not selected)
   - Coding: `1=Ausgewählt` or no coding
   - Output: "Ausgewählt/Nicht ausgewählt" columns with counts and percentages

2. **Ordinal Matrix** (Likert scales):
   - Data: Numeric codes (1-5) or labeled text ("1 (stimme gar nicht zu)")
   - Coding: `1=Stimme gar nicht zu;2=Stimme eher nicht zu;...`
   - Output: Categorical table + numeric statistics (mean, median, SD)

3. **Numeric Matrix**:
   - Data: Numeric values without categorical meaning
   - No coding or minimal coding
   - Output: Numeric statistics only

### Label Parsing - Streamlined Architecture (NEW)

**Problem Solved**: Previously, label extraction logic was duplicated in 3+ locations with 100+ lines of redundant code for:
- Matrix table creation
- Matrix numeric statistics
- Matrix cross-tabulations

**Solution**: Two new helper functions eliminate redundancy:

### 1. `get_matrix_labels()` - Unified Label Extraction

Centralized function for retrieving labels for matrix variables:

```r
labels <- get_matrix_labels(
  data,              # Dataset
  matrix_vars,       # Vector of matrix item names
  matrix_name,       # Base matrix name (e.g., "ZS01")
  var_config,        # Variable config row (optional)
  matrix_coding      # Coding string from config (optional)
)
```

**Fallback Strategy** (tried in order):
1. RDS labels from first matrix item (`attr(data[[matrix_vars[1]]], "labels")`)
2. RDS labels from matrix variable itself (if exists in data)
3. Parse config coding string
4. Extract from var_config if provided

**Benefits**:
- Single source of truth for matrix label extraction
- Consistent behavior across all analysis types
- Reduces code duplication by ~70 lines per usage

### 2. `map_response_labels()` - Intelligent Pattern Matching

Maps raw response values to human-readable labels with pattern recognition:

```r
response_labels <- map_response_labels(
  unique_responses,   # Vector of raw values: c("AO01", "AO02", "1", "2")
  labels,            # Label mapping: c("1" = "Label1", "2" = "Label2")
  verbose = TRUE     # Print debug output
)
```

**Pattern Recognition** (tried in order for each response):
1. **Direct match**: `"1"` matches label key `"1"`
2. **AO-pattern**: `"AO01"` tries → `"AO01"`, `"AO1"`, `"1"`, `"01"`
3. **A-pattern**: `"A1"` tries → `"A1"`, `"1"`
4. **Generic prefix**: `"XYZ123"` tries → `"123"`

**Example**:
```r
# Data has: "AO01", "AO02", "AO03"
# Labels: c("1" = "Sehr gut", "2" = "Gut", "3" = "Befriedigend")
# Result: c("AO01" = "Sehr gut", "AO02" = "Gut", "AO03" = "Befriedigend")
```

**Benefits**:
- Eliminates 40-60 lines of duplicate mapping code per usage
- Consistent pattern matching across all analysis functions
- Debug output helps troubleshoot label issues

### Usage in Analysis Functions

**Before** (duplicated in 3+ places):
```r
# 50+ lines of label extraction
labels <- NULL
labels <- get_value_labels_with_priority(data, matrix_vars[1], ...)
if (is.null(labels)) labels <- parse_coding(...)
if (is.null(labels)) ...

# 40+ lines of response mapping
for (response in unique_responses) {
  if (response %in% names(labels)) { ... }
  if (grepl("^AO\\d+$", response)) { ... }
  if (grepl("^A\\d+$", response)) { ... }
  ...
}
```

**After** (2 lines):
```r
labels <- get_matrix_labels(data, matrix_vars, matrix_name, var_config, matrix_coding)
response_labels <- map_response_labels(unique_responses, labels, verbose = TRUE)
```

### Where Applied

1. **`create_matrix_table()`** (line ~487): Main matrix analysis
2. **`create_matrix_table()`** (line ~636): Numeric statistics section  
3. **`create_matrix_categorical_crosstab()`** (line ~3015): Cross-tabulation

**Code Reduction**: ~250 lines eliminated, functionality preserved.

---

For matrix items, labels are retrieved with strict priority:

1. **Variable label attribute** (`attr(data[[var]], "label")`) - from SPSS/RDS
2. **Custom variable labels** (`custom_var_labels`)
3. **Labelled package labels** (`labelled::var_label()`)
4. **Intelligent extraction** from variable name (strips matrix prefix)
5. **Fallback**: Formatted variable name

### Critical Gotcha: Label Reversal

RDS files from SPSS may have **reversed label structures**:

```r
# Normal (expected):
c("1" = "Stimme zu", "2" = "Stimme nicht zu")

# Reversed (SPSS export issue):
c("Stimme zu" = "1", "Stimme nicht zu" = "2")
```

**Solution** in `get_value_labels_with_priority()`:
```r
# Detect reversal by comparing average string lengths
avg_value_len <- mean(nchar(values))
avg_name_len <- mean(nchar(names_vals))

if (avg_value_len < avg_name_len && avg_value_len <= 10) {
  # Reverse: names become values, values become names
  labels <- setNames(names_vals, values)
}
```

---

## Variable Name Sanitization

**Critical Pattern**: R sanitizes variable names on import, affecting config matching.

### Sanitization Rules

```r
# make.names() converts:
"ZS01[001]"  →  "ZS01.001."
"AS03[other]"  →  "AS03.other."
"GP-05"  →  "GP.05"
```

### Config Update Pattern

After loading data, `update_config_variable_names()` updates ALL config references:

1. **Variablen sheet**: Direct variable names
2. **Kreuztabellen sheet**: Both variable_1 and variable_2
3. **Regressionen sheet**: dependent_var and independent_vars (split by `;`)
4. **Textantworten sheet**: text_variable and sort_variable

**Interaction terms** are handled specially:
```r
# "geschlecht*bildungsfern" → both parts sanitized separately
"geschlecht*bildungsfern"  →  "geschlecht*bildungsfern"  (if both exist)
"SD01*AS03"  →  "SD01*AS03"  (sanitized names matched individually)
```

### [other] Variables

LimeSurvey "other" responses require special handling:

```r
# Config: AS03[other]
# In data: AS03.other.

# Pattern matching in find_other_variable_simple():
patterns <- paste0("^", base_var_sanitized, c("\\.other\\.", "_other$", "\\.other$"))
```

---

## Index Creation Pattern

**Indices** are composite scores from multiple matrix items.

### Definition Syntax

```r
index_definitions <- list(
  generate_index_definition(
    name = "zufriedenheit_index",           # Variable name for new index
    label = "Zufriedenheits-Index",         # Human-readable label
    prefix = "ZF01",                        # Matrix variable prefix
    range = 1:5                             # Items: ZF01[001] to ZF01[005]
  ),
  
  # Binary matrices (checkbox grids)
  list(
    name = "netzwerk_nutzung_index",
    label = "Netzwerk-Nutzungs-Index",
    vars_original = paste0("NW01[", sprintf("%03d", 1:20), "]"),
    binary = TRUE                           # Critical flag
  )
)
```

### Index Calculation

**Standard indices** (ordinal):
```r
# Extracts numeric values (1-5), computes rowMeans()
# Missing values handled gracefully (requires ≥1 valid value per row)
```

**Binary indices** (checkboxes):
```r
# Converts: empty/NA → 0, "1"/Y → 1
# Computes proportion selected (mean of 0/1 values)
```

### Critical Order

Indices are created **BEFORE** custom variables:
```r
# In load_and_prepare_data():
index_result <- create_survey_indices(data, config, index_definitions)
data <- index_result$data
config <- index_result$config

# THEN custom variables can reference indices
data <- add_custom_vars(data)
```

---

## Weighting Pattern

If `WEIGHTS = TRUE`:

1. **Survey object creation** (`create_survey_object()`):
   - Converts all factors to character (survey package requirement)
   - Creates `svydesign()` object with weights

2. **Weighted statistics** used throughout:
   - `svymean()`, `svyquantile()`, `svyvar()` for numeric vars
   - `svytable()` for frequencies
   - Fallback to unweighted if errors occur

3. **Result tracking**: All outputs include `weighted = TRUE/FALSE` flag

---

## Output Structure

The system generates an **Excel workbook** with multiple sheets:

1. **Deskriptive_Statistiken**: All variable summaries
2. **Kreuztabellen**: Cross-tabulations (absolute + relative frequencies)
3. **Statistische_Tests**: Test results (chi², t-test, ANOVA, etc.)
4. **Regressionsanalysen**: Regression coefficients and model fit
5. **Textantworten**: Processed open-ended responses
6. **Variablen_Übersicht**: Variable metadata overview

**Sheet creation** uses `openxlsx` package with formatting:
- Headers styled (bold, background color)
- Numeric columns formatted appropriately
- Auto-column-width

---

## Common Pitfalls and Solutions

### 1. Variable Not Found Warning

**Symptom**: `WARNUNG: Variable SD01 nicht gefunden`

**Cause**: R sanitizes variable names on import (brackets → dots)

**Solution**: System auto-updates config, but check:
- Original name: `SD01[001]`
- Sanitized name: `SD01.001.`
- Config references are updated automatically

### 2. Matrix Items Not Recognized

**Symptom**: `WARNUNG: Keine Matrix-Items gefunden für ZS01`

**Cause**: Pattern matching fails due to unexpected format

**Debug**:
```r
# Check actual variable names in data
names(data)[grepl("ZS01", names(data))]

# Verify matrix pattern matches one of:
# ZS01[001], ZS01.001., ZS01_001, ZS01-001
```

**Solution**: Ensure matrix_name in config matches the prefix exactly

### 3. Missing Labels in Output

**Symptom**: Numeric codes displayed instead of text labels

**Cause**: Labels not found in RDS attributes or config coding

**Solution**:
1. Check if RDS file has labels: `attr(data$SD01, "labels")`
2. Add coding to config: `1=Label1;2=Label2;...`
3. Ensure label priority system runs (check console output)

### 4. Regression Fails

**Symptom**: `FEHLER bei Regression: ...`

**Causes**:
- Missing variables (check sanitized names)
- Insufficient complete cases
- Interaction term syntax error

**Solution**:
- Verify all variables exist: `var_name %in% names(data)`
- Check complete cases: `sum(complete.cases(data[, c(dep_var, indep_vars)]))`
- Interaction syntax: `var1*var2` (asterisk, no spaces)

### 5. Weighted Analysis Fails

**Symptom**: Fallback to unweighted statistics

**Cause**: Weight variable missing or has issues

**Solution**:
- Verify weight variable exists: `WEIGHT_VAR %in% names(data)`
- Check for NA/negative weights: `summary(data[[WEIGHT_VAR]])`
- Set `WEIGHTS <- FALSE` to disable weighting

### 6. Dichotomous Matrix Showing Wrong N

**Symptom**: Total N is smaller than expected for checkbox grid

**Cause**: Old logic counted only rows with ≥1 selection

**Fix Applied**: Now uses full sample N (`nrow(data)`) for dichotomous matrices:
```r
if (is_dichotomous_matrix) {
  total_n <- nrow(data)  # Full sample, not filtered
  count_1 <- sum(var_data == "1", na.rm = TRUE)
  count_0_or_empty <- total_n - count_1
}
```

---

## Known Issues Fixed

### Vector Indexing Bug in Matrix Tables (FIXED - Commit ed2fae3)

**Issue**: Dichotomous matrix tables (Y/N, 1/0 responses) failed with "Ersetzung hat 2 Zeilen, Daten haben 1"

**Root Cause**: Unsafe named vector indexing in loops caused R's vector recycling to return multiple values

**Fix**: Use explicit `which()` indexing with safety checks (line 665-676 in __AnalysisFunctions.R)

**Details**: See `BUG_ANALYSIS_AND_FIX.md` for complete technical analysis

---

## Testing Approach

**No formal test suite**, but manual testing workflow:

1. **Sample Data Setup**:
   - Use real survey data (anonymized)
   - Create small test config Excel file
   - Test various data types (numeric, nominal, ordinal, matrix)

2. **Execution**:
   ```r
   source("Analysis-Cockpit.R")
   # Check console output for warnings/errors
   ```

3. **Output Validation**:
   - Open generated Excel file
   - Verify frequencies match manual counts
   - Check labels are correct
   - Validate cross-tabulation totals

4. **Edge Cases**:
   - Empty variables (all NA)
   - Single-value variables
   - Large matrices (20+ items)
   - Missing data handling

---

## Package Dependencies

**Required packages** (auto-installed):

- **readxl**: Read Excel config file
- **openxlsx**: Write formatted Excel output
- **dplyr**: Data manipulation
- **tidyr**: Data reshaping
- **stringr**: String operations
- **psych**: Descriptive statistics
- **survey**: Weighted analyses
- **haven**: Read SPSS/Stata files
- **labelled**: Handle labeled data
- **lme4**: Multilevel models (regressions)

**Optional dependencies**:
- **rstudioapi**: Auto-set working directory (if running in RStudio)

---

## Git Workflow

**Current state**:
- Main branch only
- Recent commits show iterative development on `__AnalysisFunctions.R`
- No feature branches or pull requests in history

**Commit pattern**: Direct commits to main with descriptive messages referencing file names

---

## Key Design Principles

1. **Declarative Configuration**: Non-programmers define analyses via Excel, not code

2. **Label Preservation**: SPSS labels from RDS files take priority over config

3. **Robust Error Handling**: Analyses continue even if individual variables fail

4. **German Domain Language**: All output, messages, and column names in German (academic context)

5. **Flexible Variable Naming**: Handles multiple sanitization formats for matrix variables

6. **Dual Output**: Console messages + log file for debugging

7. **Composite Indices**: Automatic creation of mean-based indices from matrix items

8. **Weighted Analysis Optional**: Can be toggled without changing config

---

## Memory Files

No .cursorrules, claude.md, agents.md, or similar files found. This CRUSH.md serves as the primary documentation for AI agents.

---

## Working with This Project

### Adding a New Variable Type

1. Add to valid_types in `validate_variable_config()`: (__AnalysisFunctions.R:196)
2. Create `create_<type>_table()` function following pattern: (__AnalysisFunctions.R:2000+)
3. Add case to switch statement in `create_descriptive_tables()`: (__AnalysisFunctions.R:1978)
4. Update README.md with documentation

### Adding a New Statistical Test

1. Add test type to Kreuztabellen documentation (README)
2. Implement test logic in cross-tabulation functions (__AnalysisFunctions.R:2600+)
3. Add result formatting for Excel output

### Modifying Index Creation Logic

**Location**: `create_survey_indices()` (__AnalysisFunctions.R:1688)

**Pattern**:
```r
# Extract sub-data
subdata <- data[vars_present]

# Convert to numeric (ordinal extraction or binary conversion)
for (var in names(subdata)) {
  if (def$binary) {
    # Binary: empty → 0, "1" → 1
  } else {
    # Ordinal: extract numeric from coded text
  }
}

# Calculate index (mean)
index_vec <- create_numeric_index_safe(subdata, label)

# Add to data and config
data[[name]] <- index_vec
config <- add_index_to_config(config, name, label, vars_present)
```

### Debugging Tips

**Enable debug output**: Uncomment cat() statements:
```r
# cat("DEBUG: Parsing coding:", coding_string, "\n")
# cat("DEBUG: Gefundene Labels:", paste(names(labels), "=", labels), "\n")
```

**Check intermediate data**:
```r
# After loading data
str(data)
summary(data)

# Check sanitized names
names(data)

# Verify config update
config$variablen$variable_name
```

**Trace function execution**:
```r
# Add at function start
cat("Entering function:", sys.call(), "\n")
```

---

## When to Edit Each File

### Edit Analysis-Cockpit.R if:
- Changing data file paths
- Enabling/disabling weighting
- Adding custom calculated variables
- Defining new composite indices
- Changing output settings (decimal places, significance level)

### Edit Analysis-Config.xlsx if:
- Adding/removing variables from analysis
- Changing variable labels or coding
- Defining new cross-tabulations
- Adding regression models
- Processing text responses

### Edit __AnalysisFunctions.R if:
- Fixing bugs in analysis logic
- Adding new statistical tests
- Modifying matrix handling logic
- Changing label extraction priority
- Improving error handling
- Adding new data type support

### Edit Prepare-SPSS-Data.R if:
- Changing SPSS import logic
- Modifying label preservation
- Adding data preprocessing steps

### Edit README.md if:
- Documenting new features
- Adding troubleshooting guides
- Updating configuration examples

---

## Special Patterns to Preserve

### 1. Survey Object Creation

**Always convert factors to character** before creating survey object:
```r
# CRITICAL: survey package requires character, not factor
factor_vars <- sapply(survey_data, is.factor)
if (any(factor_vars)) {
  for (var_name in names(survey_data)[factor_vars]) {
    survey_data[[var_name]] <- as.character(survey_data[[var_name]])
  }
}
survey_obj <- svydesign(ids = ~1, weights = survey_data[[weight_var]], data = survey_data)
```

### 2. Label Attribute Preservation

**During any transformation, preserve label attributes**:
```r
# Save labels
original_labels <- attr(data[[var_name]], "labels")
original_label <- attr(data[[var_name]], "label")

# Transform data
data[[var_name]] <- some_transformation(data[[var_name]])

# Restore labels
if (!is.null(original_labels)) {
  attr(data[[var_name]], "labels") <- original_labels
}
if (!is.null(original_label)) {
  attr(data[[var_name]], "label") <- original_label
}
```

### 3. Matrix Item Iteration

**Always filter out [other] variables** when finding matrix items:
```r
matrix_vars <- matrix_vars[!grepl("other", matrix_vars, ignore.case = TRUE)]
```

### 4. Response Category Sorting

**Use intelligent sorting** for ordinal scales:
```r
# Don't use: sort(unique_responses)
# Instead use:
unique_responses <- sort_response_categories(unique_responses)
```

This recognizes common patterns like "Stimme gar nicht zu" → "Stimme voll zu"

---

## Final Notes

- **Language Barrier**: When modifying user-facing text, maintain German language
- **Excel Dependencies**: Config file structure is fixed - changing it breaks the system
- **No Type Checking**: R is dynamically typed; rely on defensive programming
- **Large Functions**: Some functions exceed 200 lines due to complex matrix logic
- **Global State**: Uses global variables (WEIGHTS, WEIGHT_VAR, etc.) defined in Analysis-Cockpit.R
- **Logging**: Extensive console output is intentional for user feedback

This system prioritizes **flexibility** (Excel-driven config) and **robustness** (handles varied survey formats) over code elegance.
