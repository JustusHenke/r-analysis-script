# Matrix Crosstab Statistical Tests - Implementation

## Summary

Statistical tests for matrix crosstabs have been successfully implemented. Instead of showing "Statistische Tests für Matrix-Variablen nicht unterstützt", the system now runs statistical tests for each individual matrix item.

## Changes Made

### 1. Test Execution in `create_crosstabs()` (lines 4166-4194)

**Before:** Matrix crosstabs returned dummy test result saying tests are not supported.

**After:** 
- Detects matrix crosstabs via `"matrix_items" %in% names(crosstab_result)`
- Identifies group variable (either var1 or var2, depending on which is the matrix)
- Iterates through all matrix items (`matrix_items <- crosstab_result$matrix_items`)
- Runs `perform_statistical_test()` for each item vs. group variable
- Stores individual test results in `item_tests` list
- Returns unified test_result structure with all item tests

### 2. Return Structure

Matrix crosstab test result now includes:
```r
list(
  test = "Matrix-Kreuztabelle - chi_square Test",  # or other test type
  result = "Tests für N Matrix-Items durchgeführt",
  p_value = NA,
  statistic = NA,
  item_tests = list(
    "FP01.001." = list(test="Chi-Quadrat", statistic=..., p_value=...),
    "FP01.002." = list(test="Chi-Quadrat", statistic=..., p_value=...),
    ...
  )
)
```

### 3. Excel Export in `export_statistical_tests()` (lines 6414-6524)

**Test Summary Table (lines 6414-6450):**
- For normal crosstabs: Single row per analysis
- For matrix crosstabs: Multiple rows, one per item
- Item name appended in brackets: "FP01a_x_F2 [FP01.001.]"

**Detailed Test Results (lines 6478-6523):**
- **Matrix Tests:** Shows tabular summary of all matrix item tests
  - Columns: Item, Test, Statistik, p_Wert, Ergebnis
  - Formatted with openxlsx styling
- **Normal Tests:** Shows original detailed format (Parameter-Wert pairs)

## Supported Test Types

All existing test types automatically work with matrix items:
- `chi_square` - Default for categorical matrix responses
- `t_test` - For numeric matrix items (when applicable)
- `anova` - For multiple group comparisons
- `correlation` - For numeric correlations
- `mann_whitney` - Non-parametric alternative

## Example Usage

In Analysis-Config.xlsx, Kreuztabellen sheet:
```
| analysis_name    | variable_1 | variable_2 | statistical_test |
|------------------|-----------|-----------|-----------------|
| FP01a_x_F2       | FP01       | F2        | chi_square       |
| FP01a_x_F1       | FP01       | F1        | chi_square       |
```

**Result in Excel export:**
- Statistische_Tests sheet now shows:
  - Übersicht aller Tests: Individual rows for each matrix item
  - Detaillierte Testergebnisse: Summary table for each matrix analysis

## Technical Details

### Matrix Item Detection
- Uses existing `is_matrix_variable()` function
- Extracts items from `crosstab_result$matrix_items`
- Identifies group variable from var1_is_matrix/var2_is_matrix flags

### Test Iteration
- Loopsthrough each matrix item
- Calls `perform_statistical_test(data, item, group_var, test_type, survey_obj, config)`
- Handles weighted (survey) and unweighted tests automatically
- Collects results in named list for organized output

### Excel Formatting
- Uses existing header_style and table_style
- Matrix test table: 5-column format (Item, Test, Statistik, p_Wert, Ergebnis)
- Maintains consistent look with other report sections

## Benefits

1. **Complete Statistical Analysis:** Matrix questions now get proper statistical tests
2. **Individual Item Assessment:** Each matrix item tested separately for relationships
3. **Comprehensive Reports:** Excel export shows both summary and detailed results
4. **No Configuration Changes:** Works with existing config - just enable desired test type
5. **Flexible:** Supports all test types (chi-square, t-test, ANOVA, correlation, Mann-Whitney)

## Files Modified

- `__AnalysisFunctions.R`: 
  - Lines 4166-4194: Test execution logic
  - Lines 6414-6524: Excel export logic
- `Analysis-Cockpit.R`: Minor formatting (1 change)

## Testing Recommendations

1. Test with dichotomous matrix (checkbox grid) + categorical grouping variable
   - Expected: Chi-square tests for each item
2. Test with ordinal matrix (Likert scale) + categorical grouping variable
   - Expected: Chi-square tests (or other appropriate test)
3. Test with numeric matrix + categorical grouping variable
   - Expected: t-test or ANOVA depending on groups
4. Verify weighted analysis works with matrix tests
5. Check Excel export formatting with 10+ matrix items

