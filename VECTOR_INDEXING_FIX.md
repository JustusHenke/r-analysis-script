# Matrix Table Response Label Indexing Fix

## Problem

When processing matrix tables with dichotomous data (Y/N or 1/0 responses), the code was failing with:
```
FEHLER bei Variable E2 : Ersetzung hat 2 Zeilen, Daten haben 1
```

This error indicates that **vector replacement was attempting to assign 2 values to 1 element**.

## Root Cause

In the normal (non-dichotomous) matrix categorical table creation (line 666), the code was using direct vector indexing:

```r
response_label <- response_labels[as.character(response)]
```

**Problem**: When `response_labels` is a named vector with limited elements, this indexing can return:
- Multiple values if the name matches multiple elements (shouldn't happen, but vector recycling can cause issues)
- Issues with partial matching in named vector subsetting

## Solution

Changed the indexing approach to use `which()` with explicit index extraction:

```r
# OLD (line 666):
response_label <- response_labels[as.character(response)]

# NEW (lines 666-676):
response_char <- as.character(response)
response_label <- NA_character_

# Versuche direkten Match in response_labels names
matching_idx <- which(names(response_labels) == response_char)
if (length(matching_idx) > 0) {
  response_label <- response_labels[matching_idx[1]]
} else {
  # Fallback zu rauem Wert
  response_label <- response_char
}
```

**Why this works**:
1. `which()` returns the index positions (0 or 1 in this case)
2. Explicit `[1]` ensures we take only the first match
3. Clear fallback to raw value if no match
4. No vector recycling or implicit coercion issues

## Why Matrix Crosstables Work But Matrix Tables Don't

**Matrix crosstables** (`create_matrix_categorical_crosstab`):
- Use simpler mapping with direct `if` statements (no indexed vector assignment)
- Don't have the complex column name building loop
- Mapping doesn't trigger vector recycling issues

**Matrix tables** (`create_matrix_table`):
- Build many columns dynamically in a loop
- Use indexed vector assignment in the loop
- R's vector recycling/coercion can cause issues when indexing returns unexpected lengths

## Files Changed

- `__AnalysisFunctions.R` (lines 665-676): Safe vector indexing in categorical table creation

## Testing

Manually test with:
1. Dichotomous matrices (Y/N responses)
2. Verify "Ja" and "Nicht Gewählt" labels appear correctly
3. Verify counts and percentages are correct
4. Check that multiple matrix items work

No automated tests available.
