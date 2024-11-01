# DSStatTools VBA Toolpack 📊

This VBA toolpack provides a range of statistical functions that can be used as Excel formulas to process and analyze data directly in your spreadsheet. To use this toolpack, import the `.bas` file into the VBA editor in Excel.

### Import Instructions

1. Open the **Visual Basic for Applications (VBA)** editor (`Alt + F11`).
2. Go to **File > Import File...** and select `DSStatTools_Basic.bas`.
3. The functions are now available to be used as Excel formulas.

---

## Module: DSStatTools_Basic 📈
This module includes functions for advanced data manipulation and statistical analysis, such as counting occurrences based on patterns, handling comma-separated lists, merging ranges, and calculating interquartile ranges (IQR).

### Functions

#### `DS_CountIfAny`

Counts cells within a range that match any of up to four specified patterns.

- **Usage:** `=DS_CountIfAny(A1:A10, "*Pattern*", "*AnotherPattern*")`
- **Parameters:**
  - `cellRange` (Range): The range of cells to search. **(Required)**
  - `comp1` (Variant): The first pattern to match. **(Required)**
  - `comp2`, `comp3`, `comp4` (Variant): Additional patterns to match. **(Optional)**

#### `DS_CommaListAsArray`

Converts a comma-separated list within cells into an array.

- **Usage:** `=DS_CommaListAsArray(A1:A10)`
- **Parameters:**
  - `cellRange` (Range): Range of cells containing comma-separated values. **(Required)**
  - `castNumber` (Variant): Converts values to numbers if specified. **(Optional)**

#### `DS_CSVStringAsArray`

Converts a single comma-separated string to an array.

- **Usage:** `=DS_CSVStringAsArray("value1,value2,value3")`
- **Parameters:**
  - `csvString` (String): The comma-separated string to convert. **(Required)**
  - `castNumber` (Variant): Converts values to numbers if specified. **(Optional)**

#### `DS_CommaListCountValue`

Counts occurrences of a specific value within a comma-separated list across cells.

- **Usage:** `=DS_CommaListCountValue(A1:A10, "searchTerm")`
- **Parameters:**
  - `cellRange` (Range): The range containing comma-separated values. **(Required)**
  - `comp` (Variant): The value to count within each cell's list. **(Required)**

#### `DS_CommaListCountEntries`

Counts the total entries in a comma-separated list within cells.

- **Usage:** `=DS_CommaListCountEntries(A1:A10)`
- **Parameters:**
  - `cellRange` (Range): The range containing comma-separated lists. **(Required)**

#### `DS_CommaListValueAt`

Returns the value at a specific index in a comma-separated list within cells.

- **Usage:** `=DS_CommaListValueAt(A1:A10, 2)`
- **Parameters:**
  - `cellRange` (Range): Range containing comma-separated values. **(Required)**
  - `index` (Integer): Index position (1-based) to retrieve the value. **(Required)**
  - `decimalSeparator` (Variant): Specifies the decimal separator if needed. **(Optional)**

#### `DS_RangeToArray`

Converts a range of cells to an array, excluding empty cells.

- **Usage:** `=DS_RangeToArray(A1:A10)`
- **Parameters:**
  - `cellRange` (Range): Range of cells to convert. **(Required)**

#### `DS_CommaListMake`

Combines a range of values into a single comma-separated string.

- **Usage:** `=DS_CommaListMake(A1:A10)`
- **Parameters:**
  - `cellRange` (Range): The range of values to combine. **(Required)**
  - `decimalSeparator` (Variant): Specifies a decimal separator if needed. **(Optional)**

#### `DS_PrintIQR`

Calculates and prints the Interquartile Range (IQR) for a set of values.

- **Usage:** `=DS_PrintIQR(A1:A10, 2)`
- **Parameters:**
  - `cellRange` (Range or Array): Range of cells or array of values for IQR calculation. **(Required)**
  - `decimals` (Variant): Number of decimal places for IQR values. **(Optional)**

#### `DS_MergeRangesToArray`

Merges multiple cell ranges into a single array, excluding empty cells.

- **Usage:** `=DS_MergeRangesToArray(A1:A10, B1:B10)`
- **Parameters:**
  - `cellRange1` (Range): The primary range to merge. **(Required)**
  - `cellRange2` to `cellRange10` (Range): Additional ranges to merge. **(Optional)**
 



## Module: DSStatTools_Helpers 🛠️
This module provides a variety of helper functions for data processing in Excel, including functions for converting strings to numbers, counting occurrences of substrings, checking numeric values, sorting arrays, and managing array operations.

### Functions

#### `DS_StringToNumber`

Converts a string to a number, handling optional decimal and thousands separators.

- **Usage:** `=DS_StringToNumber("1,234.56", ".", ",")`
- **Parameters:**
  - `value` (Variant): The value to convert. **(Required)**
  - `decimalSeparator` (Variant): Custom decimal separator, defaults to system setting. **(Optional)**
  - `thousandsSeparator` (Variant): Custom thousands separator, defaults to system setting. **(Optional)**

#### `DS_StringCountOccurrences`

Counts the occurrences of a specified substring within a text.

- **Usage:** `=DS_StringCountOccurrences("hello world", "o")`
- **Parameters:**
  - `strText` (String): The text to search within. **(Required)**
  - `strFind` (String): The substring to find. **(Required)**

#### `DS_IsNumeric`

Checks if a value is numeric, considering thousands and decimal separators.

- **Usage:** `=DS_IsNumeric("1,000.50")`
- **Parameters:**
  - `value` (Variant): The value to check for numeric status. **(Optional)**

#### `DS_JoinArrays`

Merges two arrays into one.

- **Usage:** `=DS_JoinArrays(array1, array2)`
- **Parameters:**
  - `array1` (Array): First array to join. **(Required)**
  - `array2` (Array): Second array to join. **(Required)**

#### `DS_AppendToArray`

Appends a value or array of values to an existing array.

- **Usage:** `=DS_AppendToArray(myArray, value)`
- **Parameters:**
  - `myArray` (Array): The array to which values will be appended. **(Required)**
  - `value` (Variant): Value or array of values to append. **(Required)**

#### `DS_InArray`

Checks if a value exists within an array.

- **Usage:** `=DS_InArray(myArray, "searchValue")`
- **Parameters:**
  - `myArray` (Array): The array to search. **(Required)**
  - `myValue` (Variant): The value to search for. **(Required)**

#### `DS_Occurrences`

Counts occurrences of one or more patterns within a cell range.

- **Usage:** `=DS_Occurrences(A1:A10, "pattern1", "pattern2")`
- **Parameters:**
  - `cellRange` (Range): Range of cells to search. **(Required)**
  - `comp` (Variant): First pattern to match. **(Required)**
  - `comp2`, `comp3`, `comp4` (Variant): Additional patterns to match. **(Optional)**

#### `DS_OccurrencesNot`

Counts cells that do not match specific patterns within a range.

- **Usage:** `=DS_OccurrencesNot(A1:A10, "pattern", "excludePattern")`
- **Parameters:**
  - `cellRange` (Range): Range of cells to evaluate. **(Required)**
  - `comp` (Variant): Pattern to match against. **(Required)**
  - `andNot2`, `andNot3`, `andNot4` (Variant): Additional patterns to exclude. **(Optional)**

#### `DS_PatternMatch`

Matches a value to a specified pattern, supporting numeric comparisons.

- **Usage:** `=DS_PatternMatch(A1, ">=10")`
- **Parameters:**
  - `val` (Variant): Value to match. **(Required)**
  - `comp` (Variant): Pattern or condition to evaluate. **(Required)**

#### `DS_Max`

Returns the maximum value within a range or array.

- **Usage:** `=DS_Max(A1:A10)`
- **Parameters:**
  - `cellRange` (Range): Range or array to find the maximum in. **(Required)**

#### `DS_Min`

Returns the minimum value within a range or array.

- **Usage:** `=DS_Min(A1:A10)`
- **Parameters:**
  - `cellRange` (Range): Range or array to find the minimum in. **(Required)**

#### `DS_ValueInArray`

Determines if a specified value is present within an array.

- **Usage:** `=DS_ValueInArray("needle", haystackArray)`
- **Parameters:**
  - `needle` (Variant): Value to search for. **(Required)**
  - `haystack` (Array): Array to search within. **(Required)**

#### `DS_OffsetValues`

Sets array values based on indexed positions.

- **Usage:** `=DS_OffsetValues(indexes, 5, "Val1", "Val2")`
- **Parameters:**
  - `indexes` (Range/Array): List of index values. **(Required)**
  - `reach` (Integer): Number of values to return. **(Required)**
  - `value1` (Variant): Value for matching indexes. **(Required)**
  - `value2` (Variant): Default value if no match. **(Required)**

#### `DS_FirstOrDefault`

Returns the first non-empty or non-zero value from a list of inputs.

- **Usage:** `=DS_FirstOrDefault(val1, val2, val3)`
- **Parameters:**
  - `val1`, `val2`, ..., `val10` (Variant): Values to evaluate in order. **(Optional)**




## Module: DSStatTools_ANOVA 📊
This module provides functions for performing non-parametric and parametric analysis of variance (ANOVA), allowing users to calculate p-values for one-way ANOVA and Kruskal-Wallis tests directly in Excel. These functions are useful for comparing groups to determine if there are statistically significant differences between them.

### Functions

#### `DS_KruskalWallisP`
Calculates the p-value for a Kruskal-Wallis H test, a non-parametric alternative to one-way ANOVA for comparing medians among multiple independent groups.

- **Usage:** `=DS_KruskalWallisP(A1:A10, B1:B10, C1:C10)`
- **Parameters:**
  - `valueRange1` (Range): First group of values to compare. **(Required)**
  - `valueRange2` (Range): Second group of values to compare. **(Required)**
  - `valueRange3` to `valueRange10` (Range): Additional groups to compare. **(Optional)**

#### `DS_ANOVAOneWayP`
Calculates the p-value for a one-way ANOVA, assessing whether there are significant differences in means across multiple independent groups.

- **Usage:** `=DS_ANOVAOneWayP(A1:A10, B1:B10, C1:C10)`
- **Parameters:**
  - `valueRange1` (Range): First group of values to compare. **(Required)**
  - `valueRange2` (Range): Second group of values to compare. **(Required)**
  - `valueRange3` to `valueRange10` (Range): Additional groups to compare. **(Optional)**




## Module: DSStatTools_ChiSq 🔍
This module provides tools for performing chi-square tests, including calculating the chi-square statistic, determining degrees of freedom, and computing the p-value. These functions are helpful for statistical analyses involving contingency tables.

### Functions

#### `DS_ChiSquare`

Calculates the chi-square test statistic for a contingency table.

- **Usage:** `=DS_ChiSquare(A1:B10)`
- **Parameters:**
  - `cellRange` (Range): The range of cells representing the contingency table. **(Required)**
- **Description:** This function computes the chi-square test statistic based on the observed and expected frequencies in the contingency table defined by `cellRange`.

#### `DS_ChiSquareDof`

Calculates the degrees of freedom for a contingency table used in a chi-square test.

- **Usage:** `=DS_ChiSquareDof(A1:B10)`
- **Parameters:**
  - `cellRange` (Range): The range of cells representing the contingency table. **(Required)**
- **Description:** Returns the degrees of freedom for the chi-square test based on the dimensions of the `cellRange` table. Degrees of freedom are calculated as `(rows - 1) * (columns - 1)`.

#### `DS_ChiSquareP`

Computes the p-value for the chi-square test, assessing the statistical significance of observed frequencies.

- **Usage:** `=DS_ChiSquareP(A1:B10)`
- **Parameters:**
  - `cellRange` (Range): The range of cells representing the contingency table. **(Required)**
- **Description:** Calculates the p-value of the chi-square test, indicating the probability that the observed distribution occurred by chance. A lower p-value suggests that the observed frequencies deviate significantly from the expected distribution.



## Module: DSStatTools_Correlation 🔗
This module provides a collection of functions for calculating various types of correlation coefficients and their statistical significance, including point-biserial, Spearman, and Pearson correlations. These functions are designed to be called directly from Excel formulas, making it easy for users to perform correlation analysis within Excel spreadsheets.

### Functions

#### `DS_Correlation_PointBiserialR`

Calculates the point-biserial correlation coefficient between a metric variable and a binary variable.

- **Usage:** `=DS_Correlation_PointBiserialR(A1:A10, B1:B10)`
- **Parameters:**
  - `metricRange` (Range): Range containing the metric (continuous) values. **(Required)**
  - `binaryRange` (Range): Range containing the binary (categorical) values. **(Required)**

#### `DS_Correlation_PointBiserialP`

Calculates the p-value associated with the point-biserial correlation coefficient.

- **Usage:** `=DS_Correlation_PointBiserialP(A1:A10, B1:B10)`
- **Parameters:**
  - `metricRange` (Range): Range containing the metric (continuous) values. **(Required)**
  - `binaryRange` (Range): Range containing the binary (categorical) values. **(Required)**

#### `DS_Correlation_SpearmanR`

Calculates the Spearman rank correlation coefficient between two variables.

- **Usage:** `=DS_Correlation_SpearmanR(A1:A10, B1:B10)`
- **Parameters:**
  - `cellRange1` (Range): First range of values. **(Required)**
  - `cellRange2` (Range): Second range of values. **(Required)**

#### `DS_Correlation_SpearmanP`

Calculates the p-value associated with the Spearman rank correlation coefficient.

- **Usage:** `=DS_Correlation_SpearmanP(A1:A10, B1:B10)`
- **Parameters:**
  - `cellRange1` (Range): First range of values. **(Required)**
  - `cellRange2` (Range): Second range of values. **(Required)**

#### `DS_Correlation_Spearman95CI`

Calculates the 95% confidence interval for the Spearman rank correlation coefficient.

- **Usage:** `=DS_Correlation_Spearman95CI(A1:A10, B1:B10, 3)`
- **Parameters:**
  - `cellRange1` (Range): First range of values. **(Required)**
  - `cellRange2` (Range): Second range of values. **(Required)**
  - `decimals` (Number): Number of decimal places for the confidence interval bounds. **(Optional, default = 2)**

#### `DS_Correlation_PearsonR`

Calculates the Pearson correlation coefficient between two variables.

- **Usage:** `=DS_Correlation_PearsonR(A1:A10, B1:B10)`
- **Parameters:**
  - `cellRange1` (Range): First range of values. **(Required)**
  - `cellRange2` (Range): Second range of values. **(Required)**

#### `DS_Correlation_PearsonP`

Calculates the p-value associated with the Pearson correlation coefficient.

- **Usage:** `=DS_Correlation_PearsonP(A1:A10, B1:B10)`
- **Parameters:**
  - `cellRange1` (Range): First range of values. **(Required)**
  - `cellRange2` (Range): Second range of values. **(Required)**

#### `DS_Correlation_Pearson95CI`

Calculates the 95% confidence interval for the Pearson correlation coefficient.

- **Usage:** `=DS_Correlation_Pearson95CI(A1:A10, B1:B10, 3)`
- **Parameters:**
  - `cellRange1` (Range): First range of values. **(Required)**
  - `cellRange2` (Range): Second range of values. **(Required)**
  - `decimals` (Number): Number of decimal places for the confidence interval bounds. **(Optional, default = 2)**




## Module: DSStatTools_ExactFisher 🔬
This module includes functions for calculating Fisher's exact test for a 2x2 contingency table, providing an exact probability value. It is useful in statistical analysis where sample sizes are small, and exact probabilities are needed rather than approximations.

### Functions

#### `DS_ExactFisher2x2P`

Calculates the exact p-value for a 2x2 contingency table using Fisher's exact test.

- **Usage:** `=DS_ExactFisher2x2P(A1:B2)`
- **Parameters:**
  - `cellRange` (Range): A 2x2 range representing the contingency table. The top-left and top-right cells represent the counts in the first row, and the bottom-left and bottom-right cells represent the counts in the second row. **(Required)**
  
  **Example**: If you have a 2x2 table of data in cells `A1:B2`, calling `=DS_ExactFisher2x2P(A1:B2)` will return the exact p-value for the Fisher test based on those values.





## Module: DSStatTools_Normal 📊
This module provides statistical functions that allow users to test the normality of a dataset using the Shapiro-Wilk test, suitable for ranges of data. It offers tailored calculations for smaller and larger data sets to maximize accuracy.

### Functions

#### `DS_ShapiroWilkP`

Calculates the p-value of the Shapiro-Wilk test for normality on a given range of data, adjusting for small and large sample sizes to provide a robust result. This function returns a p-value indicating the probability that the data follows a normal distribution.

- **Usage:** `=DS_ShapiroWilkP(A1:A10)`
- **Parameters:**
  - `cellRange` (Range): The range of cells containing numeric data to test for normality. **(Required)**

**Note:** Returns -1 if the sample size is outside the allowed range of 3 to 2000 data points.




## Module: DSStatTools_Query 🔍

This module provides functions for selecting data from a range based on specified conditions, including single and multiple conditions. It also includes functionality to retrieve unique values from a dataset. The functions are designed to facilitate advanced querying and filtering of data in Excel.

### Functions

#### `DS_Select`

Selects values from `cellRange` that meet a specified condition in `conditionRange`.

- **Usage:** `=DS_Select(A1:A10, B1:B10, "*Pattern*")`
- **Parameters:**
  - `cellRange` (Range): Range of cells from which values are selected if they match the condition. **(Required)**
  - `conditionRange` (Range): Range containing values to compare with `comparison`. Must be the same size as `cellRange`. **(Required)**
  - `comparison` (String): Pattern or condition to match values in `conditionRange`. **(Required)**

#### `DS_SelectAND`

Selects values from `cellRange` that meet multiple specified conditions in up to four condition ranges. All conditions must be met.

- **Usage:** `=DS_SelectAND(A1:A10, B1:B10, "*Pattern*", C1:C10, "*AnotherPattern*")`
- **Parameters:**
  - `cellRange` (Range): Range of cells from which values are selected if they match all conditions. **(Required)**
  - `conditionRange` (Range): Range to compare with `comp` for the first condition. Must match `cellRange` in size. **(Required)**
  - `comp` (String): First comparison pattern for `conditionRange`. **(Required)**
  - `conditionRange2`, `conditionRange3`, `conditionRange4` (Range): Additional ranges for the 2nd, 3rd, and 4th conditions, respectively. **(Optional)**
  - `comp2`, `comp3`, `comp4` (String): Comparison patterns for `conditionRange2`, `conditionRange3`, and `conditionRange4`, respectively. **(Optional)**

#### `DS_SelectOR`

Selects values from `cellRange` that meet at least one of up to four specified conditions across condition ranges.

- **Usage:** `=DS_SelectOR(A1:A10, B1:B10, "*Pattern*", C1:C10, "*AnotherPattern*")`
- **Parameters:**
  - `cellRange` (Range): Range of cells from which values are selected if they match any condition. **(Required)**
  - `conditionRange` (Range): Range to compare with `comp` for the first condition. Must match `cellRange` in size. **(Required)**
  - `comp` (String): First comparison pattern for `conditionRange`. **(Required)**
  - `conditionRange2`, `conditionRange3`, `conditionRange4` (Range): Additional ranges for the 2nd, 3rd, and 4th conditions, respectively. **(Optional)**
  - `comp2`, `comp3`, `comp4` (String): Comparison patterns for `conditionRange2`, `conditionRange3`, and `conditionRange4`, respectively. **(Optional)**

#### `DS_UniqueValues`

Retrieves unique values from `cellRange`, removing any duplicates.

- **Usage:** `=DS_UniqueValues(A1:A10)`
- **Parameters:**
  - `cellRange` (Range or Array): Range or array containing values to be filtered for uniqueness. **(Required)**





 ## Module: DSStatTools_UTest 🧪
This module provides statistical functions for hypothesis testing, specifically focused on the Mann-Whitney U Test (or Wilcoxon rank-sum test) for comparing two independent samples. These functions are useful for determining if there is a statistically significant difference between two sample distributions without assuming normality.

### Functions

#### `DS_UTestP`

Calculates the p-value for a Mann-Whitney U Test between two samples, indicating if there is a significant difference between the distributions.

- **Usage:** `=DS_UTestP(A1:A10, B1:B10, 2)`
- **Parameters:**
  - `cellRange1` (Range or Array): The first sample data range or array. **(Required)**
  - `cellRange2` (Range or Array): The second sample data range or array. **(Required)**
  - `sided` (Variant): Specifies if the test is one-tailed (1) or two-tailed (2). Defaults to 2 if omitted. **(Optional)**




---
...ROC function documentation coming soon...
