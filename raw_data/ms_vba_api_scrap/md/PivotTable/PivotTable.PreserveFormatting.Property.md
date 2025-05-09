# PivotTable PreserveFormatting Property

## Business Description
True if formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.

## Behavior
Trueif formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.For query tables, this property isTrueif any formatting common to the first five rows of data are applied to new rows of data in the query table. Unused cells aren't formatted. The property isFalseif the last AutoFormat applied to the query table is applied to new rows of data. The default value isTrue.

## Example Usage
No VBA example available.