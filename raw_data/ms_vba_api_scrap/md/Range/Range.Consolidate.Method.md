# Range Consolidate Method

## Business Description
Consolidates data from multiple ranges on multiple worksheets into a single range on a single worksheet. Variant.

## Behavior
Consolidates data from multiple ranges on multiple worksheets into a single range on a single worksheet.Variant.

## Example Usage
```vba
Worksheets("Sheet1").Range("A1").Consolidate_ 
 Sources:=Array("Sheet2!R1C1:R37C6", "Sheet3!R1C1:R37C6"), _ 
 Function:=xlSum
```