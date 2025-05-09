# PivotTable ConvertToFormulas Method

## Business Description
The ConvertToFormulas method is new in Microsoft Office Excel 2007 and is used for converting a PivotTable to cube formulas. Read/write Boolean.

## Behavior
TheConvertToFormulasmethod is new in Microsoft Office Excel 2007 and is used for converting a PivotTable to cube formulas.  Read/writeBoolean.

## Example Usage
```vba
Sub ConvertToCubeFormulas() 
 ActiveSheet.PivotTables("PivotTable1").ConvertToFormulas False 
End Sub
```