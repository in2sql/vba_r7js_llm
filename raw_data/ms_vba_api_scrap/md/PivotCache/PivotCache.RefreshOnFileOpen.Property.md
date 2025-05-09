# PivotCache RefreshOnFileOpen Property

## Business Description
True if the PivotTable cache is automatically updated each time the workbook is opened. The default value is False. Read/write Boolean.

## Behavior
Trueif the PivotTable cache is automatically updated each time the workbook is opened. The default value isFalse. Read/writeBoolean.

## Example Usage
```vba
ActiveWorkbook.PivotCaches(1).RefreshOnFileOpen= True
```