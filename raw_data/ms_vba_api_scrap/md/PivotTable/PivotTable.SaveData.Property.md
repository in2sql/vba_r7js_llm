# PivotTable SaveData Property

## Business Description
True if data for the PivotTable report is saved with the workbook. False if only the report definition is saved. Read/write Boolean.

## Behavior
Trueif data for the PivotTable report is saved with the workbook.Falseif only the report definition is saved. Read/writeBoolean.

## Example Usage
```vba
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
pvtTable.SaveData= True
```