# PivotTable PrintTitles Property

## Business Description
True if the print titles for the worksheet are set based on the PivotTable report. False if the print titles for the worksheet are used. The default value is False. Read/write Boolean.

## Behavior
Trueif the print titles for the worksheet are set based on the PivotTable report.Falseif the print titles for the worksheet are used. The default value isFalse. Read/writeBoolean.

## Example Usage
```vba
ActiveSheet.PivotTables("PivotTable4").PrintTitles= True
```