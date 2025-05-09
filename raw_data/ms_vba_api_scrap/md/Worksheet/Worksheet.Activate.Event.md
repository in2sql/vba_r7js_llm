# Worksheet Activate Event

## Business Description
Occurs when a workbook, worksheet, chart sheet, or embedded chart is activated.

## Behavior
Occurs when a workbook, worksheet, chart sheet, or embedded chart is activated.

## Example Usage
```vba
Private Sub Worksheet_Activate() 
 Range("a1:a10").Sort Key1:=Range("a1"), Order:=xlAscending 
End Sub
```