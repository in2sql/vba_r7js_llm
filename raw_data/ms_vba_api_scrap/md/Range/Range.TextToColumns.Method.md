# Range TextToColumns Method

## Business Description
Parses a column of cells that contain text into several columns.

## Behavior
Parses a column of cells that contain text into several columns.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
ActiveSheet.Paste 
Selection.TextToColumnsDataType:=xlDelimited, _ 
 ConsecutiveDelimiter:=True, Space:=True
```