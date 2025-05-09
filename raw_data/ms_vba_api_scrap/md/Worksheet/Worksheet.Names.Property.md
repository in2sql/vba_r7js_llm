# Worksheet Names Property

## Business Description
Returns a Names collection that represents all the worksheet-specific names (names defined with the "WorksheetName!" prefix). Read-only Names object.

## Behavior
Returns aNamescollection that represents all the worksheet-specific names (names defined with the "WorksheetName!" prefix). Read-onlyNamesobject.

## Example Usage
```vba
ActiveWorkbook.Names.Add Name:="myName", RefersToR1C1:= _ 
 "=Sheet1!R1C1"
```