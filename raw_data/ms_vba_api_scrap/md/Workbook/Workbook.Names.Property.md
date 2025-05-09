# Workbook Names Property

## Business Description
Returns a Names collection that represents all the names in the specified workbook (including all worksheet-specific names). Read-only Names object.

## Behavior
Returns aNamescollection that represents all the names in the specified workbook (including all worksheet-specific names). Read-onlyNamesobject.

## Example Usage
```vba
ActiveWorkbook.Names.Add Name:="myName", RefersToR1C1:= _ 
 "=Sheet1!R1C1"
```