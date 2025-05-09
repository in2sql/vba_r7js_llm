# PivotTable InnerDetail Property

## Business Description
Returns or sets the name of the field that will be shown as detail when the ShowDetail property is True for the innermost row or column field. Read/write String.

## Behavior
Returns or sets the name of the field that will be shown as detail when theShowDetailproperty isTruefor the innermost row or column field. Read/writeString.

## Example Usage
```vba
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
MsgBox pvtTable.InnerDetail
```