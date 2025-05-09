# Workbook IconSets Property

## Business Description
This property is used to filter data in a workbook based on a cell icon from the IconSet collection. Read-only.

## Behavior
This property is used to filter data in a workbook based on a cell icon from theIconSetcollection. Read-only.

## Example Usage
```vba
Selection.AutoFilter Field:=1, Criteria1:=ActiveWorkbook.IconSets(xl3Arrows).Item(1), Operator:=xlFilterIcon
```