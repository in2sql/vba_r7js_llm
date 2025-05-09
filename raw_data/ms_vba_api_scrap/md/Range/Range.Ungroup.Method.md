# Range Ungroup Method

## Business Description
Promotes a range in an outline (that is, decreases its outline level). The specified range must be a row or column, or a range of rows or columns. If the range is in a PivotTable report, this method ungroups the items contained in the range.

## Behavior
Promotes a range in an outline (that is, decreases its outline level). The specified range must be a row or column, or a range of rows or columns. If the range is in a PivotTable report, this method ungroups the items contained in the range.

## Example Usage
```vba
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
Set groupRange = pvtTable.PivotFields("ORDER_DATE").DataRange 
groupRange.Cells(1).Ungroup
```