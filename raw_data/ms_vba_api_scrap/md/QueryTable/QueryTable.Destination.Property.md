# QueryTable Destination Property

## Business Description
Returns the cell in the upper-left corner of the query table destination range (the range where the resulting query table will be placed). The destination range must be on the worksheet that contains the QueryTable object. Read-only Range.

## Behavior
Returns the cell in the upper-left corner of the query table destination range (the range where the resulting query table will be placed). The destination range must be on the worksheet that contains theQueryTableobject. Read-onlyRange.

## Example Usage
```vba
Set d = Worksheets(1).QueryTables(1).DestinationWith ActiveWindow 
 .ScrollColumn = d.Column 
 .ScrollRow = d.Row 
End With
```