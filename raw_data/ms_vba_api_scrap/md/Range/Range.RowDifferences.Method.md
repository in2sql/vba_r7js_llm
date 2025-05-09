# Range RowDifferences Method

## Business Description
Returns a Range object that represents all the cells whose contents are different from those of the comparison cell in each row.

## Behavior
Returns aRangeobject that represents all the cells whose contents are different from those of the comparison cell in each row.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
Set c1 = ActiveSheet.Rows(1).RowDifferences( _ 
 comparison:=ActiveSheet.Range("D1")) 
c1.Select
```