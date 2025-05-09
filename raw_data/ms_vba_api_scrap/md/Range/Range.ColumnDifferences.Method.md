# Range ColumnDifferences Method

## Business Description
Returns a Range object that represents all the cells whose contents are different from the comparison cell in each column.

## Behavior
Returns aRangeobject that represents all the cells whose contents are different from the comparison cell in each column.

## Example Usage
```vba
Sub CompDiff() 
'Setting up data to be compared 
 Range("A1").Value = "Rod" 
 Range("A2").Value = "Bill" 
 Range("A3").Value = "John" 
 Range("A4").Value = "Rod" 
 Range("A5").Value = "Kelly" 
 Range("A6").Value = "Rod" 
 Range("A7").Value = "Paddy" 
 Range("A8").Value = "Rod" 
 Range("A9").Value = "Rod" 
 Range("A10").Value = "Rod" 
 
'Code to do the comparison, selects the values that are unlike A1 
Worksheets("Sheet1").Activate 
Set r1 = ActiveSheet.Columns("A").ColumnDifferences( _ 
 Comparison:=ActiveSheet.Range("A1")) 
r1.Select 
End Sub
```