# Looping Through a Range of Cells

## Business Description
When using Visual Basic, you often need to run the same block of statements on each cell in a range of cells. To do this, you combine a looping statement and one or more methods to identify each cell, one at a time, and run the operation.

## Behavior
When using Visual Basic, you often need to run the same block of statements on each cell in a range of cells. To do this, you combine a looping statement and one or more methods to identify each cell, one at a time, and run the operation.

## Example Usage
```vba
Sub RoundToZero1() 
 For Counter = 1 To 20 
 Set curCell = Worksheets("Sheet1").Cells(Counter, 3) 
 If Abs(curCell.Value) < 0.01 Then curCell.Value = 0 
 Next Counter 
End Sub
```