# Watch Object

## Business Description
Represents a range which is tracked when the worksheet is recalculated. The Watch object allows users to verify the accuracy of their models and debug problems they encounter.

## Behavior
Represents a range which is tracked when the worksheet is recalculated. TheWatchobject allows users to verify the accuracy of their models and debug problems they encounter.

## Example Usage
```vba
Sub AddWatch() 
 
 With Application 
 .Range("A1").Formula = 1 
 .Range("A2").Formula = 2 
 .Range("A3").Formula = "=Sum(A1:A2)" 
 .Range("A3").Select 
 .Watches.Add Source:=ActiveCell 
 End With 
 
End Sub
```