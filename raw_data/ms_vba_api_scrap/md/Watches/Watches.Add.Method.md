# Watches Add Method

## Business Description
Adds a range which is tracked when the worksheet is recalculated.

## Behavior
Adds a range which is tracked when the worksheet is recalculated.

## Example Usage
```vba
Sub AddWatch() 
 
 With Application 
 .Range("A1").Formula = 1 
 .Range("A2").Formula = 2 
 .Range("A3").Formula = "=Sum(A1:A2)" 
 .Range("A3").Select 
 .Watches.AddSource:=ActiveCell 
 End With 
 
End Sub
```