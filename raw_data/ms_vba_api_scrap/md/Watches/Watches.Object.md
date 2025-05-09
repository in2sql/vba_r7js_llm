# Watches Object

## Business Description
A collection of all the Watch objects in a specified application.

## Behavior
A collection of all theWatchobjects in a specified application.

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