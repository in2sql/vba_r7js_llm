# Validation ShowError Property

## Business Description
True if the data validation error message will be displayed whenever the user enters invalid data. Read/write Boolean.

## Behavior
Trueif the data validation error message will be displayed whenever the user enters invalid data. Read/writeBoolean.

## Example Usage
```vba
With Worksheets(1).Range("A10").Validation 
 .Add Type:=xlValidateWholeNumber, _ 
 AlertStyle:=xlValidAlertStop, _ 
 Operator:=xlBetween, Formula1:="5", _ 
 Formula2:="10" 
 .ErrorMessage = "value must be between 5 and 10" 
 .ShowInput = False 
 .ShowError= True 
End With
```