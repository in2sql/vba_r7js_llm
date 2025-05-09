# Validation ShowInput Property

## Business Description
True if the data validation input message will be displayed whenever the user selects a cell in the data validation range. Read/write Boolean.

## Behavior
Trueif the data validation input message will be displayed whenever the user selects a cell in the data validation range. Read/writeBoolean.

## Example Usage
```vba
With Worksheets(1).Range("A10").Validation 
 .Add Type:=xlValidateWholeNumber, _ 
 AlertStyle:=xlValidAlertStop, _ 
 Operator:=xlBetween, Formula1:="5", _ 
 Formula2:="10" 
 .ErrorMessage = "value must be between 5 and 10" 
 .ShowInput= False 
 .ShowError = True 
End With
```