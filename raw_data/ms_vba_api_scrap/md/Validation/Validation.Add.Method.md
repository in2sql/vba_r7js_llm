# Validation Add Method

## Business Description
Adds data validation to the specified range.

## Behavior
Adds data validation to the specified range.

## Example Usage
```vba
With Range("e5").Validation 
 .AddType:=xlValidateWholeNumber, _ 
 AlertStyle:= xlValidAlertStop, _ 
 Operator:=xlBetween, Formula1:="5", Formula2:="10" 
 .InputTitle = "Integers" 
 .ErrorTitle = "Integers" 
 .InputMessage = "Enter an integer from five to ten" 
 .ErrorMessage = "You must enter a number from five to ten" 
End With
```