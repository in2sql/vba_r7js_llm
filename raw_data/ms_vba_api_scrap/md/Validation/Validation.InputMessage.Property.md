# Validation InputMessage Property

## Business Description
Returns or sets the data validation input message. Read/write String.

## Behavior
Returns or sets the data validation input message. Read/writeString.

## Example Usage
```vba
With Range("e5").Validation 
 .Add Type:=xlValidateWholeNumber, _ 
 AlertStyle:= xlValidAlertStop, _ 
 Operator:=xlBetween, Formula1:="5", Formula2:="10" 
 .InputTitle = "Integers" 
 .ErrorTitle = "Integers" 
 .InputMessage= "Enter an integer from five to ten" 
 .ErrorMessage = "You must enter a number from five to ten" 
End With
```