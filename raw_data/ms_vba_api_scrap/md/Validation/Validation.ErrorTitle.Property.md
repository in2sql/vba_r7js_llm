# Validation ErrorTitle Property

## Business Description
Returns or sets the title of the data-validation error dialog box. Read/write String.

## Behavior
Returns or sets the title of the data-validation error dialog box. Read/writeString.

## Example Usage
```vba
With Range("e5").Validation 
 .Add xlValidateWholeNumber, _ 
 xlValidAlertInformation, xlBetween, "5", "10" 
 .InputTitle = "Integers" 
 .ErrorTitle= "Integers" 
 .InputMessage = "Enter an integer from five to ten" 
 .ErrorMessage = "You must enter a number from five to ten" 
End With
```