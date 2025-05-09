# Validation IMEMode Property

## Business Description
Returns or sets the description of the Japanese input rules. Can be one of the XlIMEMode constants listed in the following table. Read/write Long.

## Behavior
Returns or sets the description of the Japanese input rules. Can be one of theXlIMEModeconstants listed in the following table. Read/writeLong.

## Example Usage
```vba
With Range("E5").Validation 
    .Add Type:=xlValidateWholeNumber, _ 
        AlertStyle:= xlValidAlertStop, _ 
        Operator:=xlBetween, Formula1:="5", Formula2:="10" 
    .InputTitle = "整数値" 
    .ErrorTitle = "整数値" 
    .InputMessage = "5から10の整数を入カしてください。" 
    .ErrorMessage = "入カできるのは5から10までの値です。" 
    .IMEMode= xlIMEModeAlpha 
End With
```