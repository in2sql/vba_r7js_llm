# Validation InCellDropdown Property

## Business Description
True if data validation displays a drop-down list that contains acceptable values. Read/write Boolean.

## Behavior
Trueif data validation displays a drop-down list that contains acceptable values. Read/writeBoolean.

## Example Usage
```vba
With Range("e5").Validation 
 .Add xlValidateList, xlValidAlertStop, xlBetween, "=$A$1:$A$10" 
 .InCellDropdown= True 
End With
```