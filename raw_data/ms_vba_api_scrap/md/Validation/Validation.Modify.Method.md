# Validation Modify Method

## Business Description
Modifies data validation for a range.

## Behavior
Modifies data validation for a range.

## Example Usage
```vba
Range("e5").Validation _ 
 .ModifyxlValidateList, xlValidAlertStop, _ 
 xlBetween, "=$A$1:$A$10"
```