# Validation Object

## Business Description
Represents data validation for a worksheet range.

## Behavior
Represents data validation for a worksheet range.

## Example Usage
```vba
Range("e5").Validation _ 
 .ModifyxlValidateList, xlValidAlertStop, "=$A$1:$A$10"
```