# Range Text Property

## Business Description
Returns or sets the text for the specified object. Read-only String.

## Behavior
Returns or sets the text for the specified object. Read-onlyString.

## Example Usage
```vba
Set c = Worksheets("Sheet1").Range("B14") 
c.Value = 1198.3 
c.NumberFormat = "$#,##0_);($#,##0)" 
MsgBox c.Value 
MsgBox c.Text
```