# How to: Refer to Cells by Using a Range Object

## Business Description
If you set an object variable to a Range object, you can easily manipulate the range by using the variable name.

## Behavior
If you set an object variable to aRangeobject, you can easily manipulate the range by using the variable name.

## Example Usage
```vba
Sub Random() 
 Dim myRange As Range 
 Set myRange = Worksheets("Sheet1").Range("A1:D5") 
 myRange.Formula = "=RAND()" 
 myRange.Font.Bold = True 
End Sub
```