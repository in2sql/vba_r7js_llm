# How to: Refer to Cells by Using Shortcut Notation

## Business Description
You can use either the A1 reference style or a named range within brackets as a shortcut for the Range property. You do not have to type the word "Range" or use quotation marks, as shown in the following examples.

## Behavior
You can use either the A1 reference style or a named range within brackets as a shortcut for theRangeproperty. You do not have to type the word "Range" or use quotation marks, as shown in the following examples.

## Example Usage
```vba
Sub ClearRange() 
 Worksheets("Sheet1").[A1:B5].ClearContents 
End Sub 
 
Sub SetValue() 
 [MyRange].Value = 30 
End Sub
```