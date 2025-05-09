# Using Excel Worksheet Functions in Visual Basic

## Business Description
You can use most Microsoft Excel worksheet functions in your Visual Basic statements. For a list of the worksheet functions you can use, see List of Worksheet Functions Available to Visual Basic.

## Behavior
You can use most Microsoft Excel worksheet functions in your Visual Basic statements. For a list of the worksheet functions you can use, seeList of Worksheet Functions Available to Visual Basic.

## Example Usage
```vba
Sub UseFunction() 
 Dim myRange As Range 
 Set myRange = Worksheets("Sheet1").Range("A1:C10") 
 answer = Application.WorksheetFunction.Min(myRange) 
 MsgBox answer 
End Sub
```