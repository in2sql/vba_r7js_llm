# WorksheetFunction Object

## Business Description
Used as a container for Microsoft Excel worksheet functions that can be called from Visual Basic.

## Behavior
Used as a container for Microsoft Excel worksheet functions that can be called from Visual Basic.

## Example Usage
```vba
Set myRange = Worksheets("Sheet1").Range("A1:C10") 
answer = Application.WorksheetFunction.Min(myRange) 
MsgBox answer
```