# TextFrame2 Creator Property

## Business Description
Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.

## Behavior
Returns a 32-bit integer that indicates the application in which this object was created. Read-onlyLong.

## Example Usage
```vba
Sub FindCreator() 
 
 Dim myObject As Excel.Workbook 
 Set myObject = ActiveWorkbook 
 If myObject.TextFrame2.Creator = &h5843454c Then 
 MsgBox "This is a Microsoft Excel object." 
 Else 
 MsgBox "This is not a Microsoft Excel object." 
 End If 
 
End Sub
```