# Tab Object

## Business Description
Represents a tab in a chart or a worksheet.

## Behavior
Represents a tab in a chart or a worksheet.

## Example Usage
```vba
Sub CheckTab() 
 
 ' Determine if color index of 1st tab is set to none. 
 If Worksheets(1).Tab.ColorIndex = xlColorIndexNone Then 
 MsgBox "The color index is set to none for the first " & _ 
 "worksheet tab." 
 Else 
 MsgBox "The color index for the tab of the first worksheet " & _ 
 "is not set none." 
 End If 
 
End Sub
```