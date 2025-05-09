# Worksheet Tab Property

## Business Description
Returns a Tab object for a worksheet.

## Behavior
Returns aTabobject for a worksheet.

## Example Usage
```vba
Sub CheckTab() 
 
 ' Determine if color index of 1st tab is set to none. 
 If Worksheets(1).Tab.ColorIndex = xlColorIndexNone Then 
 MsgBox "The color index is set to none for the 1st " & _ 
 "worksheet tab." 
 Else 
 MsgBox "The color index for the tab of the 1st worksheet " & _ 
 "is not set none." 
 End If 
 
End Sub
```