# Areas.Count property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub DisplayColumnCount() 
 Dim iAreaCount As Integer 
 Dim i As Integer 
 
 Worksheets("Sheet1").Activate 
 iAreaCount = Selection.Areas.Count 
 
 If iAreaCount <= 1 Then 
 MsgBox "The selection contains " & Selection.Columns.Count & " columns." 
 Else 
 For i = 1 To iAreaCount 
 MsgBox "Area " & i & " of the selection contains " & _ 
 Selection.Areas(i).Columns.Count & " columns." 
 Next i 
 End If 
End Sub
```

## Example
```vba
Sub DisplayColumnCount() 
 Dim iAreaCount As Integer 
 Dim i As Integer 
 
 Worksheets("Sheet1").Activate 
 iAreaCount = Selection.Areas.Count 
 
 If iAreaCount <= 1 Then 
 MsgBox "The selection contains " & Selection.Columns.Count & " columns." 
 Else 
 For i = 1 To iAreaCount 
 MsgBox "Area " & i & " of the selection contains " & _ 
 Selection.Areas(i).Columns.Count & " columns." 
 Next i 
 End If 
End Sub
```

