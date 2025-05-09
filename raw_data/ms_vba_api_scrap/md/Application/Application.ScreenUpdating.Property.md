# Application.ScreenUpdating property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Dim elapsedTime(2) 
Application.ScreenUpdating = True 
For i = 1 To 2 
 If i = 2 Then Application.ScreenUpdating = False 
 startTime = Time 
 Worksheets("Sheet1").Activate 
 For Each c In ActiveSheet.Columns 
 If c.Column Mod 2 = 0 Then 
 c.Hidden = True 
 End If 
 Next c 
 stopTime = Time 
 elapsedTime(i) = (stopTime - startTime) * 24 * 60 * 60 
Next i 
Application.ScreenUpdating = True 
MsgBox "Elapsed time, screen updating on: " & elapsedTime(1) & _ 
 " sec." & Chr(13) & _ 
 "Elapsed time, screen updating off: " & elapsedTime(2) & _ 
 " sec."
```

## Remarks
Turn screen updating off to speed up your macro code. You won't be able to see what the macro is doing, but it will run faster.

## Example
```vba
Dim elapsedTime(2) 
Application.ScreenUpdating = True 
For i = 1 To 2 
 If i = 2 Then Application.ScreenUpdating = False 
 startTime = Time 
 Worksheets("Sheet1").Activate 
 For Each c In ActiveSheet.Columns 
 If c.Column Mod 2 = 0 Then 
 c.Hidden = True 
 End If 
 Next c 
 stopTime = Time 
 elapsedTime(i) = (stopTime - startTime) * 24 * 60 * 60 
Next i 
Application.ScreenUpdating = True 
MsgBox "Elapsed time, screen updating on: " & elapsedTime(1) & _ 
 " sec." & Chr(13) & _ 
 "Elapsed time, screen updating off: " & elapsedTime(2) & _ 
 " sec."
```

