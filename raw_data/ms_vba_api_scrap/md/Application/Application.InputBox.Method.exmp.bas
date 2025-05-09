Set myRange = Application.InputBox(prompt := "Sample", type := 8)

' ===== Next Example =====

Worksheets("Sheet1").Activate 
Set myCell = Application.InputBox( _ 
    prompt:="Select a cell", Type:=8)

' ===== Next Example =====

Sub Cbm_Value_Select()
   'Set up the variables.
   Dim rng As Range
   
   'Use the InputBox dialog to set the range for MyFunction, with some simple error handling.
   Set rng = Application.InputBox("Range:", Type:=8)
   If rng.Cells.Count <> 3 Then
     MsgBox "Length, width and height are needed -" & _
         vbLf & "please select three cells!"
      Exit Sub
   End If
   
   'Call MyFunction by value using the active cell.
   ActiveCell.Value = MyFunction(rng)
End Sub

Function MyFunction(rng As Range) As Double
   MyFunction = rng(1) * rng(2) * rng(3)
End Function