Attribute VB_Name = "Module1"
Sub URLPictureInsert()
Dim Pshp As Shape
Dim xRg As Range
Dim xCol As Long
Cells.RowHeight = 90
Cells.ColumnWidth = 30
On Error Resume Next
Application.ScreenUpdating = False
Set Rng = ActiveSheet.Range(Range("C2"), Range("C2").End(xlDown))
For Each cell In Rng
filenam = cell
ActiveSheet.Pictures.Insert(filenam).Select
Set Pshp = Selection.ShapeRange.Item(1)
If Pshp Is Nothing Then GoTo lab
xCol = cell.Column + 1
Set xRg = Cells(cell.Row, xCol)
With Pshp
.LockAspectRatio = msoFalse
If .Width > xRg.Width Then .Width = xRg.Width * 2 / 2
If .Height > xRg.Height Then .Height = xRg.Height * 2 / 2
.Top = xRg.Top + (xRg.Height - .Height)
.Left = xRg.Left + (xRg.Width - .Width)
End With
lab:
Set Pshp = Nothing
Range("C2").Select
Next
Columns("C").Hidden = True
Application.ScreenUpdating = True
End Sub
