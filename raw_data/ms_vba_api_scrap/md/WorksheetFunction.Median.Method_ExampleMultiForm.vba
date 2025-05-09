Option Explicit

Private Sub btnCancel_Click()
    Unload Me
    End
End Sub


Private Sub opavg_Click()
ThisWorkbook.Sheets("ChartData").Visible = True
ThisWorkbook.Sheets("Stats").Select
End Sub

Private Sub opcol_Click()
ThisWorkbook.Sheets("ChartData").Visible = True
ThisWorkbook.Sheets("ChartData").Select
Dim column As Chart, columndata As Range, fname As String
Set columndata = ActiveSheet.Range("C5:C9")
Set column = ActiveSheet.Shapes.AddChart(xlColumnClustered).Chart
column.SeriesCollection.NewSeries
column.SeriesCollection(1).Name = "Assets"
column.SeriesCollection(1).Values = columndata
column.SeriesCollection(1).XValues = ActiveSheet.Range("B5:B9")
Set column = ActiveSheet.ChartObjects(1).Chart
column.Parent.Width = 390
column.Parent.Height = 198
fname = ThisWorkbook.Path & Application.PathSeparator & "pnguyentempcharts.gif"
column.Export Filename:=fname, FilterName:="GIF"
Image1.Picture = LoadPicture(fname)
ActiveSheet.ChartObjects(1).Delete

End Sub

Private Sub opline_Click()
ThisWorkbook.Sheets("ChartData").Visible = True
ThisWorkbook.Sheets("ChartData").Select
Dim line As Chart, linedata As Range, fname As String
Set linedata = ActiveSheet.Range("C5:C9")
Set line = ActiveSheet.Shapes.AddChart(xlLine).Chart
line.SeriesCollection.NewSeries
line.SeriesCollection(1).Name = "Assets"
line.SeriesCollection(1).Values = linedata
line.SeriesCollection(1).XValues = ActiveSheet.Range("B5:B9")
Set line = ActiveSheet.ChartObjects(1).Chart
line.Parent.Width = 390
line.Parent.Height = 198
fname = ThisWorkbook.Path & Application.PathSeparator & "pnguyentempcharts.gif"
line.Export Filename:=fname, FilterName:="GIF"
Image1.Picture = LoadPicture(fname)
ActiveSheet.ChartObjects(1).Delete

End Sub

Private Sub oppie_Click()
ThisWorkbook.Sheets("ChartData").Visible = True
ThisWorkbook.Sheets("ChartData").Select
Dim pie As Chart, piedata As Range, fname As String
Set piedata = ActiveSheet.Range("C5:C9")
Set pie = ActiveSheet.Shapes.AddChart(xlPie).Chart
pie.SeriesCollection.NewSeries
pie.SeriesCollection(1).Name = "Assets"
pie.SeriesCollection(1).Values = piedata
pie.SeriesCollection(1).XValues = ActiveSheet.Range("B5:B9")
Set pie = ActiveSheet.ChartObjects(1).Chart
pie.Parent.Width = 390
pie.Parent.Height = 198
fname = ThisWorkbook.Path & Application.PathSeparator & "pnguyentempcharts.gif"
pie.Export Filename:=fname, FilterName:="GIF"
Image1.Picture = LoadPicture(fname)
ActiveSheet.ChartObjects(1).Delete

End Sub


Private Sub UserForm_Initialize()
MultiPage1.Value = 0
End Sub

Private Sub btnOK_Click()
Dim Data As String, min, max, avg, std, med As Double, rng As Range
Data = RefEdit1.Value
Set rng = Range(Data)
min = WorksheetFunction.min(rng)
max = WorksheetFunction.max(rng)
avg = WorksheetFunction.Average(rng)
std = WorksheetFunction.StDev(rng)
med = WorksheetFunction.Median(rng)

MsgBox "Average : " & avg & vbNewLine _
& "Standard Deviation : " & std & vbNewLine _
& "Max : " & max & vbNewLine _
& "Min : " & min & vbNewLine _
& "Median: " & med
End Sub

Private Sub btnok1_Click()
ThisWorkbook.Sheets("RetrieveData").Visible = True
ThisWorkbook.Sheets("RetrieveData").Select
Select Case True
    Case optibm.Value
        Call IBMM
    Case optapple.Value
        Call APPLEM
    Case optgoogle.Value
        Call GOOGLEM
End Select
End Sub
