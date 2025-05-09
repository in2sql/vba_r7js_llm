Attribute VB_Name = "Module89"
Sub recurly_subs_current_chart()

' makes a chart for the CurrentStatus window

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

    ActiveSheet.Shapes.AddChart2(297, xlColumnStacked100).Select
    ActiveChart.SetSourceData Source:=Range("A19:E23")
    ActiveChart.PlotBy = xlColumns
    ActiveChart.ClearToMatchStyle
    ActiveChart.ChartStyle = 304
    ActiveChart.ChartColor = 13
    ActiveChart.ApplyLayout (4)
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).ApplyDataLabels
    ActiveChart.FullSeriesCollection(2).Select
    ActiveChart.FullSeriesCollection(2).ApplyDataLabels
    ActiveChart.FullSeriesCollection(3).Select
    ActiveChart.FullSeriesCollection(3).ApplyDataLabels
    ActiveChart.FullSeriesCollection(4).Select
    ActiveChart.FullSeriesCollection(4).ApplyDataLabels
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    With Selection.Format.TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Solid
    End With

' move chart

    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Chart 1").IncrementLeft 536.25
    ActiveSheet.Shapes("Chart 1").IncrementTop 6.75

Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub

