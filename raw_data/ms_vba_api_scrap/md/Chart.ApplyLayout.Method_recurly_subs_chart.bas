Attribute VB_Name = "Module84"
Sub recurly_subs_chart()

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

' makes the first chart

    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("B11:F11")
    ActiveChart.SetSourceData Source:=Range("B11:F11,B17:F17")
    ActiveChart.ChartColor = 13
    ActiveChart.ClearToMatchStyle
    ActiveChart.ChartStyle = 209
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).ApplyDataLabels
    ActiveChart.FullSeriesCollection(1).Points(2).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ActiveChart.FullSeriesCollection(1).Points(3).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent5
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ActiveChart.FullSeriesCollection(1).Points(4).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent4
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Points(5).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(112, 48, 160)
        .Transparency = 0
        .Solid
    End With
    
' makes the second chart

    ActiveSheet.Shapes.AddChart2(297, xlColumnStacked100).Select
    ActiveChart.SetSourceData Source:=Range("A19:F23")
    ActiveChart.PlotBy = xlColumns
    ActiveChart.ClearToMatchStyle
    ActiveChart.ApplyLayout (4)
    ActiveChart.ChartStyle = 304
    ActiveChart.ChartColor = 13
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).ApplyDataLabels
    ActiveChart.FullSeriesCollection(2).Select
    ActiveChart.FullSeriesCollection(2).ApplyDataLabels
    ActiveChart.FullSeriesCollection(3).Select
    ActiveChart.FullSeriesCollection(3).ApplyDataLabels
    ActiveChart.FullSeriesCollection(4).Select
    ActiveChart.FullSeriesCollection(4).ApplyDataLabels
    ActiveChart.FullSeriesCollection(5).Select
    ActiveChart.FullSeriesCollection(5).ApplyDataLabels
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    With Selection.Format.TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Solid
    End With
    
' makes the third chart
    
    ActiveSheet.Shapes.AddChart2(297, xlColumnStacked100).Select
    ActiveChart.SetSourceData Source:=Range("A25:F29")
    ActiveChart.PlotBy = xlRows
    ActiveChart.ClearToMatchStyle
    ActiveChart.ApplyLayout (4)
    ActiveChart.ChartStyle = 304
    ActiveChart.ChartColor = 13
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).ApplyDataLabels
    ActiveChart.FullSeriesCollection(2).Select
    ActiveChart.FullSeriesCollection(2).ApplyDataLabels
    ActiveChart.FullSeriesCollection(3).Select
    ActiveChart.FullSeriesCollection(3).ApplyDataLabels
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    With Selection.Format.TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Solid
    End With
    ActiveChart.FullSeriesCollection(4).Select
    ActiveChart.FullSeriesCollection(4).ApplyDataLabels
    
' move charts

    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Chart 3").IncrementLeft 792
    ActiveSheet.Shapes("Chart 3").IncrementTop 6.75
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Chart 2").IncrementLeft 792.75
    ActiveSheet.Shapes("Chart 2").IncrementTop 228
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Chart 1").IncrementLeft 794.25
    ActiveSheet.Shapes("Chart 1").IncrementTop 454.5
    
Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub

