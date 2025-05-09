Attribute VB_Name = "Module1"
Sub Generate_Plot()


Dim myrnglabel As Range
Dim labelrange As String

Set myrnglabel = Application.InputBox(Prompt:="Select range label", _
                                    Title:="Range label", Type:=8)
labelrange = "='" & ActiveSheet.Name & "'!" & myrnglabel.Address
                                    
Application.ScreenUpdating = False 'turns off screen updating
ActiveWindow.Zoom = 100 'resets zoom to 100% so we'd have uniform aspect ratio for charts regardless of individual user zoom
                            
Set Myrange = Selection
'ActiveChart.ChartType = xlXYScatter
ActiveSheet.Shapes.AddChart2(240, xlXYScatter, Width:=510.2, Height:=396.7).Select 'creating scatter plot
ActiveChart.SetSourceData Source:=Myrange 'determining source of data for scatterplot
ActiveChart.ChartArea.Select
 

    ActiveChart.FullSeriesCollection(1).Select
    With Selection
        .MarkerStyle = 8
        .MarkerSize = 5
        .MarkerBackgroundColor = RGB(0, 32, 96)
        .MarkerForegroundColor = RGB(0, 32, 96)
    End With
    Selection.MarkerSize = 8



    ActiveChart.FullSeriesCollection(1).ApplyDataLabels
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    ActiveChart.SeriesCollection(1).DataLabels.Format.TextFrame2.TextRange. _
        InsertChartField msoChartFieldRange, labelrange, 0

'determining scatterplots dimensions (gridlines and scale) and establishing which range or parameters determine such dimensions so data can be to scale

    Selection.ShowRange = True
    Selection.ShowValue = False
    Application.CutCopyMode = False
    ActiveChart.ChartArea.Select
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = Range("C44")
    ActiveChart.Axes(xlValue).MaximumScale = Range("C45")
    ActiveChart.Axes(xlValue).CrossesAt = 0
    ActiveChart.Axes(xlValue).CrossesAt = Range("C43")
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.PlotArea.Select
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MinimumScale = Range("B44")
    ActiveChart.Axes(xlCategory).MaximumScale = Range("B45")
    ActiveChart.Axes(xlCategory).CrossesAt = 0
    ActiveChart.Axes(xlCategory).CrossesAt = Range("B43")
    ActiveChart.Axes(xlValue).Select
    Selection.TickLabelPosition = xlNone
    ActiveChart.Axes(xlCategory).Select
    Selection.TickLabelPosition = xlNone
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    Selection.Delete
    ActiveChart.Axes(xlCategory).MajorGridlines.Select
    Selection.Delete
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = Range("B47")
    ActiveChart.SetElement (msoElementChartTitleCenteredOverlay)
    
    'formatting scatterplot "plot area" or rather the background
    
      ActiveChart.PlotArea.Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.150000006
        .Transparency = 0
    End With
    ActiveChart.ChartArea.Select
    With Selection.Format.Line
        .Visible = msoFalse
    End With
    
    Set cht = ActiveChart

For Each srs In cht.SeriesCollection

    srs.Values = srs.Values
    srs.XValues = srs.XValues

Next

    
End Sub

' this is an expansion of the previous model in case we have a larger number of variables

Sub Generate_Plot_30()


Dim myrnglabel As Range
Dim labelrange As String

Set myrnglabel = Application.InputBox(Prompt:="Select range label", _
                                    Title:="Range label", Type:=8)
labelrange = "='" & ActiveSheet.Name & "'!" & myrnglabel.Address
                                    
Application.ScreenUpdating = False 'turns off screen updating
ActiveWindow.Zoom = 100 'resets zoom to 100%
                            
Set Myrange = Selection
'ActiveChart.ChartType = xlXYScatter
ActiveSheet.Shapes.AddChart2(240, xlXYScatter, Width:=510.2, Height:=396.7).Select
ActiveChart.SetSourceData Source:=Myrange
ActiveChart.ChartArea.Select
 

    ActiveChart.FullSeriesCollection(1).Select
    With Selection
        .MarkerStyle = 8
        .MarkerSize = 5
        .MarkerBackgroundColor = RGB(0, 32, 96)
        .MarkerForegroundColor = RGB(0, 32, 96)
    End With
    Selection.MarkerSize = 8


    ActiveChart.FullSeriesCollection(1).ApplyDataLabels
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    ActiveChart.SeriesCollection(1).DataLabels.Format.TextFrame2.TextRange. _
        InsertChartField msoChartFieldRange, labelrange, 0
        
    Selection.ShowRange = True
    Selection.ShowValue = False
    Application.CutCopyMode = False
    ActiveChart.ChartArea.Select
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = Range("C53")
    ActiveChart.Axes(xlValue).MaximumScale = Range("C54")
    ActiveChart.Axes(xlValue).CrossesAt = 0
    ActiveChart.Axes(xlValue).CrossesAt = Range("C52")
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.PlotArea.Select
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MinimumScale = Range("B53")
    ActiveChart.Axes(xlCategory).MaximumScale = Range("B54")
    ActiveChart.Axes(xlCategory).CrossesAt = 0
    ActiveChart.Axes(xlCategory).CrossesAt = Range("B52")
    ActiveChart.Axes(xlValue).Select
    Selection.TickLabelPosition = xlNone
    ActiveChart.Axes(xlCategory).Select
    Selection.TickLabelPosition = xlNone
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    Selection.Delete
    ActiveChart.Axes(xlCategory).MajorGridlines.Select
    Selection.Delete
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = Range("B56")
    ActiveChart.SetElement (msoElementChartTitleCenteredOverlay)
    
      ActiveChart.PlotArea.Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.150000006
        .Transparency = 0
    End With
    ActiveChart.ChartArea.Select
    With Selection.Format.Line
        .Visible = msoFalse
    End With
    
    Set cht = ActiveChart

For Each srs In cht.SeriesCollection

    srs.Values = srs.Values
    srs.XValues = srs.XValues

Next

End Sub

' this is a further expansion of the original model in case we have a larger number of variables


Sub Generate_Plot_50()


Dim myrnglabel As Range
Dim labelrange As String

Set myrnglabel = Application.InputBox(Prompt:="Select range label", _
                                    Title:="Range label", Type:=8)
labelrange = "='" & ActiveSheet.Name & "'!" & myrnglabel.Address
                                    
Application.ScreenUpdating = False 'turns off screen updating
ActiveWindow.Zoom = 100 'resets zoom to 100%
                            
Set Myrange = Selection
'ActiveChart.ChartType = xlXYScatter
ActiveSheet.Shapes.AddChart2(240, xlXYScatter, Width:=510.2, Height:=396.7).Select
ActiveChart.SetSourceData Source:=Myrange
ActiveChart.ChartArea.Select
 

    ActiveChart.FullSeriesCollection(1).Select
    With Selection
        .MarkerStyle = 8
        .MarkerSize = 5
        .MarkerBackgroundColor = RGB(0, 32, 96)
        .MarkerForegroundColor = RGB(0, 32, 96)
    End With
    Selection.MarkerSize = 8


    ActiveChart.FullSeriesCollection(1).ApplyDataLabels
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    ActiveChart.SeriesCollection(1).DataLabels.Format.TextFrame2.TextRange. _
        InsertChartField msoChartFieldRange, labelrange, 0
        
    Selection.ShowRange = True
    Selection.ShowValue = False
    Application.CutCopyMode = False
    ActiveChart.ChartArea.Select
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = Range("C73")
    ActiveChart.Axes(xlValue).MaximumScale = Range("C74")
    ActiveChart.Axes(xlValue).CrossesAt = 0
    ActiveChart.Axes(xlValue).CrossesAt = Range("C72")
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.PlotArea.Select
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MinimumScale = Range("B73")
    ActiveChart.Axes(xlCategory).MaximumScale = Range("B74")
    ActiveChart.Axes(xlCategory).CrossesAt = 0
    ActiveChart.Axes(xlCategory).CrossesAt = Range("B72")
    ActiveChart.Axes(xlValue).Select
    Selection.TickLabelPosition = xlNone
    ActiveChart.Axes(xlCategory).Select
    Selection.TickLabelPosition = xlNone
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    Selection.Delete
    ActiveChart.Axes(xlCategory).MajorGridlines.Select
    Selection.Delete
    ActiveChart.ChartTitle.Select
    Selection.Delete
    
      ActiveChart.PlotArea.Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.150000006
        .Transparency = 0
    End With
    ActiveChart.ChartArea.Select
    With Selection.Format.Line
        .Visible = msoFalse
    End With
    

End Sub


