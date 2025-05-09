Sub WorkWithChartArea()

    Dim Chrt As ChartObject
    Dim ChrtArea As ChartArea
    Dim chrtaxs As Axis
    
    ' Create reference to chart
    
    Set Chrt = ActiveSheet.ChartObjects(1)
    Set ChrtArea = Chrt.Chart.ChartArea
    
        ' Add 3D effect
        
        ChrtArea.Format.ThreeD.BevelTopType = msoBevelCircle
        
        ' Copy chart area
        ChrtArea.Copy
        
        ' Clear Format, content, clear general
        
        ' ChrtArea.Clear
        ' ChrtArea.ClearContents
        ' ChrtArea.ClearFormats
        
        ' Add shadow
        
        With ChrtArea.Format.Shadow
            .Visible = True
            .Style = msoShadowStyleOuterShadow
            .Transparency = 0.4
            .ForeColor.RGB = RGB(36, 60, 252)
        End With
        
        ' Add corners
        ChrtArea.RoundedCorners = True
        
        ' Select chart area
        ChrtArea.Select
        
        ' Set Chart Axis
        
        Set chrtaxs = Chrt.Chart.Axes(Type:=xlValue, AxisGroup:=xlPrimary)
        
        ' Change major unit
        chrtaxs.MajorUnit = 100000
        
        ' Change the scale type
        chrtaxs.ScaleType = xlScaleLogarithmic
        
        ' Change tick label position
        chrtaxs.TickLabelPosition = xlTickLabelPositionHigh
        
        ' Change orientation
        Chrt.Chart.PlotBy = xlRows
        Chrt.Chart.PlotBy = xlColumns
        
        ' Change axis cross
        
        chrtaxs.Crosses = xlAxisCrossesMaximum
        chrtaxs.CrossesAt = 50
        
        


End Sub
