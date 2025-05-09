Sub WorkWithSeries()

    Dim Chrt As ChartObject
    Dim ChtSerColl As SeriesCollection
    Dim ChtSer As Series
    
    ' Create reference to chart
    
    Set Chrt = ActiveSheet.ChartObjects(1)
    
    ' Get series collection
    
    Set ChtSerColl = Chrt.Chart.SeriesCollection
    
    ' Print the name of each series
    
    For Each ChtSer In ChtSerColl
        Debug.Print ChtSer.Name
    Next
    
    ' Selecting one series from collection
    
    Set ChtSer = ChtSerColl.Item("Profit")
    
    ' With profit
    
    With ChtSer
    
        ' Add data labels
        
        .HasDataLabels = True
        .ApplyDataLabels Type:=xlValue
        
        ' Change fill color of series
        .Format.Fill.ForeColor.RGB = RGB(34, 60, 252)
        
        ' Add some error bars
        
        .HasErrorBars = True
        
        ' Add a leader line
        .HasLeaderLines = True
        .LeaderLines.Border.Color = vbRed
        
        'Move the label away to demo
        
        ' Format the series border
        
        With .Format
            .Line.Visible = True
            .Line.Weight = 3
            .Line.DashStyle = msoLineDashDot
            .Line.ForeColor.TintAndShade = 1
            
        End With
        
        ' Print the axis it belongs to
        
        Debug.Print ChtSer.AxisGroup
        
        ' Is series filtered?
        
        Debug.Print ChtSer.IsFiltered
        
        
     End With
     
     
     
     

End Sub
