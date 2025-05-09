Attribute VB_Name = "Grafiek_Invoegen"
'===============================================
'@details       Function for building Gausgrafic
'@author        Jordi van Nistelrooij @ Webs en Systems
'@email         info@websensystems.nl
'@version       1.0.0
'@date          2024-09-10
'@copyright     Non of these scripts maybe copied or modified without permission of the author
'===============================================

Sub BuildGaus()

  Gemiddelde = Application.WorksheetFunction.Average(Selection)
  StandDevi = Application.WorksheetFunction.StDev_P(Selection)
  Laagste = Application.WorksheetFunction.Min(Selection)
  Hoogste = Application.WorksheetFunction.Max(Selection)
  
  Dim stdDev(0 To 24) As Double
  
  stdDev(0) = Laagste
  stdDev(4) = Gemiddelde - (2 * StandDevi)
  stdDev(8) = Gemiddelde - StandDevi
  stdDev(12) = Gemiddelde
  stdDev(16) = Gemiddelde + StandDevi
  stdDev(20) = Gemiddelde + (2 * StandDevi)
  stdDev(24) = Hoogste
  
  stdDev(1) = (stdDev(4) - stdDev(0)) * 0.25 + stdDev(0)
  stdDev(2) = (stdDev(4) - stdDev(0)) * 0.5 + stdDev(0)
  stdDev(3) = (stdDev(4) - stdDev(0)) * 0.75 + stdDev(0)
  
  stdDev(5) = (stdDev(8) - stdDev(4)) * 0.25 + stdDev(4)
  stdDev(6) = (stdDev(8) - stdDev(4)) * 0.5 + stdDev(4)
  stdDev(7) = (stdDev(8) - stdDev(4)) * 0.75 + stdDev(4)
  
  stdDev(9) = (stdDev(12) - stdDev(8)) * 0.25 + stdDev(8)
  stdDev(10) = (stdDev(12) - stdDev(8)) * 0.5 + stdDev(8)
  stdDev(11) = (stdDev(12) - stdDev(8)) * 0.75 + stdDev(8)
  
  stdDev(13) = (stdDev(16) - stdDev(12)) * 0.25 + stdDev(12)
  stdDev(14) = (stdDev(16) - stdDev(12)) * 0.5 + stdDev(12)
  stdDev(15) = (stdDev(16) - stdDev(12)) * 0.75 + stdDev(12)
  
  stdDev(17) = (stdDev(20) - stdDev(16)) * 0.25 + stdDev(16)
  stdDev(18) = (stdDev(20) - stdDev(16)) * 0.5 + stdDev(16)
  stdDev(19) = (stdDev(20) - stdDev(16)) * 0.75 + stdDev(16)
  
  stdDev(21) = (stdDev(24) - stdDev(20)) * 0.25 + stdDev(20)
  stdDev(22) = (stdDev(24) - stdDev(20)) * 0.5 + stdDev(20)
  stdDev(23) = (stdDev(24) - stdDev(20)) * 0.75 + stdDev(20)
  
  
  Dim nomDevA(0 To 24) As Double
  Dim nomDevB(0 To 20) As Double
  Dim nomDevC(0 To 16) As Double
  Dim nomDevD(0 To 12) As Double
  Dim nomDevE(0 To 8) As Double
  Dim nomDevF(0 To 4) As Double
  
  For a = 0 To 24
    nomDevA(a) = Application.WorksheetFunction.Norm_Dist(stdDev(a), Gemiddelde, StandDevi, False)
  Next
  
    For b = 0 To 20: nomDevB(b) = nomDevA(b): Next b
    
    For c = 0 To 16: nomDevC(c) = nomDevA(c): Next c
    For D = 0 To 12: nomDevD(D) = nomDevA(D): Next D
    For e = 0 To 8: nomDevE(e) = nomDevA(e): Next e
    For f = 0 To 4: nomDevF(f) = nomDevA(f): Next f
  
  ActiveSheet.Shapes.AddChart2(276, xlArea).Select
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).Values = nomDevA
    ActiveChart.FullSeriesCollection(1).XValues = stdDev
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Values = nomDevB
    
     ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).Values = nomDevC
    
     ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(4).Values = nomDevD
    
     ActiveChart.SeriesCollection.NewSeries
     ActiveChart.FullSeriesCollection(5).Values = nomDevE
    
     ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(6).Values = nomDevF
    ActiveChart.Axes(xlValue).Delete
    ActiveChart.ChartTitle.Text = "Gaus Grafiek"
     
     With ActiveChart.FullSeriesCollection(1).Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.4000000238
        .Transparency = 0
        .Solid
    End With
     With ActiveChart.FullSeriesCollection(2).Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.4000000238
        .Transparency = 0
        .Solid
    End With
     With ActiveChart.FullSeriesCollection(3).Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent3
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.4000000238
        .Transparency = 0
        .Solid
    End With
     With ActiveChart.FullSeriesCollection(4).Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent3
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.6000000238
        .Transparency = 0
        .Solid
    End With
     With ActiveChart.FullSeriesCollection(5).Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.6000000238
        .Transparency = 0
        .Solid
    End With
     
    With ActiveChart.FullSeriesCollection(6).Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.6000000238
        .Transparency = 0
        .Solid
    End With
    
    
    ActiveChart.Axes(xlCategory).HasMajorGridlines = True
    ActiveChart.Parent.Width = 600
    ActiveChart.Parent.Height = 400
    ActiveChart.Parent.Top = 50
    ActiveChart.Parent.Left = 50
End Sub

