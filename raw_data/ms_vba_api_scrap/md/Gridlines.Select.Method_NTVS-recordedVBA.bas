Attribute VB_Name = "Module1"
Sub CreateNTVSTableBackup()
    Dim myChartObject As ChartObject
    Dim Nk As Series
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    Range("N19:N30,O37:O48,Q37:Q48").Select
    Range("Q37").Activate
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    'ActiveSheet.ChartObjects.Add(Left:=200, Top:=200, Width:=400, Height:=300)
    'Set myChartObject = Charts.Add
    'Charts.Add
    'ActiveChart.ChartType = xlLineMarkers
    ActiveSheet.Shapes.AddChart2(332, xlLineMarkers).Select
    ActiveChart.SetSourceData Source:=Range( _
        "'TVvCheck@L=0.0001(4a)'!$Q$19:$Q$30,'TVvCheck@L=0.0001(4a)'!$O$37:$O$48,'TVvCheck@L=0.0001(4a)'!$Q$37:$Q$48" _
        )
    'Set MyNewSrs = ActiveChart.SeriesCollection.NewSeries
    'Set Nk = ActiveChart.FullSeriesCollection(1).XValues = "='TVvCheck@L=0.0001(4a)'!$N$19:$N$30"
    ActiveChart.FullSeriesCollection(1).XValues = "='TVvCheck@L=0.0001(4a)'!$N$19:$N$30"
    'ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).Name = "=""NTVS, TV Check Off"""
    'ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "=""NTVS, TV Check On"""
    'ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).Name = "=""Original TVS"""
    'ActiveChart.SeriesCollection.NewSeries
    'ActiveChart.FullSeriesCollection(4).XValues = "='TVvCheck@L=0.0001(4a)'!$N$19:$N$30"
    
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Line
        .Visible = msoFalse
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground2
        .ForeColor.TintAndShade = 0
        '.ForeColor.Brightness = -0.5
        .ForeColor.brightness = -0.1000000015
        .Transparency = 0
        .dashStyle = msoLineSolid 'msoLineSysDash
        .Weight = 2.5
    End With
    'With Selection.Format.Fill
    '    .Visible = msoFalse
    '    .Visible = msoTrue
    '    .ForeColor.ObjectThemeColor = msoThemeColorBackground2
    '    .ForeColor.TintAndShade = 0
    '    '.ForeColor.Brightness = 0.349999994
    '    '.ForeColor.Brightness = 0
    '    .Transparency = 0
    '    .Solid
    'End With
    With Selection
        .Border.LineStyle = xlContinuous
        '.Border.Color = msoThemeColorText1
        '.MarkerBackgroundColor = msoThemeColorBackground1 'RGB(255, 255, 255)
        .MarkerForegroundColor = msoThemeColorText1 'RGB(0, 0, 0) 'RGB(255, 0, 0)
        .MarkerBackgroundColorIndex = 15
        .markerStyle = 1
        .markerSize = 6
    End With
    '-------------------------------------------------------------------------------------
    ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Line
        .Visible = msoFalse
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground2
        '.ForeColor.ObjectThemeColor = RGB(255, 0, 0)
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = -0.5
        '.ForeColor.Brightness = -0.1000000015
        '.ForeColor.Brightness = 0
        .Transparency = 0
        .dashStyle = msoLineSolid
        .Weight = 2.5
    End With
    'With Selection.Format.Fill
    '    .Visible = msoFalse
    '    .Visible = msoTrue
    '    .ForeColor.ObjectThemeColor = msoThemeColorBackground1
    '    .ForeColor.TintAndShade = 0
    '    '.ForeColor.Brightness = -0.1000000015
    '    .Solid
    '    .Transparency = 0
    'End With
    With Selection
        .Border.LineStyle = xlContinuous
        '.Border.Color = msoThemeColorBackground2
        '.Border.Color = msoThemeColorBackground2
        '.MarkerBackgroundColor = msoThemeColorBackground2
        '.MarkerForegroundColor = msoThemeColorBackground1
        .MarkerForegroundColor = msoThemeColorText1
        '.MarkerBackgroundColor = RGB(0, 255, 0)
        '.MarkerForegroundColor = RGB(255, 0, 0)
        .MarkerBackgroundColorIndex = 15
        .markerStyle = 3
        .markerSize = 6
    End With
    'ActiveChart.FullSeriesCollection(2).MarkerStyle = 4
    '-------------------------------------------------------------------------------------
    ActiveChart.FullSeriesCollection(3).Select
    With Selection.Format.Line
        .Visible = msoFalse
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
        '.ForeColor.Brightness = -0.1000000015
        '.ForeColor.Brightness = -0.349999994
        .Transparency = 0
        .dashStyle = msoLineSysDash 'msoLineDash
        .Weight = 2.5
    End With
    'With Selection.Format.Fill
    '    .Visible = msoFalse
    '    .Visible = msoTrue
    '    .ForeColor.ObjectThemeColor = msoThemeColorBackground1
    '    .ForeColor.TintAndShade = 0
    '    '.ForeColor.Brightness = -0.1000000015
    '    .Solid
    '    .Transparency = 0
    'End With
    'With Selection
    '    .MarkerStyle = 8
    '    .MarkerSize = 6
    'End With
    Selection.markerStyle = -4142
    
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Total Variation vs. N" & Chr(13) & "12 Iterations " + ChrW(955) + "=0.0001"
    'Selection.Format.TextFrame2.TextRange.Characters.Text = _
    '    "Total Variation vs. N" & Chr(13) & "12 Iterations (?=0.0001) "
    
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    ActiveChart.Axes(xlValue).HasTitle = True
    ActiveChart.Axes(xlValue).axisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).axisTitle.Text = "Total Variation (TV)"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Total Variation (TV)"
    'ActiveChart.Axes(xlValue).AxisBetweenCategories = False
    ActiveChart.Axes(xlValue).MajorTickMark = xlOutside
    ActiveChart.Axes(xlValue).MinorTickMark = xlOutside
    'Selection.MajorTickMark = xlOutside
    'Selection.MinorTickMark = xlOutside
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = 10850
    Selection.Format.Line.Visible = msoFalse
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
    End With
    ActiveChart.Axes(xlValue).HasMinorGridlines = True
    
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    ActiveChart.Axes(xlCategory).HasTitle = True
    ActiveChart.Axes(xlCategory).axisTitle.Select
    ActiveChart.Axes(xlCategory, xlPrimary).axisTitle.Text = "N"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "N"
    'ActiveChart.Axes(xlCategory).CategoryType = xlCategoryScale
    ActiveChart.Axes(xlCategory).CrossesAt = 1
    'ActiveChart.Axes(xlCategory).AxisBetweenCategories = False
    ActiveChart.Axes(xlCategory).MajorTickMark = xlOutside
    ActiveChart.Axes(xlCategory).MinorTickMark = xlOutside
    'Selection.MajorTickMark = xlOutside
    'Selection.MinorTickMark = xlOutside
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).HasMajorGridlines = True
    
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    ActiveChart.HasLegend = True
    ActiveChart.Legend.Select
    ActiveChart.Legend.Position = xlBottom
    '.Legend.Position = xlLegendPositionRight
    'Selection.Position = xlBottom
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
    End With
    ActiveChart.Axes(xlValue).Select

    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    ActiveChart.PlotArea.Select
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = 10850
    Selection.Format.Line.Visible = msoFalse
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
    End With
    ActiveChart.Axes(xlValue).HasMinorGridlines = True
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    Selection.Format.Line.Visible = msoTrue
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).HasMinorGridlines = True
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinorGridlines.Select
    Selection.Format.Line.Visible = msoFalse
    
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).HasMajorGridlines = True
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).HasMajorGridlines = True

End Sub
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************************************
Sub xyScatter()
Attribute xyScatter.VB_ProcData.VB_Invoke_Func = "S\n14"
'
' xyScatter Macro
'
' Keyboard Shortcut: Ctrl+Shift+S
'
    Dim ActiveShape As Shape
    Dim UserSelection As Variant
    Range("W22").Select
    ActiveSheet.Shapes.AddChart2(332, xlLineMarkers).Select
Dim shpTemp As Excel.Shape
    
    Set shpTemp = Application.Selection.ShapeRange.Item(1)

    'Pull-in what is selected on screen
    Set UserSelection = ActiveWindow.Selection

    'Determine if selection is a shape
    'On Error GoTo NoShapeSelected
        Set ActiveShape = ActiveSheet.Shapes(UserSelection.Name)
    'On Error Resume Next

    'Do Something with your Shape variable
    MsgBox "You have selected the shape: " & ActiveShape.Name

    ActiveSheet.Shapes(UserSelection.Name).IncrementLeft 442.9091338583
    ActiveSheet.Shapes(UserSelection.Name).IncrementTop 44.7273228346
    'ActiveSheet.Shapes("Chart 4").IncrementLeft 442.9091338583
    'ActiveSheet.Shapes("Chart 4").IncrementTop 44.7273228346
    Application.CutCopyMode = False
'----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).Name = "=""NTVS, TV Check Off"""
    ActiveChart.FullSeriesCollection(1).Values = _
        "='TVvCheck@L=0.0001(4a)'!$D$2:$D$13"
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
        .Transparency = 0
    End With
    With Selection
        .markerStyle = 8
        .markerSize = 5
    End With
    Selection.markerSize = 6
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        '.ForeColor.Brightness = 0.349999994
        .Solid
    End With
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground2
        .ForeColor.TintAndShade = 0
        '.ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
    End With
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .dashStyle = msoLineSysDash
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = 1
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = 1.75
    End With
    Selection.markerSize = 5
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = 2.5
    End With
'----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "=""NTCS, TV Check On"""
    ActiveChart.FullSeriesCollection(2).Values = _
        "='TVvCheck@L=0.0001(4a)'!$E$2:$E$13"
    ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .dashStyle = msoLineSysDash
    End With
    ActiveChart.FullSeriesCollection(2).Select
    With Selection
        .markerStyle = 8
        .markerSize = 5
    End With
    Selection.markerStyle = 1
    Selection.markerSize = 6
    Selection.markerSize = 5
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground2
        .ForeColor.TintAndShade = 0
        '.ForeColor.Brightness = 0
        .Solid
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = 1.75
    End With
    
    ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
        .Transparency = 0
    End With
    ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground2
        .ForeColor.TintAndShade = 0
        '.ForeColor.Brightness = -0.1000000015
        .Transparency = 0
        .Solid
    End With
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground2
        .ForeColor.TintAndShade = 0
        '.ForeColor.Brightness = -0.1000000015
        .Transparency = 0
        .Solid
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = 2.5
    End With
    ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = 2.5
    End With
    '----------------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------------
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(3).Name = "=""Original TVS"""
    ActiveChart.FullSeriesCollection(3).Values = _
        "='TVvCheck@L=0.0001(4a)'!$F$2:$F$13"
    ActiveChart.FullSeriesCollection(3).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
        .Transparency = 0
    End With
    Selection.markerStyle = -4142
    With Selection.Format.Line
        .Visible = msoTrue
        .dashStyle = msoLineDash
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
        .Transparency = 0
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
        .Transparency = 0.0400000215
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .dashStyle = msoLineDash
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
        .Transparency = 0
    End With
    ActiveChart.FullSeriesCollection(3).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = -0.349999994
        .Transparency = 0
    End With
    '----------------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------------
    With ActiveSheet.Shapes("Chart 4").Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
    End With
    ActiveChart.Axes(xlCategory).Select
    'With Selection.Format.TextFrame2.TextRange.Font.Fill
    '    .Visible = msoTrue
    '    .ForeColor.ObjectThemeColor = msoThemeColorText1
    '    .ForeColor.TintAndShade = 0
    '    .ForeColor.Brightness = 0.349999994
    '    .Solid
    'End With
    'Selection.Format.TextFrame2.TextRange.Font.Line.Visible = msoFalse
    ActiveChart.Axes(xlCategory).CategoryType = xlCategoryScale
    ActiveChart.Axes(xlCategory).CrossesAt = 1
    ActiveChart.Axes(xlCategory).AxisBetweenCategories = False
    ActiveChart.Axes(xlCategory).MajorTickMark = xlOutside
    ActiveChart.Axes(xlCategory).MinorTickMark = xlOutside
    ActiveChart.Axes(xlCategory).Format.Fill.Visible = msoFalse
    'With Selection.Format.Line
    '    .Visible = msoTrue
    '    .ForeColor.RGB = RGB(255, 0, 0)
    '    .Transparency = 0
    'End With
    'With Selection.Format.Line
    '    .Visible = msoTrue
    '    .ForeColor.RGB = RGB(255, 255, 0)
    '    .Transparency = 0
    'End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 0)
        .ForeColor.ObjectThemeColor = msoThemeColorBackground2
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
        .Transparency = 0
    End With
    ActiveChart.Axes(xlValue).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
        .Transparency = 0
    End With
    ActiveSheet.ChartObjects("Chart 4").Activate
    ActiveChart.PlotArea.Select
    ActiveSheet.ChartObjects("Chart 4").Activate
    ActiveSheet.ChartObjects("Chart 4").Select
    'With ActiveWorkbook.ActiveSheet.ActiveChart
    '.HasTitle = True
    '.ChartTitle.Select
   ' .ChartTitle.Text = chnam
    'VerticalAlignment (msoElementChartTitleAboveChart)          '(msoElementChartTitleAboveChart)
    ActiveChart.ChartTitle.Top = 0
    '-ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    '-ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    'ActiveChart. 'SetElement (msoElementLegendRight)
    '-ActiveChart.SetElement (msoElementDataLabelTop)
    '-ActiveChart.SetElement (msoElementDataLabelNone)
    ''ActiveSheet.ChartObjects("Chart1").Chart.HasLegend = True
    ''ActiveSheet.ChartObjects("Chart1").Chart.Legend.Position = xlBottom
    'ActiveSheet.ChartObjects("Chart1").Chart.HasLegend = True
    'ActiveSheet.ChartObjects("Chart1").Chart.Legend.Position = xlBottom
    ActiveChart.HasLegend = True
    ActiveChart.Legend.Position = xlBottom
    '.Legend.Position = xlLegendPositionRight

    ActiveChart.Legend.Select
    Selection.Position = xlBottom
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
    End With
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = 10850
    Selection.Format.Line.Visible = msoFalse
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
    End With
    Selection.MajorTickMark = xlOutside
    Selection.MinorTickMark = xlOutside
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    Selection.Format.Line.Visible = msoTrue
    ActiveChart.PlotArea.Select
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).HasMinorGridlines = True
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).HasMajorGridlines = True
    
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinorGridlines.Select
    Selection.Format.Line.Visible = msoFalse
    ActiveChart.Axes(xlCategory).HasTitle = True
    ActiveChart.Axes(xlCategory).axisTitle.Select
    ActiveChart.Axes(xlCategory, xlPrimary).axisTitle.Text = "N"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "N"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 1).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 1).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    
    ActiveChart.Axes(xlValue).HasTitle = True
    ActiveChart.Axes(xlValue).axisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).axisTitle.Text = "Total Variation (TV)"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Total Variation (TV)"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 20).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 5).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(6, 15).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Total Variation vs. N "
    Selection.Format.TextFrame2.TextRange.Characters.Text = _
        "Total Variation vs. N "
    With Selection.Format.TextFrame2.TextRange.Characters(1, 22).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 22).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    ActiveWorkbook.Save
    ActiveChart.ChartTitle.Text = "Total Variation vs. N" & Chr(13) & "12 Iterations (?=0.0001) "
    Selection.Format.TextFrame2.TextRange.Characters.Text = _
        "Total Variation vs. N" & Chr(13) & "12 Iterations (?=0.0001) "
    With Selection.Format.TextFrame2.TextRange.Characters(1, 22).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 22).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(23, 25).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(23, 13).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(36, 2).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(38, 1).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "Calibri"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "Calibri"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(39, 8).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(47, 1).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    ActiveChart.Legend.Select
    Selection.Width = 328.711
    Selection.Left = 3.144
    Selection.Width = 351.711
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Total Variation vs. N" & Chr(13) & "12 Iterations (?=0.0001) "
    ActiveChart.Axes(xlValue).axisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).axisTitle.Text = "Total Variation (TV)"
    ActiveChart.Axes(xlCategory).axisTitle.Select
    ActiveChart.Axes(xlCategory, xlPrimary).axisTitle.Text = "N"
    ActiveChart.Axes(xlCategory).Select
    
    
    ActiveChart.ChartTitle.Select
    Selection.Left = 99.402
    Selection.Top = 0
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    ActiveChart.ChartArea.Select
    ActiveChart.PlotArea.Select
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Chart 4").IncrementLeft -0.8275590551
    ActiveSheet.Shapes("Chart 4").IncrementTop -1.6551968504
    ActiveWorkbook.Save
End Sub
