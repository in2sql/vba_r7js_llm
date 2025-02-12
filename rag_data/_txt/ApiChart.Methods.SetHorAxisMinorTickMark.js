## Description / Описание

This code sets values in specific cells, adds a scatter chart with specific properties, sets the chart title and axis minor tick marks, and customizes marker fills and outlines.

Этот код устанавливает значения в определенные ячейки, добавляет точечную диаграмму с определенными свойствами, устанавливает заголовок диаграммы и отметки на осях, а также настраивает заполнение и контуры маркеров.

```vba
' VBA code equivalent
Sub CreateFinancialOverviewChart()
    ' Get the active sheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set values in cells
    oWorksheet.Range("B1").Value = 2014
    oWorksheet.Range("C1").Value = 2015
    oWorksheet.Range("D1").Value = 2016
    oWorksheet.Range("A2").Value = "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"
    oWorksheet.Range("B2").Value = 200
    oWorksheet.Range("B3").Value = 250
    oWorksheet.Range("C2").Value = 240
    oWorksheet.Range("C3").Value = 260
    oWorksheet.Range("D2").Value = 280
    oWorksheet.Range("D3").Value = 280
    
    ' Add a scatter chart
    Dim oChart As Chart
    Set oChart = oWorksheet.Shapes.AddChart2(201, xlXYScatter).Chart
    oChart.SetSourceData Source:=oWorksheet.Range("A1:D3")
    
    ' Set chart title
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "Financial Overview"
    oChart.ChartTitle.Font.Size = 13
    
    ' Set minor tick marks for horizontal axis
    With oChart.Axes(xlCategory)
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkOutside
    End With
    
    ' Customize marker fills and outlines for first series
    With oChart.SeriesCollection(1).Format.Fill
        .ForeColor.RGB = RGB(51, 51, 51)
        .Solid
    End With
    With oChart.SeriesCollection(1).Format.Line
        .Weight = 0.5
        .ForeColor.RGB = RGB(51, 51, 51)
    End With
    
    ' Customize marker fills and outlines for second series
    With oChart.SeriesCollection(2).Format.Fill
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
    With oChart.SeriesCollection(2).Format.Line
        .Weight = 0.5
        .ForeColor.RGB = RGB(255, 111, 61)
    End With
End Sub
```

```javascript
// JavaScript code using OnlyOffice API

// This example specifies the minor tick mark for the horizontal axis.
var oWorksheet = Api.GetActiveSheet();

// Set values in cells
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Add a scatter chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "scatter", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Set minor tick marks for horizontal axis
oChart.SetHorAxisMinorTickMark("out");

// Customize marker fills and outlines
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetMarkerFill(oFill, 0, 0, true);

var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
oChart.SetMarkerOutLine(oStroke, 0, 0, true);

oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetMarkerFill(oFill, 1, 0, true);

oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)));
oChart.SetMarkerOutLine(oStroke, 1, 0, true);
```