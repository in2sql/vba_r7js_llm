### Description:
This code creates a scatter chart in a spreadsheet, sets data values, titles, axis tick marks, marker fills, and outlines.
Этот код создает точечную диаграмму в таблице, устанавливает значения данных, заголовки, отметки оси, заполнение маркеров и контуры.

```vba
' VBA Code to create a scatter chart with specific settings

Sub CreateFinancialChart()
    ' Get the active worksheet
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
    
    ' Add scatter chart
    Dim oChart As ChartObject
    Set oChart = oWorksheet.ChartObjects.Add(Left:=100, Top:=100, Width:=360, Height:=270)
    With oChart.Chart
        .ChartType = xlXYScatter
        .SetSourceData Source:=oWorksheet.Range("A1:D3")
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .Axes(xlValue).MajorGridlines.TickMark = xlTickMarkCross
        ' Set marker fill and outline
        ' Note: VBA has limited capabilities compared to OnlyOffice API for custom marker fills
    End With
End Sub
```

```javascript
// JavaScript Code to create a scatter chart with specific settings using OnlyOffice API

// Get the active worksheet
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

// Add scatter chart with specified dimensions and data range
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "scatter", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Set major tick mark for vertical axis
oChart.SetVertAxisMajorTickMark("cross");

// Create and set marker fill for first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetMarkerFill(oFill, 0, 0, true);

// Create and set marker outline for first series
var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
oChart.SetMarkerOutLine(oStroke, 0, 0, true);

// Create and set marker fill for second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetMarkerFill(oFill, 1, 0, true);

// Create and set marker outline for second series
oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)));
oChart.SetMarkerOutLine(oStroke, 1, 0, true);
```