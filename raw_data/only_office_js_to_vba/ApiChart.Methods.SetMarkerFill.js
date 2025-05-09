# Code Description / Описание кода

This code populates an Excel worksheet with financial data, creates a scatter chart titled "Financial Overview," and customizes the marker fills and outlines for the chart series.

Этот код заполняет лист Excel финансовыми данными, создает точечную диаграмму с заголовком "Financial Overview" и настраивает заливку и контуры маркеров для серий диаграммы.

```vba
' VBA code to create a financial overview chart with customized markers

Sub CreateFinancialOverviewChart()
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

    ' Add a scatter chart
    Dim oChart As Chart
    Set oChart = oWorksheet.Shapes.AddChart2(227, xlXYScatter, 100, 70, 360, 240).Chart
    oChart.SetSourceData Source:=oWorksheet.Range("A1:D3")
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "Financial Overview"

    ' Customize first series marker fill and outline
    With oChart.SeriesCollection(1)
        .MarkerBackgroundColor = RGB(51, 51, 51)
        .MarkerForegroundColor = RGB(51, 51, 51)
        .MarkerSize = 7
        .Format.Line.Weight = 0.5
        .Format.Line.ForeColor.RGB = RGB(51, 51, 51)
    End With

    ' Customize second series marker fill and outline
    With oChart.SeriesCollection(2)
        .MarkerBackgroundColor = RGB(255, 111, 61)
        .MarkerForegroundColor = RGB(255, 111, 61)
        .MarkerSize = 7
        .Format.Line.Weight = 0.5
        .Format.Line.ForeColor.RGB = RGB(255, 111, 61)
    End With
End Sub
```

```javascript
// JavaScript code to create a financial overview chart with customized markers

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

// Add a scatter chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "scatter", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);
oChart.SetTitle("Financial Overview", 13);

// Customize first series marker fill and outline
var oFill1 = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetMarkerFill(oFill1, 0, 0, true);
var oStroke1 = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
oChart.SetMarkerOutLine(oStroke1, 0, 0, true);

// Customize second series marker fill and outline
var oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetMarkerFill(oFill2, 1, 0, true);
var oStroke2 = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)));
oChart.SetMarkerOutLine(oStroke2, 1, 0, true);
```