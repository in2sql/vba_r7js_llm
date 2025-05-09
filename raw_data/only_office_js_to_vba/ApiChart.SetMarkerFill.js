**Description / Описание**

This code sets up data in an Excel worksheet, creates a scatter chart titled "Financial Overview," and customizes the marker fills and outlines for two data series.

Этот код заполняет данные в рабочем листе Excel, создает диаграмму типа "Точечная" с заголовком "Финансовый обзор" и настраивает заливку маркеров и контуры для двух серий данных.

```javascript
// This example sets the fill to the marker in the specified chart series.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
oWorksheet.GetRange("B1").SetValue(2014); // Set value 2014 in cell B1
oWorksheet.GetRange("C1").SetValue(2015); // Set value 2015 in cell C1
oWorksheet.GetRange("D1").SetValue(2016); // Set value 2016 in cell D1
oWorksheet.GetRange("A2").SetValue("Projected Revenue"); // Set label in A2
oWorksheet.GetRange("A3").SetValue("Estimated Costs"); // Set label in A3
oWorksheet.GetRange("B2").SetValue(200); // Set value 200 in B2
oWorksheet.GetRange("B3").SetValue(250); // Set value 250 in B3
oWorksheet.GetRange("C2").SetValue(240); // Set value 240 in C2
oWorksheet.GetRange("C3").SetValue(260); // Set value 260 in C3
oWorksheet.GetRange("D2").SetValue(280); // Set value 280 in D2
oWorksheet.GetRange("D3").SetValue(280); // Set value 280 in D3
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "scatter", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000); // Add scatter chart
oChart.SetTitle("Financial Overview", 13); // Set chart title with font size 13
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)); // Create dark gray fill
oChart.SetMarkerFill(oFill, 0, 0, true); // Set marker fill for first series
var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51))); // Create dark gray stroke
oChart.SetMarkerOutLine(oStroke, 0, 0, true); // Set marker outline for first series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create orange fill
oChart.SetMarkerFill(oFill, 1, 0, true); // Set marker fill for second series
oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))); // Create orange stroke
oChart.SetMarkerOutLine(oStroke, 1, 0, true); // Set marker outline for second series
```

```vba
' This VBA code sets up data in the active worksheet, creates a scatter chart titled "Financial Overview",
' and customizes the marker fills and outlines for two data series.

Sub CreateFinancialOverviewChart()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet ' Get the active worksheet

    ' Set headers
    oWorksheet.Range("B1").Value = 2014
    oWorksheet.Range("C1").Value = 2015
    oWorksheet.Range("D1").Value = 2016
    oWorksheet.Range("A2").Value = "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"

    ' Set data
    oWorksheet.Range("B2").Value = 200
    oWorksheet.Range("B3").Value = 250
    oWorksheet.Range("C2").Value = 240
    oWorksheet.Range("C3").Value = 260
    oWorksheet.Range("D2").Value = 280
    oWorksheet.Range("D3").Value = 280

    ' Add scatter chart
    Dim oChart As Chart
    Set oChart = oWorksheet.Shapes.AddChart2(201, xlXYScatter).Chart
    oChart.SetSourceData Source:=oWorksheet.Range("A1:D3")
    oChart.ChartTitle.Text = "Financial Overview"
    oChart.ChartTitle.Font.Size = 13

    ' Customize first series markers
    With oChart.SeriesCollection(1)
        .MarkerBackgroundColor = RGB(51, 51, 51) ' Dark gray fill
        .MarkerForegroundColor = RGB(51, 51, 51) ' Dark gray outline
        .MarkerSize = 7
        .MarkerStyle = xlMarkerStyleCircle
    End With

    ' Customize second series markers
    With oChart.SeriesCollection(2)
        .MarkerBackgroundColor = RGB(255, 111, 61) ' Orange fill
        .MarkerForegroundColor = RGB(255, 111, 61) ' Orange outline
        .MarkerSize = 7
        .MarkerStyle = xlMarkerStyleCircle
    End With
End Sub
```