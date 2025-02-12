# Code Description / Описание кода

This code sets cell values, creates a 3D bar chart, sets the chart title, applies colors to chart series, and outlines the chart legend.

Этот код устанавливает значения ячеек, создает 3D столбчатую диаграмму, устанавливает заголовок диаграммы, применяет цвета к сериям диаграммы и очерчивает легенду диаграммы.

```vba
' VBA Code

Sub CreateFinancialOverviewChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As Chart

    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Set cell values
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280

    ' Add a 3D bar chart
    Set chartObj = ws.ChartObjects.Add(Left:=100, Top:=70, Width:=360, Height:=360)
    Set chart = chartObj.Chart
    chart.SetSourceData Source:=ws.Range("A1:D3")
    chart.ChartType = xlBar3DClustered

    ' Set chart title
    chart.HasTitle = True
    chart.ChartTitle.Text = "Financial Overview"
    chart.ChartTitle.Font.Size = 13

    ' Set series fill colors
    chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' Dark gray
    chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Orange

    ' Outline the legend
    With chart.Legend.Format.Line
        .Visible = msoTrue
        .Weight = 0.5
        .ForeColor.RGB = RGB(51, 51, 51)
    End With
End Sub
```

```javascript
// JavaScript Code

// This example sets the outline to the chart legend.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
oWorksheet.GetRange("B1").SetValue(2014); // Set cell B1 to 2014
oWorksheet.GetRange("C1").SetValue(2015); // Set cell C1 to 2015
oWorksheet.GetRange("D1").SetValue(2016); // Set cell D1 to 2016
oWorksheet.GetRange("A2").SetValue("Projected Revenue"); // Set cell A2
oWorksheet.GetRange("A3").SetValue("Estimated Costs"); // Set cell A3
oWorksheet.GetRange("B2").SetValue(200); // Set cell B2
oWorksheet.GetRange("B3").SetValue(250); // Set cell B3
oWorksheet.GetRange("C2").SetValue(240); // Set cell C2
oWorksheet.GetRange("C3").SetValue(260); // Set cell C3
oWorksheet.GetRange("D2").SetValue(280); // Set cell D2
oWorksheet.GetRange("D3").SetValue(280); // Set cell D3

// Add a 3D bar chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Set series fill colors
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)); // Dark gray
oChart.SetSeriesFill(oFill, 0, false);
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Orange
oChart.SetSeriesFill(oFill, 1, false);

// Outline the legend
var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
oChart.SetLegendOutLine(oStroke);
```