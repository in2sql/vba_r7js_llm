### Description / Описание
This code sets up data in an Excel worksheet, creates a 3D bar chart based on the data, and customizes the chart's title and series fill colors.
Этот код настраивает данные в рабочем листе Excel, создает 3D столбчатую диаграмму на основе данных и настраивает заголовок диаграммы и цвета заполнения серий.

```vba
' VBA Code to create and format a 3D bar chart in Excel

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As Chart
    Dim fillColor As Long
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set values in cells
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
    Set chartObj = ws.ChartObjects.Add(Left:=100, Top:=50, Width:=360, Height:=270)
    Set chart = chartObj.Chart
    chart.SetSourceData Source:=ws.Range("A1:D3")
    chart.ChartType = xl3DColumnClustered
    
    ' Set chart title
    With chart
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
    End With
    
    ' Set series fill colors
    ' First series fill color (RGB 51, 51, 51)
    fillColor = RGB(51, 51, 51)
    chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = fillColor
    
    ' Second series fill color (RGB 255, 111, 61)
    fillColor = RGB(255, 111, 61)
    chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = fillColor
    
    ' Set chart title fill color (RGB 128, 128, 128)
    fillColor = RGB(128, 128, 128)
    chart.ChartTitle.Format.Fill.ForeColor.RGB = fillColor
End Sub
```

```javascript
// JavaScript Code to create and format a 3D bar chart using OnlyOffice API

// This example sets the fill to the chart title.
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

// Add a 3D bar chart
var oChart = oWorksheet.AddChart(
    "'Sheet1'!$A$1:$D$3",
    true,
    "bar3D",
    2,
    100 * 36000,
    70 * 36000,
    0,
    2 * 36000,
    7,
    3 * 36000
);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Set series fill colors
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)); // First series
oChart.SetSeriesFill(oFill, 0, false);

oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Second series
oChart.SetSeriesFill(oFill, 1, false);

// Set chart title fill color
oFill = Api.CreateSolidFill(Api.CreateRGBColor(128, 128, 128));
oChart.SetTitleFill(oFill);
```