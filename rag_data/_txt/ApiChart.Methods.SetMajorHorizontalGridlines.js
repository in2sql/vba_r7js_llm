## Description / Описание

**English:**  
This script populates data in an Excel sheet and creates a 3D bar chart to visualize the financial overview, including projected revenue and estimated costs for the years 2014-2016. It customizes the chart's title, series colors, and the major horizontal gridlines.

**Русский:**  
Этот скрипт заполняет данные в Excel-таблице и создает 3D-гистограмму для визуализации финансового обзора, включая прогнозируемую прибыль и оцененные расходы за годы 2014-2016. Он настраивает заголовок диаграммы, цвета серий и основные горизонтальные сетки.

```javascript
// This example specifies the visual properties of the major horizontal gridline.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
oWorksheet.GetRange("B1").SetValue(2014); // Set year 2014 in cell B1
oWorksheet.GetRange("C1").SetValue(2015); // Set year 2015 in cell C1
oWorksheet.GetRange("D1").SetValue(2016); // Set year 2016 in cell D1
oWorksheet.GetRange("A2").SetValue("Projected Revenue"); // Set label in A2
oWorksheet.GetRange("A3").SetValue("Estimated Costs"); // Set label in A3
oWorksheet.GetRange("B2").SetValue(200); // Set projected revenue for 2014
oWorksheet.GetRange("B3").SetValue(250); // Set estimated costs for 2014
oWorksheet.GetRange("C2").SetValue(240); // Set projected revenue for 2015
oWorksheet.GetRange("C3").SetValue(260); // Set estimated costs for 2015
oWorksheet.GetRange("D2").SetValue(280); // Set projected revenue for 2016
oWorksheet.GetRange("D3").SetValue(280); // Set estimated costs for 2016

// Add a 3D bar chart with specified range and properties
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);
oChart.SetTitle("Financial Overview", 13); // Set chart title with font size

// Create and set fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)); // Dark gray fill
oChart.SetSeriesFill(oFill, 0, false); // Apply fill to first series

// Create and set fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Orange fill
oChart.SetSeriesFill(oFill, 1, false); // Apply fill to second series

// Create and set stroke for major horizontal gridlines
var oStroke = Api.CreateStroke(1 * 15000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))); // Orange stroke
oChart.SetMajorHorizontalGridlines(oStroke); // Apply stroke to major horizontal gridlines
```

```vba
' This example specifies the visual properties of the major horizontal gridline.

Sub CreateFinancialOverviewChart()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet ' Get the active worksheet
    
    ' Populate headers
    ws.Range("B1").Value = 2014 ' Set year 2014 in cell B1
    ws.Range("C1").Value = 2015 ' Set year 2015 in cell C1
    ws.Range("D1").Value = 2016 ' Set year 2016 in cell D1
    
    ' Populate labels
    ws.Range("A2").Value = "Projected Revenue" ' Set label in A2
    ws.Range("A3").Value = "Estimated Costs" ' Set label in A3
    
    ' Populate data
    ws.Range("B2").Value = 200 ' Set projected revenue for 2014
    ws.Range("B3").Value = 250 ' Set estimated costs for 2014
    ws.Range("C2").Value = 240 ' Set projected revenue for 2015
    ws.Range("C3").Value = 260 ' Set estimated costs for 2015
    ws.Range("D2").Value = 280 ' Set projected revenue for 2016
    ws.Range("D3").Value = 280 ' Set estimated costs for 2016
    
    ' Add a 3D bar chart with specified range and properties
    Dim cht As Chart
    Set cht = ws.Shapes.AddChart2(251, xlBarClustered, 200, 100, 360, 270).Chart ' Add clustered bar chart
    cht.SetSourceData Source:=ws.Range("'Sheet1'!$A$1:$D$3") ' Set data range
    cht.ChartType = xl3DBarClustered ' Set chart type to 3D Bar Clustered
    
    cht.HasTitle = True
    cht.ChartTitle.Text = "Financial Overview" ' Set chart title
    cht.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 13 ' Set font size
    
    ' Set fill color for the first series
    cht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' Dark gray fill
    
    ' Set fill color for the second series
    cht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Orange fill
    
    ' Set major horizontal gridlines
    With cht.Axes(xlValue).MajorGridlines.Format.Line
        .Weight = 1 ' Set line weight
        .ForeColor.RGB = RGB(255, 111, 61) ' Set line color to orange
    End With
End Sub
```