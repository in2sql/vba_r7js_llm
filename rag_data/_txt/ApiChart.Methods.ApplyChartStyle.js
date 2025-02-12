**Description / Описание**

This script sets up financial data in an Excel worksheet, creates a 3D bar chart with customized styles, and applies specific fill and outline colors to the chart series.

Этот скрипт настраивает финансовые данные в рабочем листе Excel, создает 3D столбчатую диаграмму с настраиваемыми стилями и применяет определенные цвета заливки и контура к сериям диаграммы.

```vba
' VBA Code to replicate the OnlyOffice JS functionality

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As Chart
    Dim fillColor1 As Long
    Dim fillColor2 As Long
    Dim strokeColor1 As Long
    Dim strokeColor2 As Long
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Populate the cells with data
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
    
    ' Add a 3D Bar Chart
    Set chartObj = ws.ChartObjects.Add(Left:=100, Width:=360, Top:=70, Height:=270)
    Set chart = chartObj.Chart
    chart.SetSourceData Source:=ws.Range("A1:D3")
    chart.ChartType = xl3DBarClustered
    
    ' Set chart title
    With chart
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
    End With
    
    ' Apply chart style
    chart.ApplyStyle (2)
    
    ' Define colors
    fillColor1 = RGB(51, 51, 51)       ' Dark Gray
    fillColor2 = RGB(255, 111, 61)     ' Orange
    strokeColor1 = RGB(51, 51, 51)     ' Dark Gray
    strokeColor2 = RGB(255, 111, 61)   ' Orange
    
    ' Set fill and outline for first series
    With chart.SeriesCollection(1).Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor1
        .Solid
    End With
    With chart.SeriesCollection(1).Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = strokeColor1
        .Weight = 0.5
    End With
    
    ' Set fill and outline for second series
    With chart.SeriesCollection(2).Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor2
        .Solid
    End With
    With chart.SeriesCollection(2).Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = strokeColor2
        .Weight = 0.5
    End With
End Sub
```

```javascript
// This example sets a style to the chart by style ID.
var oWorksheet = Api.GetActiveSheet();
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

// Add a 3D Bar Chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Apply chart style
oChart.ApplyChartStyle(2);

// Define fill and stroke for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);
var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
oChart.SetSeriesOutLine(oStroke, 0, false);

// Define fill and stroke for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)));
oChart.SetSeriesOutLine(oStroke, 1, false); 
```