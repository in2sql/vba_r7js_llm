**Description / Описание**

This code populates an Excel worksheet with financial data for the years 2014-2016, creates a 3D bar chart titled "Financial Overview," and customizes the chart's series colors and minor vertical gridlines.

Этот код заполняет рабочий лист Excel финансовыми данными за 2014-2016 годы, создает 3D столбчатую диаграмму с заголовком "Обзор Финансов" и настраивает цвета серий и дополнительные вертикальные сетки диаграммы.

```vba
' VBA Code to populate worksheet, create a 3D bar chart, set title, customize series colors, and set minor vertical gridlines

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As Chart
    Dim rng As Range
    Dim fillColor1 As Long
    Dim fillColor2 As Long
    Dim gridlineColor As Long
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Populate header years
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    
    ' Populate row labels
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    
    ' Populate financial data
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    
    ' Define the range for the chart
    Set rng = ws.Range("A1:D3")
    
    ' Add a 3D bar chart
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=360, Top:=100, Height:=270)
    Set chart = chartObj.Chart
    chart.ChartType = xl3DColumn
    
    ' Set chart data source
    chart.SetSourceData Source:=rng
    
    ' Set chart title
    chart.HasTitle = True
    chart.ChartTitle.Text = "Financial Overview"
    chart.ChartTitle.Font.Size = 13
    
    ' Define fill colors
    fillColor1 = RGB(51, 51, 51)   ' Dark Gray
    fillColor2 = RGB(255, 111, 61) ' Orange
    
    ' Customize series colors
    chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = fillColor1
    chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = fillColor2
    
    ' Define gridline color
    gridlineColor = RGB(255, 111, 61) ' Orange
    
    ' Customize minor vertical gridlines
    With chart.Axes(xlValue).MajorGridlines
        .Format.Line.Weight = 1
        .Format.Line.ForeColor.RGB = gridlineColor
    End With
End Sub
```

```javascript
// JavaScript Code using OnlyOffice API to populate worksheet, create a 3D bar chart, set title, customize series colors, and set minor vertical gridlines

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Populate header years
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Populate row labels
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Populate financial data
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Add a 3D bar chart with specific positioning and size
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set chart title and font size
oChart.SetTitle("Financial Overview", 13);

// Create and set the first series fill color (dark gray)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set the second series fill color (orange)
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Create a stroke for minor vertical gridlines with orange color
var oStroke = Api.CreateStroke(1 * 5000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)));

// Set minor vertical gridlines with the defined stroke
oChart.SetMinorVerticalGridlines(oStroke);
```