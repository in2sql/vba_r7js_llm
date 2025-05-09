# Description / Описание
This code populates an Excel worksheet with financial data, creates a scatter chart based on that data, sets the chart title, and customizes the marker fill and outline for the chart series.

Этот код заполняет рабочий лист Excel финансовыми данными, создает точечную диаграмму на основе этих данных, устанавливает заголовок диаграммы и настраивает заполнение маркеров и их контуры для серий диаграммы.

---

## Excel VBA Code

```vba
' This macro populates the worksheet with data, creates a scatter chart,
' sets the chart title, and customizes marker fills and outlines.

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As Chart
    Dim fillColor1 As Long
    Dim fillColor2 As Long
    Dim strokeColor As Long
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Populate header years
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    
    ' Populate category labels
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    
    ' Populate financial data
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    
    ' Add a scatter chart
    Set chartObj = ws.ChartObjects.Add(Left:=100, Top:=70, Width:=300, Height:=200)
    Set chart = chartObj.Chart
    chart.SetSourceData Source:=ws.Range("A1:D3")
    chart.ChartType = xlXYScatter
    
    ' Set chart title
    chart.HasTitle = True
    chart.ChartTitle.Text = "Financial Overview"
    
    ' Define colors
    fillColor1 = RGB(51, 51, 51)      ' Dark gray
    fillColor2 = RGB(255, 111, 61)    ' Orange
    strokeColor = RGB(51, 51, 51)     ' Dark gray
    
    ' Customize marker for first series
    With chart.SeriesCollection(1).Points(1)
        .MarkerBackgroundColor = fillColor1
        .MarkerForegroundColor = strokeColor
        .MarkerSize = 7
        .Border.Weight = 0.5
    End With
    
    ' Customize marker for second series
    With chart.SeriesCollection(2).Points(1)
        .MarkerBackgroundColor = fillColor2
        .MarkerForegroundColor = strokeColor
        .MarkerSize = 7
        .Border.Weight = 0.5
    End With
End Sub
```

---

## OnlyOffice JavaScript Code

```javascript
// This example sets the outline to the marker in the specified chart series.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Populate header years
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Populate category labels
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Populate financial data
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Add a scatter chart with specified position and size
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "scatter", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Create and set marker fill for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetMarkerFill(oFill, 0, 0, true);

// Create stroke for marker outline
var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
oChart.SetMarkerOutLine(oStroke, 0, 0, true);

// Create and set marker fill for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetMarkerFill(oFill, 1, 0, true);
oChart.SetMarkerOutLine(oStroke, 1, 0, true);
```