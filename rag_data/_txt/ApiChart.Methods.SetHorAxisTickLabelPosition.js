### Description / Описание
**English:** This script populates a worksheet with financial data for the years 2014 to 2016, adds a 3D bar chart titled "Financial Overview," sets the horizontal axis tick label position to high, and applies specific fill colors to the chart series.

**Russian:** Этот скрипт заполняет рабочий лист финансовыми данными за 2014-2016 годы, добавляет 3D-столбчатую диаграмму с заголовком "Финансовый обзор", устанавливает положение меток делений горизонтальной оси на высокий уровень и применяет определенные цвета заливки к сериям диаграммы.

```vba
' Excel VBA equivalent code

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As Chart
    Dim rng As Range
    
    ' Set active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Populate data
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
    
    ' Define data range
    Set rng = ws.Range("A1:D3")
    
    ' Add a 3D Bar chart
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=360, Top:=100, Height:=270)
    Set chart = chartObj.Chart
    chart.SetSourceData Source:=rng
    chart.ChartType = xlBar3DClustered
    
    ' Set chart title
    chart.HasTitle = True
    chart.ChartTitle.Text = "Financial Overview"
    chart.ChartTitle.Font.Size = 13
    
    ' Set horizontal axis tick label position to high
    chart.Axes(xlCategory).TickLabelPosition = xlTickLabelPositionHigh
    
    ' Set series fill colors
    With chart.SeriesCollection(1).Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(51, 51, 51)
        .Solid
    End With
    
    With chart.SeriesCollection(2).Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61)
        .Solid
    End With
End Sub
```

```javascript
// OnlyOffice JS equivalent code

// Set the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Populate data
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

// Add a 3D Bar chart with specified dimensions and position
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set chart title and font size
oChart.SetTitle("Financial Overview", 13);

// Set horizontal axis tick label position to high
oChart.SetHorAxisTickLabelPosition("high");

// Create and set fill for first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set fill for second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```