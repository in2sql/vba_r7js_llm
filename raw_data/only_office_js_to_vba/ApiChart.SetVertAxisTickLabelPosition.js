# Description / Описание

**English:**  
This script populates a worksheet with financial data for the years 2014 to 2016, creates a 3D bar chart titled "Financial Overview," sets the vertical axis tick labels to a high position, and applies specific fill colors to the chart series.

**Russian:**  
Этот скрипт заполняет таблицу финансовыми данными за 2014–2016 годы, создает 3D-гистограмму с заголовком "Финансовый обзор", устанавливает положение меток делений вертикальной оси в верхнее положение и применяет определенные цвета заливки к сериям диаграммы.

## Excel VBA Code

```vba
' Populate worksheet with financial data and create a 3D bar chart with specific formatting

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As Chart

    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Populate headers
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016

    ' Populate row labels
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"

    ' Populate data
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280

    ' Add a 3D bar chart
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=360, Top:=100, Height:=270)
    Set chart = chartObj.Chart
    chart.SetSourceData Source:=ws.Range("A1:D3")
    chart.ChartType = xlBar3DClustered

    ' Set chart title
    chart.HasTitle = True
    chart.ChartTitle.Text = "Financial Overview"

    ' Set vertical axis tick label position to high
    chart.Axes(xlValue).TickLabelPosition = xlTickLabelPositionHigh

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

## OnlyOffice JavaScript Code

```javascript
// Populate the worksheet with financial data and create a 3D bar chart with specific formatting

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Populate headers
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Populate row labels
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Populate data
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Add a 3D bar chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Set vertical axis tick label position to high
oChart.SetVertAxisTickLabelPosition("high");

// Set series fill colors
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```