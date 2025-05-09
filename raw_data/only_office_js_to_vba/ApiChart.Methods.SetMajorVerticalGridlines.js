### Description / Описание
**English:** This script populates specific cells with data, creates a 3D bar chart titled "Financial Overview," sets custom fill colors for the chart series, and defines the major vertical gridlines' stroke properties.

**Russian:** Этот скрипт заполняет определенные ячейки данными, создает 3D столбчатую диаграмму с заголовком "Финансовый обзор", устанавливает пользовательские цвета заливки для серий диаграммы и определяет свойства штриха основных вертикальных сеток.

---

#### VBA Code
```vba
' This VBA script populates cells, creates a 3D bar chart, sets chart title,
' custom series fill colors, and defines major vertical gridlines.

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As Chart
    Dim series1 As Series, series2 As Series
    Dim gridlines As Gridlines

    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Populate cells with data
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
    Set chartObj = ws.ChartObjects.Add(Left:=200, Top:=100, Width:=360, Height:=270)
    Set chart = chartObj.Chart
    chart.SetSourceData Source:=ws.Range("'Sheet1'!$A$1:$D$3")
    chart.ChartType = xlBarClustered ' Equivalent to "bar3D"

    ' Set chart title
    With chart
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
    End With

    ' Set series fill colors
    Set series1 = chart.SeriesCollection(1)
    series1.Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' Dark Gray

    Set series2 = chart.SeriesCollection(2)
    series2.Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Orange

    ' Set major vertical gridlines
    With chart.Axes(xlValue).MajorGridLines
        .Format.Line.Weight = 1.5
        .Format.Line.ForeColor.RGB = RGB(255, 111, 61) ' Orange
    End With
End Sub
```

#### JavaScript Code
```javascript
// This script populates cells, creates a 3D bar chart, sets chart title,
// custom series fill colors, and defines major vertical gridlines.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Populate cells with data
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
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Set series fill colors
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)); // Dark Gray
oChart.SetSeriesFill(oFill, 0, false);

oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Orange
oChart.SetSeriesFill(oFill, 1, false);

// Set major vertical gridlines
var oStroke = Api.CreateStroke(1 * 15000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))); // Orange stroke
oChart.SetMajorVerticalGridlines(oStroke);
```