### Description / Описание

This code populates a worksheet with financial data and creates a 3D bar chart to visualize the projected revenue and estimated costs over three years. It also customizes the chart's appearance, including the title, series colors, and major vertical gridlines.

Этот код заполняет рабочий лист финансовыми данными и создает 3D столбчатую диаграмму для визуализации прогнозируемых доходов и оцененных затрат за три года. Также он настраивает внешний вид диаграммы, включая заголовок, цвета серий и основные вертикальные сетки.

```vba
' VBA code equivalent to the OnlyOffice JS example
Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim cht As Chart
    Dim rng As Range
    Dim fillColor1 As Long
    Dim fillColor2 As Long

    ' Get active worksheet
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

    ' Define range for chart
    Set rng = ws.Range("A1:D3")

    ' Add a 3D Bar chart
    Set cht = ws.Shapes.AddChart2(251, xlBarClustered, 200, 100, 360, 270).Chart ' Parameters: Style, Type, Left, Top, Width, Height

    cht.SetSourceData Source:=rng

    ' Set chart title
    cht.HasTitle = True
    cht.ChartTitle.Text = "Financial Overview"
    cht.ChartTitle.Font.Size = 13

    ' Set series fill colors
    fillColor1 = RGB(51, 51, 51)      ' Dark gray
    fillColor2 = RGB(255, 111, 61)    ' Orange

    cht.SeriesCollection(1).Format.Fill.ForeColor.RGB = fillColor1
    cht.SeriesCollection(2).Format.Fill.ForeColor.RGB = fillColor2

    ' Set major vertical gridlines
    With cht.Axes(xlValue)
        .MajorGridlines.Format.Line.Weight = 1
        .MajorGridlines.Format.Line.ForeColor.RGB = RGB(255, 111, 61)
    End With
End Sub
```

```js
// This example specifies the visual properties of the major vertical gridline.
var oWorksheet = Api.GetActiveSheet();

// Set header values for years
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set labels for revenue and costs
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set financial data
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Add a 3D bar chart with specified parameters
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Create and set series fill colors
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Create and set major vertical gridlines
var oStroke = Api.CreateStroke(1 * 15000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)));
oChart.SetMajorVerticalGridlines(oStroke); 
```