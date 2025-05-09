## Description / Описание

**English:**  
This code sets values in specific cells, creates a 3D bar chart with the specified data range, sets the chart title, positions the horizontal axis tick labels, and applies specific fill colors to the chart series.

**Russian:**  
Этот код устанавливает значения в определенные ячейки, создает 3D гистограмму с указанным диапазоном данных, задает заголовок диаграммы, позиционирует подписи основных делений горизонтальной оси и применяет определенные цвета заливки к сериям диаграммы.

### VBA Code

```vba
' VBA Code to create a 3D bar chart with specific data and formatting
Sub CreateFinancialChart()
    Dim oWorksheet As Worksheet
    Dim oChart As Chart
    Dim oSeries As Series

    ' Set the active worksheet
    Set oWorksheet = ActiveSheet

    ' Set values in cells
    oWorksheet.Range("B1").Value = 2014
    oWorksheet.Range("C1").Value = 2015
    oWorksheet.Range("D1").Value = 2016
    oWorksheet.Range("A2").Value = "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"
    oWorksheet.Range("B2").Value = 200
    oWorksheet.Range("B3").Value = 250
    oWorksheet.Range("C2").Value = 240
    oWorksheet.Range("C3").Value = 260
    oWorksheet.Range("D2").Value = 280
    oWorksheet.Range("D3").Value = 280

    ' Add a 3D Bar chart
    Set oChart = oWorksheet.Shapes.AddChart2(Style:=xlBarClustered, _
                XlChartType:=xl3DColumn, Left:=200, Top:=100, Width:=400, Height:=300).Chart
    oChart.SetSourceData Source:=oWorksheet.Range("A1:D3")

    ' Set chart title
    With oChart
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
    End With

    ' Set horizontal axis tick label position to high
    oChart.Axes(xlCategory).TickLabelPosition = xlTickLabelPositionHigh

    ' Set fill color for the first series
    Set oSeries = oChart.SeriesCollection(1)
    oSeries.Format.Fill.ForeColor.RGB = RGB(51, 51, 51)

    ' Set fill color for the second series
    Set oSeries = oChart.SeriesCollection(2)
    oSeries.Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
End Sub
```

### JavaScript Code

```javascript
// JavaScript Code to create a 3D bar chart with specific data and formatting

// Access the active worksheet
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

// Add a 3D Bar chart with specified parameters
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Set horizontal axis tick label position to high
oChart.SetHorAxisTickLabelPosition("high");

// Create and set fill color for the first series
var oFill1 = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill1, 0, false);

// Create and set fill color for the second series
var oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill2, 1, false);
```