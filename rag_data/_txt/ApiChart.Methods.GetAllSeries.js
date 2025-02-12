### Description / Описание
**English:** This code initializes data on a worksheet, creates a combo bar-line chart, sets its title and series colors, and writes the chart series types to column F.

**Русский:** Этот код инициализирует данные на листе, создает комбинированную столбцовую и линейную диаграмму, устанавливает заголовок и цвета серий, а также записывает типы серий диаграммы в столбец F.

```javascript
// JavaScript Code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set values for years
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set labels
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set data values
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Add a combo chart to the worksheet
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "comboBarLine", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Set series fill colors
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Get all series and write their types to column F
var aAllSeries = oChart.GetAllSeries();
var oSeries, sSeriesType;
for(var nSeries = 0; nSeries < aAllSeries.length; ++nSeries) {
    oSeries = aAllSeries[nSeries];
    sSeriesType = oSeries.GetChartType();
    oWorksheet.GetRange("F" + (nSeries + 1)).SetValue((nSeries + 1) + " Series Type = " + sSeriesType);
}
```

```vba
' VBA Code

Sub CreateFinancialOverviewChart()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set values for years
    oWorksheet.Range("B1").Value = 2014
    oWorksheet.Range("C1").Value = 2015
    oWorksheet.Range("D1").Value = 2016
    
    ' Set labels
    oWorksheet.Range("A2").Value = "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"
    
    ' Set data values
    oWorksheet.Range("B2").Value = 200
    oWorksheet.Range("B3").Value = 250
    oWorksheet.Range("C2").Value = 240
    oWorksheet.Range("C3").Value = 260
    oWorksheet.Range("D2").Value = 280
    oWorksheet.Range("D3").Value = 280
    
    ' Add a combo chart
    Dim oChart As Chart
    Set oChart = oWorksheet.Shapes.AddChart2(201, xlColumnClustered, 100, 70, 300, 200).Chart
    oChart.SetSourceData Source:=oWorksheet.Range("A1:D3")
    oChart.ChartType = xlColumnClustered ' Adjust as needed for combo
    
    ' Set chart title
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "Financial Overview"
    
    ' Set series fill colors
    oChart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
    oChart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Write series types to column F
    Dim i As Integer
    For i = 1 To oChart.SeriesCollection.Count
        oWorksheet.Range("F" & i).Value = i & " Series Type = " & oChart.SeriesCollection(i).ChartType
    Next i
End Sub
```