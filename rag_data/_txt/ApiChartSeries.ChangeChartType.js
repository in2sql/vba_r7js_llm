# Description / Описание

**English:**  
This code populates specific cells in a worksheet with data, adds a combo bar and line chart, sets the chart title and series fills, changes the chart type of the first series, and records the old and new chart types in specific cells.

**Russian:**  
Этот код заполняет определённые ячейки листа данными, добавляет комбинированный столбчато-линейный график, устанавливает заголовок графика и заливку серий, изменяет тип графика первой серии и записывает старый и новый тип графика в определённые ячейки.

```vba
' Excel VBA equivalent code
Sub ModifyChart()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set values in B1, C1, D1
    oWorksheet.Range("B1").Value = 2014
    oWorksheet.Range("C1").Value = 2015
    oWorksheet.Range("D1").Value = 2016
    
    ' Set labels in A2, A3
    oWorksheet.Range("A2").Value = "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"
    
    ' Set data in B2, B3, etc.
    oWorksheet.Range("B2").Value = 200
    oWorksheet.Range("B3").Value = 250
    oWorksheet.Range("C2").Value = 240
    oWorksheet.Range("C3").Value = 260
    oWorksheet.Range("D2").Value = 280
    oWorksheet.Range("D3").Value = 280
    
    ' Add a combo bar and line chart
    Dim oChart As ChartObject
    Set oChart = oWorksheet.ChartObjects.Add(Left:=100, Top:=70, Width:=360, Height:=240)
    oChart.Chart.SetSourceData Source:=oWorksheet.Range("A1:D3")
    oChart.Chart.ChartType = xlColumnClustered ' Placeholder for combo chart
    oChart.Chart.HasTitle = True
    oChart.Chart.ChartTitle.Text = "Financial Overview"
    
    ' Set series fill colors
    oChart.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
    oChart.Chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Get the first series
    Dim oSeries As Series
    Set oSeries = oChart.Chart.SeriesCollection(1)
    
    ' Get current chart type
    Dim sSeriesType As String
    sSeriesType = oSeries.ChartType
    
    ' Record old series type in F1
    oWorksheet.Range("F1").Value = "Old Series Type = " & sSeriesType
    
    ' Change chart type to Area
    oSeries.ChartType = xlArea
    
    ' Get new chart type
    sSeriesType = oSeries.ChartType
    
    ' Record new series type in F2
    oWorksheet.Range("F2").Value = "New Series Type = " & sSeriesType
End Sub
```

```javascript
// This example populates cells with data, adds a combo bar and line chart, sets chart title and series fills,
// changes the first series chart type, and records old and new chart types in specific cells.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set values in B1, C1, D1
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set labels in A2, A3
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set data in B2, B3, etc.
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Add a combo bar and line chart with specified position and size
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "comboBarLine", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Set fill color for series 0
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Set fill color for series 1
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Get the first series
var oSeries = oChart.GetSeries(0);

// Get current chart type
var sSeriesType = oSeries.GetChartType();

// Record old series type in F1
oWorksheet.GetRange("F1").SetValue("Old Series Type = " + sSeriesType);

// Change chart type to Area
oSeries.ChangeChartType("area");

// Get new chart type
sSeriesType = oSeries.GetChartType();

// Record new series type in F2
oWorksheet.GetRange("F2").SetValue("New Series Type = " + sSeriesType);
```