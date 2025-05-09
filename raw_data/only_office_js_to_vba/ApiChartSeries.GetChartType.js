**Description / Описание:**
This code populates an Excel worksheet with financial data, creates a combination bar and line chart, customizes the chart's appearance, and displays the chart type of the first series in cell F1.

Этот код заполняет рабочий лист Excel финансовыми данными, создает комбинированную колонную и линейную диаграмму, настраивает внешний вид диаграммы и отображает тип диаграммы первого ряда в ячейке F1.

```vba
' VBA Code to replicate OnlyOffice API functionality

Sub CreateFinancialChart()
    Dim oWorksheet As Worksheet
    Dim oChart As Chart
    Dim oSeries As Series
    Dim sSeriesType As String

    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

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

    ' Add a combination chart (Column and Line)
    Set oChart = oWorksheet.Shapes.AddChart2(240, xlColumnClustered, 100, 70, 300, 200).Chart
    oChart.SetSourceData Source:=oWorksheet.Range("'Sheet1'!$A$1:$D$3")
    oChart.ChartType = xlColumnClustered

    ' Change chart type to combination
    oChart.ChartType = xlColumnClustered
    oChart.SeriesCollection.NewSeries
    oChart.SeriesCollection(2).ChartType = xlLine
    oChart.SeriesCollection(2).Name = "Line Series"

    ' Set chart title
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "Financial Overview"

    ' Customize the first series fill color to RGB(51, 51, 51)
    oChart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)

    ' Customize the second series fill color to RGB(255, 111, 61)
    oChart.SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(255, 111, 61)

    ' Get the chart type of the first series
    sSeriesType = oChart.SeriesCollection(1).ChartType

    ' Set the series type information in cell F1
    oWorksheet.Range("F1").Value = "Series Type = " & sSeriesType
End Sub
```

```javascript
// JavaScript Code to replicate OnlyOffice API functionality

// Get the active worksheet
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

// Add a combination chart (Column and Line)
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "comboBarLine", 2, 100 * 36000, 70 * 36000, 300 * 36000, 200 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Create and set fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Get the chart type of the first series
var oSeries = oChart.GetSeries(0);
var sSeriesType = oSeries.GetChartType();

// Set the series type information in cell F1
oWorksheet.GetRange("F1").SetValue("Series Type = " + sSeriesType);
```