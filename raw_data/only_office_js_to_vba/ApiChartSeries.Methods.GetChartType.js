# Description / Описание

**English**: This code manipulates an Excel worksheet by setting specific cell values, adds a combo bar-line chart with specified dimensions and series, sets the chart title, customizes the series fill colors, retrieves the chart type of the first series, and outputs the series type into a cell.

**Russian**: Этот код манипулирует рабочим листом Excel, устанавливая определенные значения ячеек, добавляет комбинированный столбцово-линейный график с заданными размерами и сериями, устанавливает заголовок графика, настраивает цвета заливки серий, получает тип графика первой серии и выводит тип серии в ячейку.

```vba
' VBA Code: Manipulate worksheet and add a combo chart

Sub CreateFinancialOverviewComboChart()
    Dim oWorksheet As Worksheet
    Dim oChartObj As ChartObject
    Dim oChart As Chart
    Dim sSeriesType As String
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set cell values for years
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
    Set oChartObj = oWorksheet.ChartObjects.Add(Left:=200, Top:=100, Width:=360, Height:=70)
    Set oChart = oChartObj.Chart
    oChart.SetSourceData Source:=oWorksheet.Range("A1:D3")
    oChart.ChartType = xlColumnClustered
    
    ' Change the second series to a line chart
    oChart.SeriesCollection(2).ChartType = xlLine
    
    ' Set chart title
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "Financial Overview"
    oChart.ChartTitle.Font.Size = 13
    
    ' Customize series fill colors
    oChart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' First series
    oChart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Second series
    
    ' Get chart type of the first series
    sSeriesType = oChart.SeriesCollection(1).ChartType
    
    ' Output series type to cell F1
    oWorksheet.Range("F1").Value = "Series Type = " & sSeriesType
End Sub
```

```javascript
// JavaScript Code: Manipulate worksheet and add a combo chart

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set years in cells B1, C1, D1
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set labels in column A
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set data values
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Add a combo bar-line chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "comboBarLine", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Customize series fill colors
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Get chart type of the first series
var oSeries = oChart.GetSeries(0);
var sSeriesType = oSeries.GetChartType();

// Output series type to cell F1
oWorksheet.GetRange("F1").SetValue("Series Type = " + sSeriesType);
```