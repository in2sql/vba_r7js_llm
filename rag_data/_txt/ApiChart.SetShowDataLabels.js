### Description / Описание
This code sets up a worksheet with specific data, creates a 3D bar chart, customizes its title and data labels, and applies specific fill colors to the chart series.
Этот код настраивает лист с определёнными данными, создаёт 3D столбчатую диаграмму, настраивает её заголовок и подписи данных, а также применяет специальные цвета заливки к сериям диаграммы.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set values for the header row
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set labels for the first column
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set data for the chart
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Add a 3D bar chart to the worksheet
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set the title of the chart
oChart.SetTitle("Financial Overview", 13);

// Configure data labels for the chart
oChart.SetShowDataLabels(false, false, true, false);

// Create and set the fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set the fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ActiveSheet

' Set values for the header row
oWorksheet.Range("B1").Value = 2014
oWorksheet.Range("C1").Value = 2015
oWorksheet.Range("D1").Value = 2016

' Set labels for the first column
oWorksheet.Range("A2").Value = "Projected Revenue"
oWorksheet.Range("A3").Value = "Estimated Costs"

' Set data for the chart
oWorksheet.Range("B2").Value = 200
oWorksheet.Range("B3").Value = 250
oWorksheet.Range("C2").Value = 240
oWorksheet.Range("C3").Value = 260
oWorksheet.Range("D2").Value = 280
oWorksheet.Range("D3").Value = 280

' Add a 3D bar chart to the worksheet
Dim oChart As ChartObject
Set oChart = oWorksheet.ChartObjects.Add(Left:=100, Top:=70, Width:=360, Height:=270)
oChart.Chart.ChartType = xlBarClustered ' VBA does not have a direct 3D bar type

' Set the source data for the chart
oChart.Chart.SetSourceData Source:=oWorksheet.Range("A1:D3")

' Set the title of the chart
oChart.Chart.HasTitle = True
oChart.Chart.ChartTitle.Text = "Financial Overview"

' Configure data labels for the chart
Dim ser As Series
For Each ser In oChart.Chart.SeriesCollection
    ser.HasDataLabels = True
    ser.DataLabels.ShowValue = True
    ' VBA does not support partial data label display like JS
Next ser

' Set the fill color for the first series
oChart.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)

' Set the fill color for the second series
oChart.Chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
```