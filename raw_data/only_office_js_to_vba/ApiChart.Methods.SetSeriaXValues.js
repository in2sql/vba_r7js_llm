### Description / Описание
This code sets up data in an Excel worksheet and creates a scatter chart with customized markers and titles.
Этот код устанавливает данные в рабочем листе Excel и создает диаграмму рассеяния с настраиваемыми маркерами и заголовками.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set values for the years
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set labels for the data series
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set data for the first series
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);

// Set value for the next year
oWorksheet.GetRange("B4").SetValue(2017);

// Set data for the second series
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);

// Set value for the next year
oWorksheet.GetRange("C4").SetValue(2018);

// Set data for the third series
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Set value for the next year
oWorksheet.GetRange("D4").SetValue(2019);

// Add a scatter chart to the worksheet
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "scatter", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);

// Set the title of the chart
oChart.SetTitle("Financial Overview", 13);

// Set the X-axis values for the series
oChart.SetSeriaXValues("'Sheet1'!$B$4:$D$4", 0);

// Create a solid fill color for the first marker
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetMarkerFill(oFill, 0, 0, true);

// Create a stroke for the first marker outline
var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
oChart.SetMarkerOutLine(oStroke, 0, 0, true);

// Create a solid fill color for the second marker
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetMarkerFill(oFill, 1, 0, true);

// Create a stroke for the second marker outline
oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)));
oChart.SetMarkerOutLine(oStroke, 1, 0, true);
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Set values for the years
oWorksheet.Range("B1").Value = 2014
oWorksheet.Range("C1").Value = 2015
oWorksheet.Range("D1").Value = 2016

' Set labels for the data series
oWorksheet.Range("A2").Value = "Projected Revenue"
oWorksheet.Range("A3").Value = "Estimated Costs"

' Set data for the first series
oWorksheet.Range("B2").Value = 200
oWorksheet.Range("B3").Value = 250

' Set value for the next year
oWorksheet.Range("B4").Value = 2017

' Set data for the second series
oWorksheet.Range("C2").Value = 240
oWorksheet.Range("C3").Value = 260

' Set value for the next year
oWorksheet.Range("C4").Value = 2018

' Set data for the third series
oWorksheet.Range("D2").Value = 280
oWorksheet.Range("D3").Value = 280

' Set value for the next year
oWorksheet.Range("D4").Value = 2019

' Add a scatter chart to the worksheet
Dim oChart As Chart
Set oChart = oWorksheet.Shapes.AddChart2(201, xlXYScatter).Chart

' Set the data source for the chart
oChart.SetSourceData Source:=oWorksheet.Range("A1:D3")

' Set the title of the chart
oChart.HasTitle = True
oChart.ChartTitle.Text = "Financial Overview"

' Set the X-axis values for the first series
oChart.SeriesCollection(1).XValues = oWorksheet.Range("B4:D4")

' Customize the first series markers
With oChart.SeriesCollection(1).Marker
    .Style = xlMarkerStyleCircle
    .BackgroundColor = RGB(51, 51, 51)
    .ForegroundColor = RGB(51, 51, 51)
    .Size = 10
End With

' Customize the second series markers
With oChart.SeriesCollection(2).Marker
    .Style = xlMarkerStyleCircle
    .BackgroundColor = RGB(255, 111, 61)
    .ForegroundColor = RGB(255, 111, 61)
    .Size = 10
End With
```