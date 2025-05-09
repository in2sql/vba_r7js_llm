**Description / Описание:**

This script sets up financial data in a spreadsheet, adds a chart, customizes its appearance, and changes the chart type of a series.
Этот скрипт заполняет финансовые данные в электронной таблице, добавляет диаграмму, настраивает ее внешний вид и изменяет тип графика для серии.

```javascript
// This example changes the type of the first series of ApiChart class and inserts the new type into the document.
var oWorksheet = Api.GetActiveSheet();

// Set year values
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set category labels
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set data values
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Add a combo chart with specified dimensions and position
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "comboBarLine", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Create and set fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Get the first series and its current chart type
var oSeries = oChart.GetSeries(0);
var sSeriesType = oSeries.GetChartType();

// Set old series type value in cell F1
oWorksheet.GetRange("F1").SetValue("Old Series Type = " + sSeriesType);

// Change the chart type of the first series to 'area'
oSeries.ChangeChartType("area");

// Get the new chart type
sSeriesType = oSeries.GetChartType();

// Set new series type value in cell F2
oWorksheet.GetRange("F2").SetValue("New Series Type = " + sSeriesType);
```

```vba
' This macro sets up financial data, adds a chart, customizes its appearance, and changes the chart type of a series.
' Этот макрос заполняет финансовые данные, добавляет диаграмму, настраивает ее внешний вид и изменяет тип графика для серии.

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim cht As Chart
    Dim srs As Series
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set year values
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    
    ' Set category labels
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    
    ' Set data values
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    
    ' Add a combo chart
    Set cht = ws.Shapes.AddChart2(227, xlColumnClustered, 100, 70, 360, 180).Chart
    cht.SetSourceData Source:=ws.Range("A1:D3")
    cht.ChartType = xlColumnClustered
    
    ' Change to combo chart
    cht.ChartType = xlLineMarkers ' Placeholder, VBA requires more steps for combo charts
    
    ' Set chart title
    cht.HasTitle = True
    cht.ChartTitle.Text = "Financial Overview"
    cht.ChartTitle.Font.Size = 13
    
    ' Customize series fill colors
    Set srs = cht.SeriesCollection(1)
    srs.Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
    
    Set srs = cht.SeriesCollection(2)
    srs.Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Get the first series and its current chart type
    Set srs = cht.SeriesCollection(1)
    Dim oldType As String
    oldType = srs.ChartType
    
    ' Set old series type value in cell F1
    ws.Range("F1").Value = "Old Series Type = " & oldType
    
    ' Change the chart type of the first series to xlArea
    srs.ChartType = xlArea
    
    ' Get the new chart type
    Dim newType As String
    newType = srs.ChartType
    
    ' Set new series type value in cell F2
    ws.Range("F2").Value = "New Series Type = " & newType
End Sub
```