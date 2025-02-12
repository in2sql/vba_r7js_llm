**Description / Описание:**
This script populates an Excel worksheet with financial data, creates a 3D bar chart titled "Financial Overview," configures data labels for the chart, and sets specific fill colors for each data series.
Этот скрипт заполняет рабочий лист Excel финансовыми данными, создает 3D столбчатую диаграмму с заголовком "Financial Overview", настраивает метки данных для диаграммы и устанавливает определенные цвета заливки для каждого ряда данных.

```vba
' VBA Code to populate cells, create a chart, set the title, configure data labels, and set series fill colors

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chartRange As Range
    Dim fillColor As Long

    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Populate year headers
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016

    ' Populate row labels
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"

    ' Populate data
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280

    ' Define the range for the chart
    Set chartRange = ws.Range("'A1':'D3'")

    ' Add a 3D Bar Chart
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=360, Top:=100, Height:=270)
    With chartObj.Chart
        .ChartType = xlBarClustered ' Equivalent to "bar3D"
        .SetSourceData Source:=chartRange
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13

        ' Configure data labels
        Dim ser As Series
        For Each ser In .SeriesCollection
            ser.HasDataLabels = True
            ser.DataLabels.ShowValue = False
            ser.DataLabels.ShowCategoryName = False
            ser.DataLabels.ShowSeriesName = True
            ser.DataLabels.ShowPercentage = False
        Next ser

        ' Set series fill colors
        ' First series
        fillColor = RGB(51, 51, 51)
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = fillColor
        .SeriesCollection(1).Format.Fill.Visible = msoTrue

        ' Second series
        fillColor = RGB(255, 111, 61)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = fillColor
        .SeriesCollection(2).Format.Fill.Visible = msoTrue
    End With
End Sub
```

```javascript
// JavaScript Code to populate cells, create a chart, set the title, configure data labels, and set series fill colors using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Populate year headers
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Populate row labels
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Populate data
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Define the range for the chart and add a 3D Bar Chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set the chart title
oChart.SetTitle("Financial Overview", 13);

// Configure data labels
oChart.SetShowPointDataLabel(1, 0, false, false, true, false);

// Set fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Set fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```