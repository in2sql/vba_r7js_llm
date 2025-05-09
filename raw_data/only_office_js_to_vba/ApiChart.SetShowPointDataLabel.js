## Description / Описание

This code populates a worksheet with financial data, creates a 3D bar chart titled "Financial Overview," customizes data labels, and sets specific fill colors for each data series.

Этот код заполняет рабочий лист финансовыми данными, создает 3D-гистограмму с заголовком "Финансовый обзор", настраивает метки данных и устанавливает определенные цвета заливки для каждого ряда данных.

```vba
' VBA Code to populate data, create a 3D bar chart, and customize its appearance

Sub CreateFinancialChart()
    Dim oWorksheet As Worksheet
    Dim oChart As Chart
    Dim oFill As Object
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Populate cells with years
    oWorksheet.Range("B1").Value = 2014
    oWorksheet.Range("C1").Value = 2015
    oWorksheet.Range("D1").Value = 2016
    
    ' Populate labels
    oWorksheet.Range("A2").Value = "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"
    
    ' Populate financial data
    oWorksheet.Range("B2").Value = 200
    oWorksheet.Range("B3").Value = 250
    oWorksheet.Range("C2").Value = 240
    oWorksheet.Range("C3").Value = 260
    oWorksheet.Range("D2").Value = 280
    oWorksheet.Range("D3").Value = 280
    
    ' Add a 3D bar chart
    Set oChart = oWorksheet.Shapes.AddChart2(201, xlBarClustered, 200, 100, 360, 270).Chart
    oChart.SetSourceData Source:=oWorksheet.Range("'Sheet1'!$A$1:$D$3")
    oChart.ChartType = xl3DBarClustered
    
    ' Set chart title
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "Financial Overview"
    oChart.ChartTitle.Font.Size = 13
    
    ' Customize data labels
    Dim series As Series
    For Each series In oChart.SeriesCollection
        series.HasDataLabels = True
        With series.DataLabels
            .ShowValue = True
            .ShowSeriesName = False
            .ShowCategoryName = False
            .ShowPercentage = False
            .ShowBubbleSize = False
        End With
    Next series
    
    ' Set series fill colors
    oChart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' Dark Gray
    oChart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Orange
End Sub
```

```javascript
// JavaScript Code to populate data, create a 3D bar chart, and customize its appearance

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Populate cells with years
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Populate labels
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Populate financial data
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Add a 3D bar chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Customize data labels
oChart.SetShowPointDataLabel(1, 0, false, false, true, false);

// Set series fill colors
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)); // Dark Gray
oChart.SetSeriesFill(oFill, 0, false);
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Orange
oChart.SetSeriesFill(oFill, 1, false);
```