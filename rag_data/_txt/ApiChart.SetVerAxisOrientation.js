# Description / Описание

**English:**  
This script populates an OnlyOffice spreadsheet with financial data for the years 2014 to 2016, creates a 3D bar chart titled "Financial Overview," and sets specific colors for each data series. It also configures the vertical axis orientation.

**Русский:**  
Этот скрипт заполняет таблицу OnlyOffice финансовыми данными за 2014-2016 годы, создает 3D столбчатую диаграмму с заголовком "Финансовый обзор" и устанавливает определенные цвета для каждого ряда данных. Также настраивается ориентация вертикальной оси.

```javascript
// This example specifies the direction of the data displayed on the vertical axis.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set year headers
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set row labels
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Populate data for Projected Revenue
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("D2").SetValue(280);

// Populate data for Estimated Costs
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D3").SetValue(280);

// Add a 3D bar chart to the worksheet
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set the chart title
oChart.SetTitle("Financial Overview", 13);

// Set vertical axis orientation
oChart.SetVerAxisOrientation(false);

// Create and set fill color for the first data series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set fill color for the second data series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```

```vba
' Description / Описание
' This VBA script populates an Excel worksheet with financial data for the years 2014 to 2016,
' creates a 3D bar chart titled "Financial Overview," and sets specific colors for each data series.
' It also configures the vertical axis orientation.

Sub CreateFinancialChart()
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set year headers
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    
    ' Set row labels
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    
    ' Populate data for Projected Revenue
    ws.Range("B2").Value = 200
    ws.Range("C2").Value = 240
    ws.Range("D2").Value = 280
    
    ' Populate data for Estimated Costs
    ws.Range("B3").Value = 250
    ws.Range("C3").Value = 260
    ws.Range("D3").Value = 280
    
    ' Add a 3D bar chart to the worksheet
    Dim cht As Chart
    Set cht = ws.Shapes.AddChart2(251, xlBarClustered, 200, 100, 360, 270).Chart ' Adjust position and size as needed
    cht.SetSourceData Source:=ws.Range("A1:D3")
    
    ' Change chart type to 3D Bar
    cht.ChartType = xl3DBarClustered
    
    ' Set the chart title
    cht.HasTitle = True
    cht.ChartTitle.Text = "Financial Overview"
    cht.ChartTitle.Font.Size = 13
    
    ' Set vertical axis orientation
    cht.Axes(xlValue).ReversePlotOrder = False
    
    ' Set fill color for the first data series
    cht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
    
    ' Set fill color for the second data series
    cht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
End Sub
```