# Description / Описание

**English:**  
This code populates an Excel worksheet with financial data for the years 2014-2016, creates a 3D bar chart titled "Financial Overview," sets the horizontal axis orientation, and applies specific fill colors to the chart series.

**Russian:**  
Этот код заполняет рабочий лист Excel финансовыми данными за 2014-2016 годы, создает 3D-гистограмму с названием "Финансовый обзор", устанавливает ориентацию горизонтальной оси и применяет определенные цвета заливки к сериям диаграммы.

```vba
' VBA Code equivalent to the OnlyOffice JS example

Sub CreateFinancialChart()
    Dim oWorksheet As Worksheet
    Dim oChart As Chart
    Dim oSeries As Series
    
    ' Set the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Populate the cells with data
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
    
    ' Add a 3D bar chart
    Set oChart = oWorksheet.Shapes.AddChart3(201, xlBarClustered, 100, 70, 360, 360).Chart
    oChart.SetSourceData Source:=oWorksheet.Range("'Sheet1'!$A$1:$D$3")
    
    ' Set the chart title
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "Financial Overview"
    oChart.ChartTitle.Font.Size = 13
    
    ' Set horizontal axis orientation
    oChart.Axes(xlCategory).ReversePlotOrder = False
    
    ' Set series fill colors
    Set oSeries = oChart.SeriesCollection(1)
    oSeries.Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' Dark Gray
    
    Set oSeries = oChart.SeriesCollection(2)
    oSeries.Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Orange

End Sub
```

```javascript
// OnlyOffice JS Code equivalent to the VBA example

// This example populates the worksheet with financial data, creates a 3D bar chart,
// sets the horizontal axis orientation, and applies specific fill colors to the chart series.

var oWorksheet = Api.GetActiveSheet();

// Populate the cells with data
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

// Add a 3D bar chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set the chart title
oChart.SetTitle("Financial Overview", 13);

// Set horizontal axis orientation
oChart.SetHorAxisOrientation(false);

// Set series fill colors
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)); // Dark Gray
oChart.SetSeriesFill(oFill, 0, false);

oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Orange
oChart.SetSeriesFill(oFill, 1, false);
```