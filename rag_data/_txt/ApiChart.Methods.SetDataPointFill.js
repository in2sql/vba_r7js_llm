**Description / Описание:**

English:  
This code sets up an Excel worksheet with projected revenue and estimated costs for the years 2014 to 2016. It adds a 3D bar chart titled "Financial Overview" and applies specific fill colors to the chart series and data points.

Russian:  
Этот код настраивает рабочий лист Excel с прогнозируемыми доходами и оценочными затратами за годы 2014–2016. Он добавляет 3D столбчатую диаграмму с заголовком "Финансовый обзор" и применяет определенные цвета заполнения к сериям и точкам данных диаграммы.

```vba
' VBA Code to replicate the OnlyOffice JS functionality

Sub CreateFinancialOverviewChart()
    Dim oWorksheet As Worksheet
    Dim oChart As Chart
    Dim oFill As Object

    ' Set the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

    ' Set values for years
    oWorksheet.Range("B1").Value = 2014
    oWorksheet.Range("C1").Value = 2015
    oWorksheet.Range("D1").Value = 2016

    ' Set labels
    oWorksheet.Range("A2").Value = "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"

    ' Set projected revenue values
    oWorksheet.Range("B2").Value = 200
    oWorksheet.Range("C2").Value = 240
    oWorksheet.Range("D2").Value = 280

    ' Set estimated costs values
    oWorksheet.Range("B3").Value = 250
    oWorksheet.Range("C3").Value = 260
    oWorksheet.Range("D3").Value = 280

    ' Add a 3D Bar Chart
    Set oChart = oWorksheet.Shapes.AddChart2(227, xlBarClustered, 200, 100, 360, 70).Chart
    oChart.SetSourceData Source:=oWorksheet.Range("'Sheet1'!$A$1:$D$3")
    oChart.ChartTitle.Text = "Financial Overview"

    ' Set fill colors for series
    ' Series 1
    oChart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
    
    ' Series 2
    oChart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Set fill color for specific data point (e.g., first data point of Series 1)
    oChart.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = RGB(128, 128, 128)
End Sub
```

```javascript
// This example shows how to set the fill to the data point.
var oWorksheet = Api.GetActiveSheet();
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
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);
oChart.SetTitle("Financial Overview", 13);
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
oFill = Api.CreateSolidFill(Api.CreateRGBColor(128, 128, 128));
oChart.SetDataPointFill(oFill, 0, 0, false); 
```