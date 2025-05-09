# Description / Описание

This code sets values in specific cells, creates a 3D bar chart, sets the chart title, and applies specific fill colors to chart series and data points.

Этот код устанавливает значения в определенные ячейки, создает 3D столбчатую диаграмму, задает заголовок диаграммы и применяет определенные цвета заполнения к сериям и данным диаграммы.

```vba
' VBA code to set cell values, create a 3D bar chart, set title, and apply fill colors
Sub CreateFinancialOverviewChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As Chart
    Dim fillColor As Long
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set values in cells
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    
    ' Add a 3D bar chart
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=3600, Top:=100, Height:=2700)
    Set chart = chartObj.Chart
    chart.SetSourceData Source:=ws.Range("A1:D3")
    chart.ChartType = xlBarClustered ' Excel does not have a built-in 3D bar type, xlBarClustered is closest
    
    ' Set chart title
    With chart
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
    End With
    
    ' Set fill colors for series
    ' Assuming two series: Projected Revenue and Estimated Costs
    If chart.SeriesCollection.Count >= 2 Then
        ' Set first series fill color to RGB(51, 51, 51)
        chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
        
        ' Set second series fill color to RGB(255, 111, 61)
        chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    End If
    
    ' Set fill color for the first data point of the first series to RGB(128, 128, 128)
    If chart.SeriesCollection.Count >= 1 Then
        If chart.SeriesCollection(1).Points.Count >= 1 Then
            chart.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = RGB(128, 128, 128)
        End If
    End If
End Sub
```

```javascript
// This example shows how to set the fill to the data point.
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

// Add a 3D bar chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Create and set fill colors for series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Create and set fill color for a specific data point
oFill = Api.CreateSolidFill(Api.CreateRGBColor(128, 128, 128));
oChart.SetDataPointFill(oFill, 0, 0, false);
```