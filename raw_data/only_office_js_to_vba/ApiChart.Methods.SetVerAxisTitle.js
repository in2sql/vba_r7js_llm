# Description / Описание

**English:** This example creates a 3D bar chart with financial data, sets the chart title and vertical axis title, and applies specific colors to the chart series.

**Russian:** Этот пример создает 3D столбчатую диаграмму с финансовыми данными, устанавливает заголовок диаграммы и заголовок вертикальной оси, а также применяет определенные цвета к сериям диаграммы.

## OnlyOffice JS Code

```javascript
// This example creates a 3D bar chart with financial data, sets the chart title and vertical axis title, and applies specific colors to the chart series.
var oWorksheet = Api.GetActiveSheet();

// Set years in the header
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set row labels
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set data values
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

// Set vertical axis title
oChart.SetVerAxisTitle("USD In Hundred Thousands", 10);

// Create and set fill color for first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set fill color for second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```

## Excel VBA Code

```vba
' This example creates a 3D bar chart with financial data, sets the chart title and vertical axis title, and applies specific colors to the chart series.

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chartRange As Range
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set years in the header
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    
    ' Set row labels
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    
    ' Set data values
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    
    ' Define the range for the chart
    Set chartRange = ws.Range("A1:D3")
    
    ' Add a 3D bar chart
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=375, Top:=50, Height:=225)
    With chartObj.Chart
        .SetSourceData Source:=chartRange
        .ChartType = xlBarClustered ' VBA does not have a direct 3D bar type, use xlColumnClustered with 3D formatting
        .ApplyChartTemplate Template:=xl3D
        
        ' Set chart title
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
        
        ' Set vertical axis title
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "USD In Hundred Thousands"
        .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 10
        
        ' Set fill color for first series
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
        
        ' Set fill color for second series
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    End With
End Sub
```