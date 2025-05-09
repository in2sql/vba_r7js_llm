**Description / Описание:**
This code sets up an Excel worksheet with projected revenue and estimated costs for the years 2014-2016. It then creates a 3D bar chart titled "Financial Overview," applies specific fill colors to the chart series, and configures the minor horizontal gridlines.

Этот код настраивает рабочий лист Excel с прогнозируемыми доходами и предполагаемыми затратами на годы 2014-2016. Затем он создает 3D столбчатую диаграмму с названием "Финансовый обзор", применяет определенные цвета заливки к сериям диаграммы и настраивает мелкие горизонтальные сетки.

```vba
' VBA Code
Sub CreateFinancialOverviewChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As Chart
    Dim rng As Range
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set cell values
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
    
    ' Define the range for the chart
    Set rng = ws.Range("A1:D3")
    
    ' Add a 3D Bar Chart
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=360, Top:=100, Height:=270)
    Set chart = chartObj.Chart
    chart.ChartType = xl3DColumn
    
    ' Set chart title
    chart.HasTitle = True
    chart.ChartTitle.Text = "Financial Overview"
    chart.ChartTitle.Font.Size = 13
    
    ' Set series fill colors
    chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' Dark Gray
    chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Orange
    
    ' Configure minor horizontal gridlines
    With chart.Axes(xlCategory).MinorGridlines
        .Format.Line.Weight = 1
        .Format.Line.ForeColor.RGB = RGB(255, 111, 61) ' Orange
    End With
End Sub
```

```javascript
// OnlyOffice JS Code
// This script sets up a worksheet with financial data and creates a customized 3D bar chart.

function createFinancialOverviewChart() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set cell values
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
    
    // Add a 3D Bar Chart
    var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);
    
    // Set chart title
    oChart.SetTitle("Financial Overview", 13);
    
    // Set series fill colors
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)); // Dark Gray
    oChart.SetSeriesFill(oFill, 0, false);
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Orange
    oChart.SetSeriesFill(oFill, 1, false);
    
    // Configure minor horizontal gridlines
    var oStroke = Api.CreateStroke(1 * 5000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))); // Orange
    oChart.SetMinorHorizontalGridlines(oStroke);
}

// Execute the function
createFinancialOverviewChart();
```