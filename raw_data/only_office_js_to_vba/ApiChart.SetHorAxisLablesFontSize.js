**Description | Описание**

This code populates an Excel sheet with financial data for the years 2014-2016, creates a 3D bar chart titled "Financial Overview", adjusts the font sizes of the chart's title and horizontal axis labels, and sets specific colors for the chart series.

Этот код заполняет лист Excel финансовыми данными за 2014-2016 годы, создает 3D столбчатую диаграмму с заголовком "Финансовый обзор", настраивает размеры шрифтов заголовка диаграммы и меток горизонтальной оси, а также устанавливает определенные цвета для серий диаграммы.

```vba
' Excel VBA Code

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chartRange As Range
    Dim series1 As Series
    Dim series2 As Series
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Populate header years
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    
    ' Populate category labels
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    
    ' Populate financial data
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    
    ' Define the range for the chart
    Set chartRange = ws.Range("'A1':'D3'")
    
    ' Add a 3D bar chart
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=360, Top:=100, Height:=270)
    With chartObj.Chart
        .SetSourceData Source:=chartRange
        .ChartType = xlBar3DClustered
        
        ' Set chart title
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
        
        ' Set horizontal axis label font size
        .Axes(xlCategory).TickLabels.Font.Size = 10
        
        ' Set series fill colors
        Set series1 = .SeriesCollection(1)
        series1.Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
        
        Set series2 = .SeriesCollection(2)
        series2.Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    End With
End Sub
```

```javascript
// OnlyOffice JS Code

// This script populates the sheet with financial data and creates a 3D bar chart with customized formatting.
function createFinancialChart(Api) {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Populate header years
    oWorksheet.GetRange("B1").SetValue(2014);
    oWorksheet.GetRange("C1").SetValue(2015);
    oWorksheet.GetRange("D1").SetValue(2016);
    
    // Populate category labels
    oWorksheet.GetRange("A2").SetValue("Projected Revenue");
    oWorksheet.GetRange("A3").SetValue("Estimated Costs");
    
    // Populate financial data
    oWorksheet.GetRange("B2").SetValue(200);
    oWorksheet.GetRange("B3").SetValue(250);
    oWorksheet.GetRange("C2").SetValue(240);
    oWorksheet.GetRange("C3").SetValue(260);
    oWorksheet.GetRange("D2").SetValue(280);
    oWorksheet.GetRange("D3").SetValue(280);
    
    // Add a 3D bar chart to the worksheet
    var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);
    
    // Set chart title and font size
    oChart.SetTitle("Financial Overview", 13);
    
    // Set horizontal axis labels font size
    oChart.SetHorAxisLablesFontSize(10);
    
    // Set fill color for the first series
    var oFill1 = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
    oChart.SetSeriesFill(oFill1, 0, false);
    
    // Set fill color for the second series
    var oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    oChart.SetSeriesFill(oFill2, 1, false);
}
```