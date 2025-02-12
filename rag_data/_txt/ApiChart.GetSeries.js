**English:** This script populates cells with financial data, creates a combination bar and line chart, customizes the chart's appearance, retrieves the chart type of the first series, and outputs it to a specific cell.

**Russian:** Этот скрипт заполняет ячейки финансовыми данными, создает комбинированный столбчатый и линейный график, настраивает внешний вид графика, получает тип графика для первой серии и выводит его в определенную ячейку.

```vba
' VBA Code

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chartSeries As Series
    Dim seriesType As String

    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Populate header years
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016

    ' Populate row labels
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"

    ' Populate financial data
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280

    ' Add a combination chart (Column and Line)
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=375, Top:=50, Height:=225)
    With chartObj.Chart
        .SetSourceData Source:=ws.Range("A1:D3")
        .ChartType = xlColumnClustered

        ' Add a secondary series as a line
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Formula = "=SERIES(""Projected Revenue"",Sheet1!$B$2:$D$2,Sheet1!$B$2:$D$2,2)"
        .SeriesCollection(2).ChartType = xlLine
        .SeriesCollection(2).AxisGroup = xlSecondary

        ' Set chart title
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"

        ' Customize series fill colors
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
        .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(255, 111, 61)

        ' Get the chart type of the first series
        seriesType = .SeriesCollection(1).ChartType

        ' Output the series type to cell F1
        ws.Range("F1").Value = "1 Series Type = " & seriesType
    End With
End Sub
```

```javascript
// OnlyOffice JS Code

// This script populates cells with financial data, creates a combination bar and line chart,
// customizes the chart's appearance, retrieves the chart type of the first series, and outputs it to a specific cell.

function createFinancialChart(Api) {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Populate header years
    oWorksheet.GetRange("B1").SetValue(2014);
    oWorksheet.GetRange("C1").SetValue(2015);
    oWorksheet.GetRange("D1").SetValue(2016);
    
    // Populate row labels
    oWorksheet.GetRange("A2").SetValue("Projected Revenue");
    oWorksheet.GetRange("A3").SetValue("Estimated Costs");
    
    // Populate financial data
    oWorksheet.GetRange("B2").SetValue(200);
    oWorksheet.GetRange("B3").SetValue(250);
    oWorksheet.GetRange("C2").SetValue(240);
    oWorksheet.GetRange("C3").SetValue(260);
    oWorksheet.GetRange("D2").SetValue(280);
    oWorksheet.GetRange("D3").SetValue(280);
    
    // Add a combination chart (Combo Bar and Line)
    var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "comboBarLine", 
        2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);
    
    // Set chart title
    oChart.SetTitle("Financial Overview", 13);
    
    // Customize series fill colors
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
    oChart.SetSeriesFill(oFill, 0, false);
    
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    oChart.SetSeriesFill(oFill, 1, false);
    
    // Get the chart type of the first series
    var oSeries = oChart.GetSeries(0);
    var sSeriesType = oSeries.GetChartType();
    
    // Output the series type to cell F1
    oWorksheet.GetRange("F1").SetValue("1 Series Type = " + sSeriesType);
}
```