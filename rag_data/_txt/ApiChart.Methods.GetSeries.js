# Description
This code populates a worksheet with financial data, creates a combination chart, formats the chart series, and inserts the chart type into a specific cell.
Этот код заполняет рабочий лист финансовыми данными, создает комбинированную диаграмму, форматирует серии диаграммы и вставляет тип диаграммы в определенную ячейку.

```vba
' VBA code to populate data, create a combo chart, format it, and insert chart type into a cell

Sub CreateFinancialOverviewChart()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

    ' Populate headers
    oWorksheet.Range("B1").Value = 2014
    oWorksheet.Range("C1").Value = 2015
    oWorksheet.Range("D1").Value = 2016

    ' Populate row labels
    oWorksheet.Range("A2").Value = "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"

    ' Populate data
    oWorksheet.Range("B2").Value = 200
    oWorksheet.Range("B3").Value = 250
    oWorksheet.Range("C2").Value = 240
    oWorksheet.Range("C3").Value = 260
    oWorksheet.Range("D2").Value = 280
    oWorksheet.Range("D3").Value = 280

    ' Add chart
    Dim oChart As ChartObject
    Set oChart = oWorksheet.ChartObjects.Add(Left:=100, Top:=70, Width:=360, Height:=360)
    With oChart.Chart
        .SetSourceData Source:=oWorksheet.Range("A1:D3")
        .ChartType = xlColumnClustered ' Base chart type

        ' Add a second series as a line chart for combination
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Values = oWorksheet.Range("B3:D3")
        .SeriesCollection(2).ChartType = xlLine

        ' Set chart title
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"

        ' Format first series
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
        
        ' Format second series
        .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(255, 111, 61)
    End With

    ' Get series type
    Dim sSeriesType As String
    sSeriesType = oChart.Chart.SeriesCollection(1).ChartType

    ' Insert series type into cell F1
    oWorksheet.Range("F1").Value = "1 Series Type = " & sSeriesType
End Sub
```

```javascript
// JavaScript code to populate data, create a combo chart, format it, and insert chart type into a cell

function createFinancialOverviewChart() {
    var oWorksheet = Api.GetActiveSheet();

    // Populate headers
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

    // Add chart
    var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "comboBarLine", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);
    
    // Set chart title
    oChart.SetTitle("Financial Overview", 13);
    
    // Format first series
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
    oChart.SetSeriesFill(oFill, 0, false);
    
    // Format second series
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    oChart.SetSeriesFill(oFill, 1, false);

    // Get series type
    var oSeries = oChart.GetSeries(0);
    var sSeriesType = oSeries.GetChartType();

    // Insert series type into cell F1
    oWorksheet.GetRange("F1").SetValue("1 Series Type = " + sSeriesType);
}
```