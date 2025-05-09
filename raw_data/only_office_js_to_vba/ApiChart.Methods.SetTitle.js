# Create a Financial Overview Chart in Excel | Создание диаграммы "Финансовый обзор" в Excel

```vba
' VBA code to create a Financial Overview chart
Sub CreateFinancialOverviewChart()
    Dim oWorksheet As Worksheet
    Dim oChart As Chart
    Dim oFill As Object

    ' Set the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

    ' Populate the header values
    oWorksheet.Range("B1").Value = 2014
    oWorksheet.Range("C1").Value = 2015
    oWorksheet.Range("D1").Value = 2016

    ' Populate the row labels
    oWorksheet.Range("A2").Value = "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"

    ' Populate the data values
    oWorksheet.Range("B2").Value = 200
    oWorksheet.Range("B3").Value = 250
    oWorksheet.Range("C2").Value = 240
    oWorksheet.Range("C3").Value = 260
    oWorksheet.Range("D2").Value = 280
    oWorksheet.Range("D3").Value = 280

    ' Add a 3D Bar Chart
    Set oChart = oWorksheet.Shapes.AddChart2(251, xlBar3DClustered, 200, 100, 360, 270).Chart
    oChart.SetSourceData Source:=oWorksheet.Range("A1:D3")

    ' Set the chart title
    With oChart
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
    End With

    ' Set the fill color for the first series
    Set oFill = oChart.SeriesCollection(1).Format.Fill
    oFill.Solid
    oFill.ForeColor.RGB = RGB(51, 51, 51)

    ' Set the fill color for the second series
    Set oFill = oChart.SeriesCollection(2).Format.Fill
    oFill.Solid
    oFill.ForeColor.RGB = RGB(255, 111, 61)
End Sub
```

```javascript
// JavaScript code to create a Financial Overview chart using OnlyOffice API
function createFinancialOverviewChart() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();

    // Populate the header values
    oWorksheet.GetRange("B1").SetValue(2014);
    oWorksheet.GetRange("C1").SetValue(2015);
    oWorksheet.GetRange("D1").SetValue(2016);

    // Populate the row labels
    oWorksheet.GetRange("A2").SetValue("Projected Revenue");
    oWorksheet.GetRange("A3").SetValue("Estimated Costs");

    // Populate the data values
    oWorksheet.GetRange("B2").SetValue(200);
    oWorksheet.GetRange("B3").SetValue(250);
    oWorksheet.GetRange("C2").SetValue(240);
    oWorksheet.GetRange("C3").SetValue(260);
    oWorksheet.GetRange("D2").SetValue(280);
    oWorksheet.GetRange("D3").SetValue(280);

    // Add a 3D Bar Chart
    var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

    // Set the chart title
    oChart.SetTitle("Financial Overview", 13);

    // Set the fill color for the first series
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
    oChart.SetSeriesFill(oFill, 0, false);

    // Set the fill color for the second series
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    oChart.SetSeriesFill(oFill, 1, false);
}
```