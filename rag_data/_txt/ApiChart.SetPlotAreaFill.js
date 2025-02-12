### Description / Описание
This code sets up data in an Excel worksheet, creates a 3D bar chart titled "Financial Overview," and applies specific fill colors to the chart series and plot area.
Этот код заполняет данные в листе Excel, создает 3D столбчатую диаграмму с заголовком "Financial Overview" и применяет определенные цвета заливки к сериям диаграммы и области построения.

```vba
' Excel VBA Equivalent Code
Sub CreateFinancialOverviewChart()
    Dim ws As Worksheet
    Dim rng As Range
    Dim chartObj As ChartObject
    Dim chartRange As Range
    Dim plotAreaFill As Shape
    Dim seriesFill1 As Shape
    Dim seriesFill2 As Shape

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

    ' Define the range for the chart
    Set chartRange = ws.Range("'Sheet1'!$A$1:$D$3")

    ' Add a 3D bar chart
    Set chartObj = ws.ChartObjects.Add(Left:=200, Top:=100, Width:=360, Height:=70)
    With chartObj.Chart
        .ChartType = xlBarClustered ' xlBarClustered is the closest equivalent
        .SetSourceData Source:=chartRange
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
    End With

    ' Set fill for series 1
    With chartObj.Chart.SeriesCollection(1).Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(51, 51, 51) ' Dark Gray
        .Solid
    End With

    ' Set fill for series 2
    With chartObj.Chart.SeriesCollection(2).Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 111, 61) ' Orange
        .Solid
    End With

    ' Set fill for plot area
    With chartObj.Chart.PlotArea.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(128, 128, 128) ' Gray
        .Solid
    End With
End Sub
```

```javascript
// OnlyOffice JS Equivalent Code
// This example sets the fill to the chart plot area.
// Этот пример устанавливает заливку для области построения диаграммы.

function createFinancialOverviewChart(Api) {
    // Get the active worksheet
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

    // Create and set fill for series 1
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)); // Dark Gray
    oChart.SetSeriesFill(oFill, 0, false);

    // Create and set fill for series 2
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Orange
    oChart.SetSeriesFill(oFill, 1, false);

    // Create and set fill for plot area
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(128, 128, 128)); // Gray
    oChart.SetPlotAreaFill(oFill);
}
```