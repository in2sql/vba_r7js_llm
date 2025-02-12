### Description / Описание
This code populates an Excel sheet with financial data and creates a 3D bar chart with customized colors and title.
Этот код заполняет лист Excel финансовыми данными и создает 3D-гистограмму с настраиваемыми цветами и заголовком.

```vba
' Excel VBA code to populate data and create a 3D bar chart

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chartRange As Range
    Dim series1 As Series
    Dim series2 As Series

    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Populate the data
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
    Set chartRange = ws.Range("A1:D3")

    ' Add a 3D bar chart
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=360, Top:=100, Height:=270)
    With chartObj.Chart
        .SetSourceData Source:=chartRange
        .ChartType = xl3DBarClustered

        ' Set the chart title
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"

        ' Set the vertical axis orientation
        .Axes(xlValue).ReversePlotOrder = True

        ' Set series fill colors
        Set series1 = .SeriesCollection(1)
        series1.Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' Dark Gray

        Set series2 = .SeriesCollection(2)
        series2.Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Orange
    End With
End Sub
```

```javascript
// OnlyOffice JavaScript code to populate data and create a 3D bar chart

// This example specifies the direction of the data displayed on the vertical axis.
var oWorksheet = Api.GetActiveSheet();

// Populate the data
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

// Set the vertical axis orientation
oChart.SetVerAxisOrientation(false);

// Set series fill colors
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)); // Dark Gray
oChart.SetSeriesFill(oFill, 0, false);

oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Orange
oChart.SetSeriesFill(oFill, 1, false);
```