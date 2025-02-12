# Script to Populate Worksheet Data and Create a 3D Bar Chart  
# Скрипт для заполнения данных листа и создания 3D столбчатой диаграммы

```vba
' VBA code to populate worksheet data and create a 3D bar chart

Sub CreateFinancialChart()
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Set header values
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016

    ' Set row labels
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    ws.Range("A4").Value = "Cost price"

    ' Set data values for 2014
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("B4").Value = 50

    ' Set data values for 2015
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("C4").Value = 120

    ' Set data values for 2016
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    ws.Range("D4").Value = 160

    ' Add a 3D Bar Chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=100, Width:=360, Top:=70, Height:=360)
    With chartObj.Chart
        .SetSourceData Source:=ws.Range("A1:D3")
        .ChartType = xlBarClustered ' 3D bar chart type
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        
        ' Add 'Cost price' series
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "Cost price"
        .SeriesCollection(2).Values = ws.Range("B4:D4")
        .SeriesCollection(2).ChartType = xlBarClustered

        ' Set fill color for first series
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
        
        ' Set fill color for second series
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    End With
End Sub
```

```javascript
// JavaScript code to populate worksheet data and create a 3D bar chart

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set header values
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set row labels
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");
oWorksheet.GetRange("A4").SetValue("Cost price");

// Set data values for 2014
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("B4").SetValue(50);

// Set data values for 2015
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("C4").SetValue(120);

// Set data values for 2016
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);
oWorksheet.GetRange("D4").SetValue(160);

// Add a 3D Bar Chart with specified dimensions and position
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);

// Set the chart title
oChart.SetTitle("Financial Overview", 13);

// Add 'Cost price' series to the chart
oChart.AddSeria("Cost price", "'Sheet1'!$B$4:$D$4");

// Create and set fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```