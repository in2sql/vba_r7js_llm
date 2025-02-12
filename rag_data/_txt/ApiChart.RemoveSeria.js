```javascript
// This code removes the specified series from the current chart.
// Этот код удаляет указанную серию из текущего графика.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set header values
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
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);

// Set the chart title
oChart.SetTitle("Financial Overview", 13);

// Remove the second series from the chart
oChart.RemoveSeria(1);

// Create a solid fill color
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Set the fill for the first series
oChart.SetSeriesFill(oFill, 0, false);

// Add a note to the worksheet
oWorksheet.GetRange("A5").SetValue("The Estimated Costs series was removed from the current chart.");
```

```vba
' This code removes the specified series from the current chart.
' Этот код удаляет указанную серию из текущего графика.

Sub RemoveSeriesFromChart()
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
    
    ' Set data values
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    
    ' Add a 3D bar chart
    Dim cht As Chart
    Set cht = ws.Shapes.AddChart3(240, xl3DBarClustered, 100, 70, 350, 225).Chart
    
    ' Set the chart title
    cht.HasTitle = True
    cht.ChartTitle.Text = "Financial Overview"
    
    ' Remove the second series from the chart
    cht.SeriesCollection(2).Delete
    
    ' Set the fill color for the first series
    cht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    
    ' Add a note to the worksheet
    ws.Range("A5").Value = "The Estimated Costs series was removed from the current chart."
End Sub
```