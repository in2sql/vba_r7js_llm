**Description / Описание**

This code demonstrates how to remove a specified series from the current chart.
Этот код демонстрирует, как удалить определенный ряд из текущей диаграммы.

```vba
' VBA Code to remove a specified series from the current chart

Sub ModifyChart()
    Dim ws As Worksheet
    Dim cht As Chart
    Dim fillColor As Long

    ' Get the active worksheet
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

    ' Add a 3D bar chart
    Set cht = ws.Shapes.AddChart2(227, xlBarClustered, 100, 70, 300, 200).Chart
    cht.SetSourceData Source:=ws.Range("A1:D3")
    cht.ChartType = xl3DBarClustered

    ' Set chart title
    cht.HasTitle = True
    cht.ChartTitle.Text = "Financial Overview"
    cht.ChartTitle.Font.Size = 13

    ' Remove the second series (Estimated Costs)
    If cht.SeriesCollection.Count > 1 Then
        cht.SeriesCollection(2).Delete
    End If

    ' Set series fill color
    fillColor = RGB(255, 111, 61)
    cht.SeriesCollection(1).Format.Fill.ForeColor.RGB = fillColor

    ' Add a note in cell A5
    ws.Range("A5").Value = "The Estimated Costs series was removed from the current chart."
End Sub
```

```javascript
// JavaScript Code to remove a specified series from the current chart

function modifyChart() {
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

    // Remove the second series (Estimated Costs)
    oChart.RemoveSeria(1);

    // Create and set series fill color
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    oChart.SetSeriesFill(oFill, 0, false);

    // Add a note in cell A5
    oWorksheet.GetRange("A5").SetValue("The Estimated Costs series was removed from the current chart.");
}
```