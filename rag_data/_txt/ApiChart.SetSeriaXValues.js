## Description / Описание

**English:** This example sets the x-axis values from the specified range to the specified series. It is used with scatter charts only.

**Russian:** Этот пример устанавливает значения оси x из указанного диапазона для указанной серии. Используется только с диаграммами рассеяния.

### Excel VBA Code

```vba
' Excel VBA code to set up x-axis values and create a scatter chart

Sub CreateScatterChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chartRange As Range
    Dim ser As Series

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
    ws.Range("B4").Value = 2017
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("C4").Value = 2018
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    ws.Range("D4").Value = 2019

    ' Define the range for the chart
    Set chartRange = ws.Range("A1:D3")

    ' Add a scatter chart
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=375, Top:=50, Height:=225)
    With chartObj.Chart
        .ChartType = xlXYScatter
        .SetSourceData Source:=chartRange
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"

        ' Set X-axis values for the first series
        Set ser = .SeriesCollection(1)
        ser.XValues = ws.Range("B4:D4")

        ' Customize marker fill and outline for the first series
        ser.MarkerBackgroundColor = RGB(51, 51, 51)
        ser.MarkerForegroundColor = RGB(51, 51, 51)
        ser.MarkerSize = 7
        ser.MarkerStyle = xlMarkerStyleCircle

        ' Add a second series if needed and customize
        If .SeriesCollection.Count > 1 Then
            Set ser = .SeriesCollection(2)
            ser.XValues = ws.Range("B4:D4")
            ser.MarkerBackgroundColor = RGB(255, 111, 61)
            ser.MarkerForegroundColor = RGB(255, 111, 61)
            ser.MarkerSize = 7
            ser.MarkerStyle = xlMarkerStyleCircle
        End If
    End With
End Sub
```

### OnlyOffice JS Code

```javascript
// OnlyOffice JS code to set up x-axis values and create a scatter chart

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
oWorksheet.GetRange("B4").SetValue(2017);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("C4").SetValue(2018);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);
oWorksheet.GetRange("D4").SetValue(2019);

// Add a scatter chart to the worksheet
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "scatter", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);

// Set the chart title
oChart.SetTitle("Financial Overview", 13);

// Set the X-axis values for the first series
oChart.SetSeriaXValues("'Sheet1'!$B$4:$D$4", 0);

// Create and set marker fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetMarkerFill(oFill, 0, 0, true);

// Create and set marker outline for the first series
var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
oChart.SetMarkerOutLine(oStroke, 0, 0, true);

// Create and set marker fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetMarkerFill(oFill, 1, 0, true);

// Create and set marker outline for the second series
oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)));
oChart.SetMarkerOutLine(oStroke, 1, 0, true);
```