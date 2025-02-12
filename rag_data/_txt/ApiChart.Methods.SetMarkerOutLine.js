**Description / Описание**

This code sets up a worksheet with data for a financial overview and creates a customized scatter chart.

Этот код настраивает рабочий лист с данными для финансового обзора и создает точечную диаграмму с настраиваемыми маркерами.

```vba
' VBA Code to set up worksheet data and create a customized scatter chart

Sub CreateFinancialOverviewChart()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

    ' Set year headers
    oWorksheet.Range("B1").Value = 2014
    oWorksheet.Range("C1").Value = 2015
    oWorksheet.Range("D1").Value = 2016

    ' Set labels
    oWorksheet.Range("A2").Value = "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"

    ' Set data values
    oWorksheet.Range("B2").Value = 200
    oWorksheet.Range("B3").Value = 250
    oWorksheet.Range("C2").Value = 240
    oWorksheet.Range("C3").Value = 260
    oWorksheet.Range("D2").Value = 280
    oWorksheet.Range("D3").Value = 280

    ' Add scatter chart
    Dim oChart As ChartObject
    Set oChart = oWorksheet.ChartObjects.Add(Left:=100, Width:=350, Top:=70, Height:=250)
    With oChart.Chart
        .SetSourceData Source:=oWorksheet.Range("A1:D3")
        .ChartType = xlXYScatter
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"

        ' Customize marker for first series
        With .SeriesCollection(1)
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerBackgroundColor = RGB(51, 51, 51)
            .MarkerForegroundColor = RGB(51, 51, 51)
            .MarkerSize = 7
            .Format.Line.Weight = 0.5
            .Format.Line.ForeColor.RGB = RGB(51, 51, 51)
        End With

        ' Customize marker for second series
        With .SeriesCollection(2)
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerBackgroundColor = RGB(255, 111, 61)
            .MarkerForegroundColor = RGB(255, 111, 61)
            .MarkerSize = 7
            .Format.Line.Weight = 0.5
            .Format.Line.ForeColor.RGB = RGB(51, 51, 51)
        End With
    End With
End Sub
```

```javascript
// This code sets up a worksheet with data for a financial overview and creates a customized scatter chart.

var oWorksheet = Api.GetActiveSheet();

// Set year headers
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set labels
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set data values
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Add scatter chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "scatter", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Create and set marker fill for first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetMarkerFill(oFill, 0, 0, true);

// Create and set marker outline for first series
var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
oChart.SetMarkerOutLine(oStroke, 0, 0, true);

// Create and set marker fill for second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetMarkerFill(oFill, 1, 0, true);

// Set marker outline for second series
oChart.SetMarkerOutLine(oStroke, 1, 0, true);
```