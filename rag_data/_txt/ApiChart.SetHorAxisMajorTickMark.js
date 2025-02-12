```plaintext
// English Description:
// This code sets up data in specified cells and creates a scatter chart with various formatting options, including titles, axis tick marks, marker fills, and outlines.

// Russian Description:
// Этот код заполняет данные в указанных ячейках и создает диаграмму разброса с различными параметрами форматирования, включая заголовки, метки основных делений осей, заливку маркеров и контуры.
```

```vba
' VBA Code
Sub CreateFinancialChart()
    Dim oWorksheet As Worksheet
    Dim oChart As Chart
    Dim oSeries As Series
    
    ' Set the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set values in cells
    oWorksheet.Range("B1").Value = 2014
    oWorksheet.Range("C1").Value = 2015
    oWorksheet.Range("D1").Value = 2016
    
    oWorksheet.Range("A2").Value = "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"
    
    oWorksheet.Range("B2").Value = 200
    oWorksheet.Range("B3").Value = 250
    oWorksheet.Range("C2").Value = 240
    oWorksheet.Range("C3").Value = 260
    oWorksheet.Range("D2").Value = 280
    oWorksheet.Range("D3").Value = 280
    
    ' Add a scatter chart
    Set oChart = Charts.Add
    With oChart
        .ChartType = xlXYScatter
        .SetSourceData Source:=oWorksheet.Range("A1:D3")
        .Location Where:=xlLocationAsObject, Name:=oWorksheet.Name
        ' Set chart title
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
        ' Set horizontal axis major tick mark to cross
        With .Axes(xlCategory)
            .MajorTickMark = xlTickMarkCross
        End With
        ' Format first series markers
        Set oSeries = .SeriesCollection(1)
        With oSeries.MarkerBackgroundColor = RGB(51, 51, 51) ' Set marker fill color
            oSeries.MarkerForegroundColor = RGB(51, 51, 51) ' Set marker outline color
            oSeries.MarkerSize = 7
        End With
        ' Format second series markers
        Set oSeries = .SeriesCollection(2)
        With oSeries.MarkerBackgroundColor = RGB(255, 111, 61) ' Set marker fill color
            oSeries.MarkerForegroundColor = RGB(255, 111, 61) ' Set marker outline color
            oSeries.MarkerSize = 7
        End With
    End With
End Sub
```

```javascript
// OnlyOffice JS Code
// This example specifies the major tick mark "cross" for the horizontal axis.
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

// Add a scatter chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "scatter", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set chart title with font size 13
oChart.SetTitle("Financial Overview", 13);

// Set horizontal axis major tick mark to cross
oChart.SetHorAxisMajorTickMark("cross");

// Create and set marker fill for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetMarkerFill(oFill, 0, 0, true);

// Create and set marker outline for the first series
var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
oChart.SetMarkerOutLine(oStroke, 0, 0, true);

// Create and set marker fill for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetMarkerFill(oFill, 1, 0, true);

// Create and set marker outline for the second series
oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)));
oChart.SetMarkerOutLine(oStroke, 1, 0, true);
```