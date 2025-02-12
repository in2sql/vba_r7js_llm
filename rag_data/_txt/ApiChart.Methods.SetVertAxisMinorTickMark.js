**English:** This code populates the worksheet with financial data and creates a scatter chart titled "Financial Overview" with customized marker fills and outlines.

**Russian:** Этот код заполняет лист финансовыми данными и создает точечную диаграмму с заголовком "Обзор финансов", а также настраиваемыми заливками и контурами маркеров.

```javascript
// This code populates the worksheet with financial data and creates a customized scatter chart
var oWorksheet = Api.GetActiveSheet();

// Populate headers with years
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Populate row labels
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Populate financial data
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);

// Add a scatter chart with specified range and positioning
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "scatter", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set the chart title
oChart.SetTitle("Financial Overview", 13);

// Set minor tick marks for the vertical axis
oChart.SetVertAxisMinorTickMark("out");

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

```vba
' This VBA code populates the worksheet with financial data and creates a customized scatter chart
Sub CreateFinancialChart()
    Dim oWorksheet As Worksheet
    Dim oChart As Chart
    Dim oSeries As Series
    
    ' Set the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Populate headers with years
    oWorksheet.Range("B1").Value = 2014
    oWorksheet.Range("C1").Value = 2015
    oWorksheet.Range("D1").Value = 2016
    
    ' Populate row labels
    oWorksheet.Range("A2").Value = "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"
    
    ' Populate financial data
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
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        
        ' Set minor tick marks for the vertical axis
        .Axes(xlValue).MinorTickMark = xlTickMarkOutside
        
        ' Customize marker for first series
        Set oSeries = .SeriesCollection(1)
        With oSeries.MarkerBackgroundColor = RGB(51, 51, 51)
            .MarkerForegroundColor = RGB(51, 51, 51)
            .MarkerSize = 8
            .MarkerStyle = xlMarkerStyleCircle
        End With
        
        ' Customize marker for second series
        Set oSeries = .SeriesCollection(2)
        With oSeries.MarkerBackgroundColor = RGB(255, 111, 61)
            .MarkerForegroundColor = RGB(255, 111, 61)
            .MarkerSize = 8
            .MarkerStyle = xlMarkerStyleCircle
        End With
    End With
End Sub
```