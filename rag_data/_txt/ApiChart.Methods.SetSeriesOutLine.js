### Description / Описание
This code sets up a financial overview chart by populating specific cells with data and configuring a 3D bar chart with customized colors and outlines.

Этот код настраивает диаграмму финансового обзора, заполняя определенные ячейки данными и настраивая 3D-столбчатую диаграмму с пользовательскими цветами и контурами.

```javascript
// JavaScript OnlyOffice API code
// Sets up data and creates a 3D bar chart with customized colors and outlines

var oWorksheet = Api.GetActiveSheet();
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

var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);
oChart.SetTitle("Financial Overview", 13);

var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
oChart.SetSeriesOutLine(oStroke, 1, false);
```

```vba
' Excel VBA code
' Sets up data and creates a 3D bar chart with customized colors and outlines

Sub CreateFinancialOverviewChart()
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
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=200, Top:=100, Width:=360, Height:=270)
    With chartObj.Chart
        .SetSourceData Source:=ws.Range("A1:D3")
        .ChartType = xl3DBarClustered
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        
        ' Set series fill colors
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
        
        ' Set series outline
        With .SeriesCollection(2).Format.Line
            .Visible = msoTrue
            .Weight = 0.5
            .ForeColor.RGB = RGB(51, 51, 51)
        End With
    End With
End Sub
```