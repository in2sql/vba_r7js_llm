# Financial Overview Chart Creation / Создание диаграммы "Обзор Финансов"

This code sets up data in a worksheet and creates a 3D bar chart with customized titles, labels, and series fills.
Этот код заполняет данные на листе и создает 3D-столбчатую диаграмму с настраиваемыми заголовками, подписями и заливкой серий.

```vba
' VBA Code

Sub CreateFinancialOverviewChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chartRange As Range
    Dim fillColor1 As Long
    Dim fillColor2 As Long
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Populate data
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
    
    ' Define data range for chart
    Set chartRange = ws.Range("A1:D3")
    
    ' Add a 3D bar chart
    Set chartObj = ws.ChartObjects.Add(Left:=200, Top:=100, Width:=360, Height:=70)
    With chartObj.Chart
        .ChartType = xlBar3DClustered
        .SetSourceData Source:=chartRange
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
        .Axes(xlValue).TickLabels.Font.Size = 10
    End With
    
    ' Set series fill colors
    fillColor1 = RGB(51, 51, 51) ' Dark grey
    fillColor2 = RGB(255, 111, 61) ' Orange
    
    With chartObj.Chart.SeriesCollection(1)
        .Format.Fill.ForeColor.RGB = fillColor1
    End With
    
    With chartObj.Chart.SeriesCollection(2)
        .Format.Fill.ForeColor.RGB = fillColor2
    End With
End Sub
```

```javascript
// OnlyOffice JS Code

// This code sets up data and creates a 3D bar chart with custom titles, labels, and series fills
// Этот код заполняет данные и создает 3D-столбчатую диаграмму с пользовательскими заголовками, подписями и заливкой серий

function createFinancialOverviewChart() {
    var oWorksheet = Api.GetActiveSheet();
    
    // Populate data
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
    
    // Set chart title and font size
    oChart.SetTitle("Financial Overview", 13);
    
    // Set vertical axis labels font size
    oChart.SetVertAxisLablesFontSize(10);
    
    // Set series fill colors
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
    oChart.SetSeriesFill(oFill, 0, false);
    oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    oChart.SetSeriesFill(oFill, 1, false);
}
```