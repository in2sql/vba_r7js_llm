# Create Financial Overview Chart / Создать диаграмму Финансового обзора

```javascript
// English description:
// This code populates cells with financial data for 2014-2016,
// creates a 3D bar chart titled "Financial Overview",
// sets the vertical axis labels' font size,
// and applies custom fill colors to the series.

// Russian description:
// Этот код заполняет ячейки финансовыми данными за 2014-2016 годы,
// создает 3D-столбчатую диаграмму с заголовком "Финансовый обзор",
// устанавливает размер шрифта для вертикальных меток оси,
// и применяет пользовательские цвета заливки к сериям.

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
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);
oChart.SetTitle("Financial Overview", 13);
oChart.SetVertAxisLablesFontSize(10);
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```

```vba
' English description:
' This VBA code populates cells with financial data for 2014-2016,
' creates a 3D bar chart titled "Financial Overview",
' sets the vertical axis labels' font size,
' and applies custom fill colors to the series.

' Russian description:
' Этот код VBA заполняет ячейки финансовыми данными за 2014-2016 годы,
' создает 3D-столбчатую диаграмму с заголовком "Финансовый обзор",
' устанавливает размер шрифта для вертикальных меток оси,
' и применяет пользовательские цвета заливки к сериям.

Sub CreateFinancialOverviewChart()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Populate year headers
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    
    ' Populate categories
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    
    ' Populate data
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    
    ' Add 3D bar chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=200, Top:=100, Width:=360, Height:=270)
    With chartObj.Chart
        .SetSourceData Source:=ws.Range("A1:D3")
        .ChartType = xlBar3DClustered
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        
        ' Set vertical axis label font size
        .Axes(xlValue).TickLabels.Font.Size = 10
        
        ' Set series fill colors
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    End With
End Sub
```