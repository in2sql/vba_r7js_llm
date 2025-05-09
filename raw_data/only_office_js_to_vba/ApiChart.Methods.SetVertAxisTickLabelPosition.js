# Code Description / Описание кода

**English:**  
This script sets specific values in an Excel worksheet, creates a 3D bar chart titled "Financial Overview," positions the vertical axis tick labels, and applies custom fill colors to the chart series.

**Русский:**  
Этот скрипт устанавливает определенные значения в рабочем листе Excel, создает 3D столбчатую диаграмму с заголовком "Финансовый обзор", позиционирует метки делений вертикальной оси и применяет пользовательские цвета заливки к сериям диаграммы.

```vba
Sub CreateFinancialOverviewChart()
    ' Set values in the worksheet
    With ThisWorkbook.ActiveSheet
        .Range("B1").Value = 2014
        .Range("C1").Value = 2015
        .Range("D1").Value = 2016
        .Range("A2").Value = "Projected Revenue"
        .Range("A3").Value = "Estimated Costs"
        .Range("B2").Value = 200
        .Range("B3").Value = 250
        .Range("C2").Value = 240
        .Range("C3").Value = 260
        .Range("D2").Value = 280
        .Range("D3").Value = 280
        
        ' Add a 3D bar chart
        Dim oChart As ChartObject
        Set oChart = .ChartObjects.Add(Left:=100, Top:=70, Width:=360, Height:=360)
        With oChart.Chart
            .SetSourceData Source:=.Parent.Range("'Sheet1'!$A$1:$D$3")
            .ChartType = xlBar3DClustered
            .HasTitle = True
            .ChartTitle.Text = "Financial Overview"
            
            ' Set vertical axis tick label position to high
            .Axes(xlValue).TickLabels.Position = xlTickLabelPositionHigh
            
            ' Set fill color for series 1
            .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
            
            ' Set fill color for series 2
            .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
        End With
    End With
End Sub
```

```javascript
// Set values in the worksheet
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

// Add a 3D bar chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);

// Set the chart title
oChart.SetTitle("Financial Overview", 13);

// Set vertical axis tick label position to high
oChart.SetVertAxisTickLabelPosition("high");

// Create and set fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```