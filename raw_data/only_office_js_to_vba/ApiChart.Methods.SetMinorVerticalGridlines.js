### Description / Описание

**English:** This code sets values in specified cells, adds a 3D bar chart, sets its title, applies specific fill colors to the chart series, and configures the minor vertical gridlines.

**Russian:** Этот код устанавливает значения в указанных ячейках, добавляет 3D столбчатую диаграмму, устанавливает ее заголовок, применяет определенные цвета заливки к сериям диаграммы и настраивает второстепенные вертикальные сетки.

```vba
' VBA Code

Sub CreateFinancialOverviewChart()
    ' Set values in headers
    With ActiveSheet
        .Range("B1").Value = 2014
        .Range("C1").Value = 2015
        .Range("D1").Value = 2016
        
        ' Set labels
        .Range("A2").Value = "Projected Revenue"
        .Range("A3").Value = "Estimated Costs"
        
        ' Set data
        .Range("B2").Value = 200
        .Range("B3").Value = 250
        .Range("C2").Value = 240
        .Range("C3").Value = 260
        .Range("D2").Value = 280
        .Range("D3").Value = 280
        
        ' Add a 3D bar chart
        Dim chartObj As ChartObject
        Set chartObj = .ChartObjects.Add(Left:=100, Top:=70, Width:=360, Height:=360)
        With chartObj.Chart
            .SetSourceData Source:=.Parent.Range("A1:D3")
            .ChartType = xl3DBarClustered ' Set chart type to 3D Bar
            .HasTitle = True
            .ChartTitle.Text = "Financial Overview"
            
            ' Set series fill colors
            .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' Series 1 color
            .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Series 2 color
            
            ' Configure minor vertical gridlines
            With .Axes(xlValue)
                .HasMinorGridlines = True
                .MinorGridline.Border.Color = RGB(255, 111, 61) ' Set gridline color
                .MinorGridline.Border.LineStyle = xlContinuous ' Set gridline style
                .MinorGridline.Border.Weight = xlThin ' Set gridline weight
            End With
        End With
    End With
End Sub
```

```javascript
// JavaScript Code

// This script sets cell values, adds a 3D bar chart, sets its title,
// applies fill colors to chart series, and configures the minor vertical gridlines.

function createFinancialOverviewChart() {
    var oWorksheet = Api.GetActiveSheet();
    
    // Set header values
    oWorksheet.GetRange("B1").SetValue(2014);
    oWorksheet.GetRange("C1").SetValue(2015);
    oWorksheet.GetRange("D1").SetValue(2016);
    
    // Set labels
    oWorksheet.GetRange("A2").SetValue("Projected Revenue");
    oWorksheet.GetRange("A3").SetValue("Estimated Costs");
    
    // Set data
    oWorksheet.GetRange("B2").SetValue(200);
    oWorksheet.GetRange("B3").SetValue(250);
    oWorksheet.GetRange("C2").SetValue(240);
    oWorksheet.GetRange("C3").SetValue(260);
    oWorksheet.GetRange("D2").SetValue(280);
    oWorksheet.GetRange("D3").SetValue(280);
    
    // Add a 3D bar chart
    var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);
    
    // Set chart title
    oChart.SetTitle("Financial Overview", 13);
    
    // Set series fill colors
    var oFill1 = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
    oChart.SetSeriesFill(oFill1, 0, false); // Series 1
    
    var oFill2 = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
    oChart.SetSeriesFill(oFill2, 1, false); // Series 2
    
    // Set minor vertical gridlines
    var oStroke = Api.CreateStroke(1 * 5000, Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)));
    oChart.SetMinorVerticalGridlines(oStroke);
}
```