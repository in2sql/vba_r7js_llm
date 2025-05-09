### Description
This code sets values in specific cells, creates a 3D bar chart, sets the chart's title and series names, and applies specific fill colors to the chart series.

### Описание
Этот код задаёт значения в конкретных ячейках, создаёт 3D столбчатую диаграмму, устанавливает заголовок диаграммы и названия рядов данных, а также применяет определённые цвета заливки к рядам диаграммы.

```javascript
// This example sets values in cells, creates a 3D bar chart, sets titles and series names, and applies fill colors.
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
oChart.SetSeriaName("Projected Sales", 0);
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false); 
```

```vba
' This code sets values in cells, creates a 3D bar chart, sets the chart's title and series names, and applies specific fill colors to the chart series.
Sub CreateFinancialChart()
    ' Set values in cells
    With ActiveSheet
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
    End With
    
    ' Add a 3D Bar Chart
    Dim oChart As Chart
    Set oChart = ActiveSheet.Shapes.AddChart2(240, xl3DBarClustered, 100, 70, 200, 150).Chart ' Chart Type 240 corresponds to xl3DBarClustered
    oChart.SetSourceData Source:=ActiveSheet.Range("A1:D3")
    
    ' Set chart title
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "Financial Overview"
    oChart.ChartTitle.Font.Size = 13
    
    ' Set series name
    oChart.SeriesCollection(1).Name = "Projected Sales"
    
    ' Set series fill colors
    oChart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' Set first series fill color
    oChart.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Set second series fill color
End Sub
```