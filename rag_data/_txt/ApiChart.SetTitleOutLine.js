**English: This code populates an Excel worksheet with financial data, creates a 3D bar chart based on the data, and applies specific styling to the chart elements.**

**Russian: Этот код заполняет рабочий лист Excel финансовыми данными, создает 3D столбчатую диаграмму на основе данных и применяет определенное форматирование к элементам диаграммы.**

```vba
' VBA Code

Sub CreateFinancialChart()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
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
    
    ' Add a 3D Bar Chart
    Dim oChart As ChartObject
    Set oChart = oWorksheet.ChartObjects.Add(Left:=200, Width:=360, Top:=100, Height:=250)
    With oChart.Chart
        .SetSourceData Source:=oWorksheet.Range("A1:D3")
        .ChartType = xlBarClustered ' 3D bar chart equivalent
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        .ChartTitle.Font.Size = 13
        
        ' Set series fill colors
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
        
        ' Set chart title outline
        With .ChartTitle.Format.Line
            .Visible = msoTrue
            .Weight = 0.5
            .ForeColor.RGB = RGB(51, 51, 51)
        End With
    End With
End Sub
```

```javascript
// OnlyOffice JS Code

// This example sets the outline to the chart title.
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

// Add a 3D Bar Chart
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);

// Set chart title
oChart.SetTitle("Financial Overview", 13);

// Set series fill colors
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Set chart title outline
var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
oChart.SetTitleOutLine(oStroke);
```