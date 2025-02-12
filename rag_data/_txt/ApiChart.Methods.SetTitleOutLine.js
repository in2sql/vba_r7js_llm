### Description / Описание
**English:** This script populates a worksheet with financial data for the years 2014 to 2016, creates a 3D bar chart titled "Financial Overview," sets custom fill colors for the series, and applies an outline to the chart title.

**Русский:** Этот скрипт заполняет рабочий лист финансовыми данными за 2014–2016 годы, создает 3D-столбчатую диаграмму с заголовком "Финансовый обзор", устанавливает пользовательские цвета заполнения для серий и применяет обводку к заголовку диаграммы.

```javascript
// This example sets the outline to the chart title.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
oWorksheet.GetRange("B1").SetValue(2014); // Set value 2014 in cell B1
oWorksheet.GetRange("C1").SetValue(2015); // Set value 2015 in cell C1
oWorksheet.GetRange("D1").SetValue(2016); // Set value 2016 in cell D1
oWorksheet.GetRange("A2").SetValue("Projected Revenue"); // Set label in A2
oWorksheet.GetRange("A3").SetValue("Estimated Costs"); // Set label in A3
oWorksheet.GetRange("B2").SetValue(200); // Set value 200 in B2
oWorksheet.GetRange("B3").SetValue(250); // Set value 250 in B3
oWorksheet.GetRange("C2").SetValue(240); // Set value 240 in C2
oWorksheet.GetRange("C3").SetValue(260); // Set value 260 in C3
oWorksheet.GetRange("D2").SetValue(280); // Set value 280 in D2
oWorksheet.GetRange("D3").SetValue(280); // Set value 280 in D3

// Add a 3D bar chart with specified parameters
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);
oChart.SetTitle("Financial Overview", 13); // Set chart title with font size 13

// Create and set the fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set the fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);

// Create a stroke for the chart title outline and apply it
var oStroke = Api.CreateStroke(0.5 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
oChart.SetTitleOutLine(oStroke); 
```

```vba
' This example sets the outline to the chart title.

Sub CreateFinancialOverviewChart()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet ' Get the active worksheet
    
    ' Populate the worksheet with data
    ws.Range("B1").Value = 2014 ' Set value 2014 in cell B1
    ws.Range("C1").Value = 2015 ' Set value 2015 in cell C1
    ws.Range("D1").Value = 2016 ' Set value 2016 in cell D1
    ws.Range("A2").Value = "Projected Revenue" ' Set label in A2
    ws.Range("A3").Value = "Estimated Costs" ' Set label in A3
    ws.Range("B2").Value = 200 ' Set value 200 in B2
    ws.Range("B3").Value = 250 ' Set value 250 in B3
    ws.Range("C2").Value = 240 ' Set value 240 in C2
    ws.Range("C3").Value = 260 ' Set value 260 in C3
    ws.Range("D2").Value = 280 ' Set value 280 in D2
    ws.Range("D3").Value = 280 ' Set value 280 in D3
    
    ' Add a 3D bar chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=375, Top:=50, Height:=225)
    With chartObj.Chart
        .SetSourceData Source:=ws.Range("A1:D3") ' Set the data range for the chart
        .ChartType = xlBarClustered ' Set chart type to 3D bar (modify as needed)
        .HasTitle = True ' Enable chart title
        .ChartTitle.Text = "Financial Overview" ' Set chart title
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 13 ' Set font size
        
        ' Set fill color for the first series
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51) ' Dark gray
        
        ' Set fill color for the second series
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61) ' Orange
        
        ' Apply outline to the chart title
        With .ChartTitle.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(51, 51, 51) ' Dark gray
            .Weight = 0.5 ' Set line weight
        End With
    End With
End Sub
```