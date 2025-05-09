**English:**  
This script populates a worksheet with financial data, creates a 3D bar chart titled "Financial Overview," adds a series for "Cost price," and sets specific colors for each series.

**Русский:**  
Этот скрипт заполняет лист финансовыми данными, создает 3D-диаграмму с заголовком "Финансовый обзор", добавляет серию для "Себестоимости" и устанавливает определенные цвета для каждой серии.

```javascript
// This example adds a new series to the chart.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
oWorksheet.GetRange("B1").SetValue(2014); // Set value 2014 in cell B1
oWorksheet.GetRange("C1").SetValue(2015); // Set value 2015 in cell C1
oWorksheet.GetRange("D1").SetValue(2016); // Set value 2016 in cell D1
oWorksheet.GetRange("A2").SetValue("Projected Revenue"); // Set header in A2
oWorksheet.GetRange("A3").SetValue("Estimated Costs"); // Set header in A3
oWorksheet.GetRange("A4").SetValue("Cost price"); // Set header in A4
oWorksheet.GetRange("B2").SetValue(200); // Set value in B2
oWorksheet.GetRange("B3").SetValue(250); // Set value in B3
oWorksheet.GetRange("B4").SetValue(50); // Set value in B4
oWorksheet.GetRange("C2").SetValue(240); // Set value in C2
oWorksheet.GetRange("C3").SetValue(260); // Set value in C3
oWorksheet.GetRange("C4").SetValue(120); // Set value in C4
oWorksheet.GetRange("D2").SetValue(280); // Set value in D2
oWorksheet.GetRange("D3").SetValue(280); // Set value in D3
oWorksheet.GetRange("D4").SetValue(160); // Set value in D4
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000); // Add a 3D bar chart
oChart.SetTitle("Financial Overview", 13); // Set chart title with font size 13
oChart.AddSeria("Cost price", "'Sheet1'!$B$4:$D$4"); // Add "Cost price" series to the chart
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)); // Create a solid fill color
oChart.SetSeriesFill(oFill, 0, false); // Set fill color for the first series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create another solid fill color
oChart.SetSeriesFill(oFill, 1, false); // Set fill color for the second series
```

```vba
' This example populates a worksheet with financial data, creates a 3D bar chart,
' adds a series for "Cost price," and sets specific colors for each series.

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet ' Get the active worksheet
    
    ' Populate headers
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    ws.Range("A4").Value = "Cost price"
    
    ' Populate data
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("B4").Value = 50
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("C4").Value = 120
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    ws.Range("D4").Value = 160
    
    ' Add a 3D Bar Chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=375, Top:=50, Height:=225)
    With chartObj.Chart
        .ChartType = xl3DBar ' Set chart type to 3D Bar
        .SetSourceData Source:=ws.Range("'Sheet1'!$A$1:$D$3")
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview" ' Set chart title
        .ChartTitle.Font.Size = 13 ' Set title font size
        
        ' Add "Cost price" series
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "Cost price"
        .SeriesCollection(2).Values = ws.Range("'Sheet1'!$B$4:$D$4")
        
        ' Set fill color for first series
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
        
        ' Set fill color for second series
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    End With
End Sub
```