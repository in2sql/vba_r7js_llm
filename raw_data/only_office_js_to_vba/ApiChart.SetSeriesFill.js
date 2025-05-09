### Description / Описание

**English:**  
This code sets specific values in cells, creates a 3D bar chart titled "Financial Overview," and applies specific fill colors to the chart series.

**Russian:**  
Этот код устанавливает определенные значения в ячейки, создает 3D столбчатую диаграмму с заголовком "Финансовый обзор" и применяет определенные цвета заливки к сериям диаграммы.

```javascript
// JavaScript Code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set header values for years
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set labels for projected revenue and estimated costs
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set data values for projected revenue
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("D2").SetValue(280);

// Set data values for estimated costs
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("D3").SetValue(280);

// Add a 3D bar chart to the worksheet
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);

// Set the title of the chart
oChart.SetTitle("Financial Overview", 13);

// Create and set the fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set the fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```

```vba
' VBA Code Equivalent

Sub CreateFinancialChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim oFill As Object
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set header values for years
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    
    ' Set labels for projected revenue and estimated costs
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    
    ' Set data values for projected revenue
    ws.Range("B2").Value = 200
    ws.Range("C2").Value = 240
    ws.Range("D2").Value = 280
    
    ' Set data values for estimated costs
    ws.Range("B3").Value = 250
    ws.Range("C3").Value = 260
    ws.Range("D3").Value = 280
    
    ' Add a 3D bar chart to the worksheet
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=375, Top:=50, Height:=225)
    With chartObj.Chart
        .ChartType = xl3DBarClustered
        .SetSourceData Source:=ws.Range("'Sheet1'!$A$1:$D$3")
        .HasTitle = True
        .ChartTitle.Text = "Financial Overview"
        
        ' Set fill color for the first series
        With .SeriesCollection(1).Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(51, 51, 51)
            .Solid
        End With
        
        ' Set fill color for the second series
        With .SeriesCollection(2).Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 111, 61)
            .Solid
        End With
    End With
End Sub
```