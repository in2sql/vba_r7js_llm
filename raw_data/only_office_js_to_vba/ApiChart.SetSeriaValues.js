### Description / Описание
**English:** This code sets specific cell values in a worksheet and creates a 3D bar chart titled "Financial Overview" with customized series and data labels.
**Русский:** Этот код устанавливает определенные значения ячеек в рабочем листе и создает 3D столбчатую диаграмму с заголовком "Финансовый обзор" с настраиваемыми сериями и метками данных.

```javascript
// JavaScript OnlyOffice API Code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set values in specific cells
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");
oWorksheet.GetRange("B2").SetValue(200);
oWorksheet.GetRange("B3").SetValue(250);
oWorksheet.GetRange("B4").SetValue(260);
oWorksheet.GetRange("C2").SetValue(240);
oWorksheet.GetRange("C3").SetValue(260);
oWorksheet.GetRange("C4").SetValue(270);
oWorksheet.GetRange("D2").SetValue(280);
oWorksheet.GetRange("D3").SetValue(280);
oWorksheet.GetRange("D4").SetValue(300);

// Add a 3D bar chart to the worksheet
var oChart = oWorksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);

// Set the chart title
oChart.SetTitle("Financial Overview", 13);

// Set series values for the chart
oChart.SetSeriaValues("'Sheet1'!$B$4:$D$4", 1);

// Configure data label visibility for each series point
oChart.SetShowPointDataLabel(1, 0, false, false, true, false);
oChart.SetShowPointDataLabel(1, 1, false, false, true, false);
oChart.SetShowPointDataLabel(1, 2, false, false, true, false);

// Create and set the fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set the fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false);
```

```vba
' Excel VBA Equivalent Code

Sub CreateFinancialOverviewChart()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set values in specific cells
    oWorksheet.Range("B1").Value = 2014
    oWorksheet.Range("C1").Value = 2015
    oWorksheet.Range("D1").Value = 2016
    oWorksheet.Range("A2").Value = "Projected Revenue"
    oWorksheet.Range("A3").Value = "Estimated Costs"
    oWorksheet.Range("B2").Value = 200
    oWorksheet.Range("B3").Value = 250
    oWorksheet.Range("B4").Value = 260
    oWorksheet.Range("C2").Value = 240
    oWorksheet.Range("C3").Value = 260
    oWorksheet.Range("C4").Value = 270
    oWorksheet.Range("D2").Value = 280
    oWorksheet.Range("D3").Value = 280
    oWorksheet.Range("D4").Value = 300
    
    ' Add a 3D bar chart to the worksheet
    Dim oChart As ChartObject
    Set oChart = oWorksheet.ChartObjects.Add(Left:=200, Width:=360, Top:=100, Height:=270)
    oChart.Chart.ChartType = xlBar3DClustered
    
    ' Set the chart data range
    oChart.Chart.SetSourceData Source:=oWorksheet.Range("'Sheet1'!$A$1:$D$3")
    
    ' Set the chart title
    oChart.Chart.HasTitle = True
    oChart.Chart.ChartTitle.Text = "Financial Overview"
    oChart.Chart.ChartTitle.Font.Size = 13
    
    ' Set series values for the chart
    With oChart.Chart.SeriesCollection.NewSeries
        .Values = oWorksheet.Range("'Sheet1'!$B$4:$D$4")
        .Name = "Financial Data"
    End With
    
    ' Configure data label visibility for each series point
    Dim ser As Series
    Set ser = oChart.Chart.SeriesCollection(1)
    Dim i As Integer
    For i = 1 To ser.Points.Count
        With ser.Points(i).DataLabel
            .ShowBubbleSize = False
            .ShowCategoryName = False
            .ShowValue = True
            .ShowSeriesName = False
        End With
    Next i
    
    ' Set fill color for the first series
    ser.Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
    
    ' Add and set fill color for the second series if exists
    If oChart.Chart.SeriesCollection.Count > 1 Then
        ser = oChart.Chart.SeriesCollection(2)
        ser.Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
    End If
End Sub
```