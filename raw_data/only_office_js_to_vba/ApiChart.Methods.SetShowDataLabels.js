**Описание:**

*This script populates an Excel worksheet with financial data for the years 2014, 2015, and 2016, then creates a 3D bar chart titled "Financial Overview". It customizes data labels and sets specific fill colors for each data series.*

*Этот скрипт заполняет рабочий лист Excel финансовыми данными за 2014, 2015 и 2016 годы, затем создает 3D столбчатую диаграмму с заголовком "Financial Overview". Он настраивает подписи данных и устанавливает определенные цвета заливки для каждого ряда данных.*

**JavaScript Code:**

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set header values
oWorksheet.GetRange("B1").SetValue(2014);
oWorksheet.GetRange("C1").SetValue(2015);
oWorksheet.GetRange("D1").SetValue(2016);

// Set row labels
oWorksheet.GetRange("A2").SetValue("Projected Revenue");
oWorksheet.GetRange("A3").SetValue("Estimated Costs");

// Set data values
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

// Customize data labels
oChart.SetShowDataLabels(false, false, true, false);

// Create and set fill color for the first series
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
oChart.SetSeriesFill(oFill, 0, false);

// Create and set fill color for the second series
oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
oChart.SetSeriesFill(oFill, 1, false); 
```

**Excel VBA Code:**

```vba
Sub CreateFinancialChart()
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set header values
    ws.Range("B1").Value = 2014
    ws.Range("C1").Value = 2015
    ws.Range("D1").Value = 2016
    
    ' Set row labels
    ws.Range("A2").Value = "Projected Revenue"
    ws.Range("A3").Value = "Estimated Costs"
    
    ' Set data values
    ws.Range("B2").Value = 200
    ws.Range("B3").Value = 250
    ws.Range("C2").Value = 240
    ws.Range("C3").Value = 260
    ws.Range("D2").Value = 280
    ws.Range("D3").Value = 280
    
    ' Add a 3D bar chart
    Dim cht As Chart
    Set cht = ws.Shapes.AddChart2(251, xlBarClustered, 100, 70, 360, 360).Chart ' Width and Height scaled as needed
    
    ' Set chart data source
    cht.SetSourceData Source:=ws.Range("A1:D3")
    
    ' Set chart title
    cht.HasTitle = True
    cht.ChartTitle.Text = "Financial Overview"
    cht.ChartTitle.Font.Size = 13
    
    ' Customize data labels
    Dim ser As Series
    For Each ser In cht.SeriesCollection
        ser.HasDataLabels = True
        With ser.DataLabels
            .ShowValue = False
            .ShowCategoryName = False
            .ShowSeriesName = True
            .ShowPercentage = False
        End With
    Next ser
    
    ' Set fill color for the first series
    cht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(51, 51, 51)
    
    ' Set fill color for the second series
    cht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 111, 61)
End Sub
```