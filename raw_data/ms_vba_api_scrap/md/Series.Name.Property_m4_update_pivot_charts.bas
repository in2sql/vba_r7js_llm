Attribute VB_Name = "m4_update_pivot_charts"
Sub refresh_plots()

Dim wb As Workbook
Dim ws_admin As Worksheet
Dim ws As Worksheet

Set wb = ThisWorkbook
Set ws_admin = wb.Sheets("Admin")

Dim colorMap
Set colorMap = CreateObject("Scripting.Dictionary")

Dim i As Integer

For i = 2 To 10
    
    If Not colorMap.Exists(Range("E" & i).Value) Then
        colorMap.Add Range("E" & i).Value, Range("F" & i).Value
        
    End If
    

Next i

Dim sheetNames As Variant
Dim sheetName As Variant

sheetNames = Array("Plot1", "Plot2")

Dim hex_color_code As String
Dim r, g, b As Integer

Dim chartObj As ChartObject

For Each sheetName In sheetNames
    Set ws = wb.Sheets(sheetName)
    
    For Each chartObj In ws.ChartObjects
        
        chartObj.Chart.HasTitle = True
        chartObj.Chart.ChartTitle.Text = ws.Range("E1").Value
        
        For Each Series In chartObj.Chart.SeriesCollection
            
            itemName = Series.Name
            hex_color_code = colorMap(itemName)
            
            r = Val("&H" & Mid(hex_color_code, 1, 2))
            g = Val("&H" & Mid(hex_color_code, 3, 2))
            b = Val("&H" & Mid(hex_color_code, 5, 2))

            Series.Format.Fill.ForeColor.RGB = RGB(r, g, b)
        
        Next Series
        
    Next chartObj
    
Next sheetName

End Sub

