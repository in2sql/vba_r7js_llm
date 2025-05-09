Attribute VB_Name = "Module1"
' this function requires that you add a reference to excel
' version 97+
Function PrintGrid(Grid As MSFlexGrid, Título As String)
   Dim ExcelApp As Excel.Application 'prepare excel to be used
   Dim ExcelWBk As Excel.Workbook    'prepare an Excel workbook
   Dim ExcelWS As Excel.Worksheet    'prepare an Excel worksheet
   Set ExcelApp = CreateObject("Excel.Application") 'open Excel
   Set ExcelWBk = ExcelApp.Workbooks.Add    'add a new workbook
   Set ExcelWS = ExcelWBk.Worksheets(1)     'select the worksheet 1
   For i% = 1 To Grid.Rows       'export data from grid to excel
       For j% = 1 To Grid.Cols
           ExcelWS.Cells(i%, j%) = Grid.TextMatrix(i% - 1, j% - 1)
       Next j%
   Next i%
   ExcelApp.DisplayAlerts = False 'really need explanation?
   For i% = 1 To Grid.Cols
       'this is the key: It resizes all the columns in the worksheet
       'so it appears to be the same size
       ExcelWS.Columns(i%).ColumnWidth = Grid.ColWidth(i% - 1) / 100
   Next i%
   ' here it starts to prepare printing
   ' will be nice if you made a simple printer selection dialog box
   ' using even the common dialogbox of printing
   ExcelWS.PageSetup.Orientation = xlLandscape
   ExcelWS.PageSetup.FitToPagesWide = 1
   paginas = CInt(Grid.Rows / 25) + 1
   ExcelWS.PageSetup.FitToPagesTall = paginas
   ExcelWS.PageSetup.PrintGridlines = True
   ExcelWS.PageSetup.LeftHeader = "This is a simple grid printing" + vbNewLine + Título
   ExcelWS.PageSetup.RightHeader = "Print date: " + CStr(Date) + vbNewLine
   ExcelWS.PageSetup.PaperSize = xlPaperA4
   ExcelWS.PrintOut ' Prints the worksheet
   ExcelWS.SaveAs Tools.ReadGetTempDir + "temp.xls" ' save the workbook
   ExcelApp.Quit ' and finally quits application.
End Function

