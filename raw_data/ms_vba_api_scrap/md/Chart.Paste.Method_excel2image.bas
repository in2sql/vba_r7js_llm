Attribute VB_Name = "Excel2Image"
Sub ExcelToImage()
    'Code from officetricks.com
    '[Convert Excel to Jpg Image - VBA Code](https://officetricks.com/convert-range-excel-to-image-vba/)
    Dim sImageFilePath As String
    Dim imageRng As Range
    
    Dim Ws As Worksheet
    For Each Ws In Worksheets
        Ws.Activate
        
        Set imageRng = ActiveSheet.UsedRange
        'sImageFilePath = ActiveSheet.Name & "_" & VBA.Format(VBA.Now, "yyyymmdd_hhnnss_aaa") & ".jpg"
        sImageFilePath = ActiveSheet.Name & "_" & ThisWorkbook.Name & ".jpg"
        
        'Create Temporary workbook to hold image
        Dim wbTemp As Workbook
        Set wbTemp = Workbooks.Add(1)
        
        'Copy image & Save to new file
        imageRng.CopyPicture xlScreen, xlPicture
        wbTemp.Activate
        With wbTemp.Worksheets("Sheet1").ChartObjects.Add(imageRng.Left, imageRng.Top, imageRng.Width, imageRng.Height)
            .Activate
            .Chart.Paste
            .Chart.Export Filename:=sImageFilePath, FilterName:="jpg"
        End With

        'Close Temp workbook
        wbTemp.Close False
        Set wbTemp = Nothing
    Next Ws
End Sub