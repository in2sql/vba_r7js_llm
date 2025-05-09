Attribute VB_Name = "CopySheets"
Public Sub CreateWeeklySheets()
    ' This sub iterates through worksheets starting from the right (most current) to find 1st occuring "Cert" or "COI" in worksheet name
    Dim i As Integer
    Dim certFound As Boolean: certFound = False
    Dim coiFound As Boolean: coiFound = False
    Dim desiredWkshs(1) As Worksheet
    Dim currentWksh As Worksheet
    
    ' iterates through worksheets from right and stores found worksheet in an array
    For i = ActiveWorkbook.Worksheets.Count To 1 Step -1
        Set currentWksh = ActiveWorkbook.Worksheets(i)
        If certFound = True And coiFound = True Then
            Exit For
        End If
        If certFound = False And Right(currentWksh.Name, 4) = "Cert" Then
            Set desiredWkshs(0) = currentWksh
            certFound = True
        End If
        If coiFound = False And Right(currentWksh.Name, 3) = "COI" Then
            Set desiredWkshs(1) = currentWksh
            coiFound = True
        End If
    Next i
    
    ' runs copy sub with the cert sheet
    If Not desiredWkshs(0) Is Nothing Then
        With desiredWkshs(0)
            CopySheet desiredWkshs(0)
        End With
    Else
        MsgBox "Worksheet with ""Cert"" in name not found"
    End If
    
    ' runs the copy sub with the coi sheet
    If Not desiredWkshs(1) Is Nothing Then
        With desiredWkshs(1)
            CopySheet desiredWkshs(1)
        End With
    Else
        MsgBox "Worksheet with ""COI"" in name not found"
    End If
    
End Sub

Private Sub CopySheet(desiredWksh As Worksheet)
    ' This method does the following:
    ' 1) copies the cert or coi sheet
    ' 2) removes the tab color of prior cert sheet and adds tab color to the new cert sheet
    ' 3) disables all filters
    ' 4) ungroups all groups
    ' 5) calls the strikethrough sub
    Dim newName As String
    Dim today As Date: today = Date
    Dim lastMonday As Date: lastMonday = DateAdd("d", -7, today - Weekday(today) + 2)
    Dim thisSunday As Date: thisSunday = today - Weekday(today) + 1
    
    If Right(desiredWksh.Name, 4) = "Cert" Then
        desiredWksh.Tab.Color = RGB(232, 232, 232)
        desiredWksh.Copy After:=Worksheets(Sheets.Count)
        ActiveSheet.Tab.Color = RGB(255, 0, 217)
        ActiveSheet.Name = Format(lastMonday, "MM.DD") & "-" & Format(thisSunday, "MM.DD.YY") & " Cert"
    Else
        desiredWksh.Copy After:=Worksheets(Sheets.Count)
        ActiveSheet.Name = Format(lastMonday, "MM.DD") & "-" & Format(thisSunday, "MM.DD.YY") & " COI"
    End If
    
    ' disables all filters
    On Error Resume Next
        ActiveSheet.ShowAllData
        
    ' ungroups all groups
    On Error Resume Next
        Rows.Ungroup
    
    DeleteStrikethrough
End Sub


Private Sub DeleteStrikethrough()
    ' This method iterates through column A starting below col titles (A2). When it finds a cell in col A with a strikethrough, it deletes the entire row
    
    Dim cell As Range
    Dim iRows As Range
    Dim i As Long
    
    Set iRows = Range("A2", Range("A3").End(xlDown))
    ' Backwards iteration is necessary to maintain correct row number count after a row is deleted
    For i = iRows.Cells.Count To 0 Step -1
        ' Note this line will find MANUALLY strikethrough rows: "If iRows.Item(i).Font.Strikethrough = True Then"
        ' DisplayFormat is used here because strikethrough is formatted by conditional formatting is different
        If iRows.Item(i).DisplayFormat.Font.Strikethrough = True Then
            iRows.Rows(i).EntireRow.Delete
        End If
    Next i
    
End Sub

