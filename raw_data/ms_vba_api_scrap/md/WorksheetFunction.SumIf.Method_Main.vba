Sub snCompletionDays()
    'From service note report, the macro creates SN completion days report
    
    'Set the variables
    Dim w1, w2 As Worksheet
    Dim i1, i2, lastR1, lastR2 As Integer
    
    Set w1 = ActiveWorkbook.Sheets(1)
    Set w2 = ActiveWorkbook.Sheets.Add(after:=Sheets(1))
    
    With w1
        'Find the last row
        lastR1 = .Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
        
        'Delete unapproved notes and extract date data from it
        For i1 = lastR1 To 4 Step -1
            If Trim(.Cells(i1, "G").Value) <> "Approved" Then
                .Rows(i1).EntireRow.Delete
            Else
                .Cells(i1, "E").Value = Left(.Cells(i1, "E").Value, 10)
                .Cells(i1, "N").Value = Left(.Cells(i1, "N").Value, 10)
                .Cells(i1, "O").Formula = "=N" & i1 & "-E" & i1
            End If
        Next i1
        
        'Update the last row variable
        lastR1 = .Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
        
        .Range("M4:M" & lastR1).AdvancedFilter Action:=xlFilterCopy, copytorange:=w2.Range("A1"), Unique:=True
    End With
    
    'Create the report in a new sheet
    With w2
        lastR2 = .Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
        
        'Fill the data
        For i2 = lastR2 To 2 Step -1
            .Cells(i2, "B").Value = Application.WorksheetFunction.SumIf(w1.Range("M:M"), .Cells(i2, "A"), w1.Range("O:O"))
            .Cells(i2, "C").Value = Application.WorksheetFunction.CountIf(w1.Range("M:M"), .Cells(i2, "A"))
            .Cells(i2, "D").Value = .Cells(i2, "B").Value / .Cells(i2, "C").Value
            If .Cells(i2, "D").Value > 1 Then .Cells(i2, "D").Interior.Color = RGB(234, 84, 84)
        Next i2
        .Range("A2:D" & lastR2).Sort Key1:=.Range("D2"), order1:=xlDescending
        
        'Add and format column names
        With .Range("A1:D1")
            .Value = Array("DSP", "Total Comp. Days", "SN #", "Mean Comp. Days")
            .Font.Bold = True
            .Interior.Color = RGB(153, 226, 224)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
        
        'Calculate the total
        .Cells(lastR2 + 1, "A").Value = "TOTAL"
        .Cells(lastR2 + 1, "B").Formula = "=sum(B2:B" & lastR2 & ")"
        .Cells(lastR2 + 1, "C").Formula = "=sum(C2:C" & lastR2 & ")"
        .Cells(lastR2 + 1, "D") = "=B" & lastR2 + 1 & "/C" & lastR2 + 1
        
        'Format the total row
        With .Range("A" & lastR2 + 1 & ":D" & lastR2 + 1)
            .Font.Bold = True
            .Interior.Color = RGB(110, 191, 63)
            .Borders(xlEdgeTop).LineStyle = xlDouble
        End With
        
        'Complete ther formatting
        .Columns("D").NumberFormat = "0.00"
        .Cells.WrapText = False
        .Columns("A").AutoFit
        .Columns("B:D").ColumnWidth = 16.45
        
    End With
End Sub
