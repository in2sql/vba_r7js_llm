Attribute VB_Name = "Module1"
Sub TurnOffFuntionality()
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.DisplayScrollBars = False

End Sub

Sub TurnOnFuntionality()
Application.Calculation = xlCalculationAutomatic
Application.DisplayStatusBar = True
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayScrollBars = True
End Sub


Sub cleanedData()
Dim file_name As String, fldr As FileDialog
Dim sItems As String
Dim LR_Source As Long, LR_Destination_beforePaste As Long, LR_Destination_afterPaste As Long

Dim tempsheet As Worksheet

Dim sourceworksheet As Worksheet, destiworksheet As Worksheet, data_input As Worksheet, data_qc As Worksheet


Set destiworksheet = ThisWorkbook.Sheets("test")

'Debug.Print destiworksheet.Name
Set fldr = Application.FileDialog(msoFileDialogFilePicker)

TurnOffFuntionality


With fldr
    .Title = "Select files"
    .Filters.Clear
    .Filters.Add "Excel Files", "*.xlsx, *.xls, *.xlsm"
    .FilterIndex = 1
    .Show
    For Each sItem In .SelectedItems
        sItems = sItem
        
        Debug.Print "Currently Processing: "; sItems
        Dim book As Workbook
        Set book = Workbooks.Open(Filename:=sItem, ReadOnly:=True)
        Dim batch_id As String
        Set sourceworksheet = book.Sheets("CSV OUTPUT")
        Set data_input = book.Sheets("DATA_INPUT")
        Set data_qc = book.Sheets("DATA_QC")
        
        Dim data As Range
        LR_Source = sourceworksheet.Columns("A").Find("*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows, LookIn:=xlValues).Row
        
        batch_id = sourceworksheet.Range("F2").Value
        
        Set data = sourceworksheet.Range("A2:I" & LR_Source)
        sourceworksheet.Range("A2:A" & LR_Source).Copy
        
        
        
        LR_Destination_beforePaste = destiworksheet.Range("S" & Rows.Count).End(xlUp).Row
        
   '     Debug.Print "LR before Paste" & LR_Destination_beforePaste
        
        ' Paste Sample Names
        
        destiworksheet.Range("S" & LR_Destination_beforePaste + 1, "S" & LR_Destination + LR_Source - 1).PasteSpecial Paste:=xlPasteValues ' paste sample name
        
        
        LR_Destination = destiworksheet.Range("S" & Rows.Count).End(xlUp).Row
 '       Debug.Print "LR After Paste" & LR_Destination
        
        'Remove duplicates
        
        destiworksheet.Range("S" & LR_Destination_beforePaste + 1, "S" & LR_Destination + LR_Source - 1).RemoveDuplicates Columns:=1, Header:=xlNo
        LR_Destination_afterPaste = destiworksheet.Range("S" & Rows.Count).End(xlUp).Row
        
      '  Debug.Print "Value in F2 batchID: "; batch_id
     '   Debug.Print "Last Row After Paste & rm dup: "; LR_Destination_afterPaste
        
        destiworksheet.Range("B" & LR_Destination_beforePaste + 1, "B" & LR_Destination_afterPaste).Value = data_input.Range("B25").Value
        destiworksheet.Range("H" & LR_Destination_beforePaste + 1, "H" & LR_Destination_afterPaste).Value = data_qc.Range("C3").Value
        
        
        ' set the OPERATOR and Instrument Name
        
        destiworksheet.Range("A" & LR_Destination_beforePaste + 1, "A" & LR_Destination_afterPaste).Value = batch_id
        
        ' final range
        Dim final_range As Range
        Set final_range = destiworksheet.Range("A" & LR_Destination_beforePaste + 1, "AU" & LR_Destination_afterPaste)
        
        
            Dim iRow As Long
                    
            For iRow = 1 To final_range.Rows.Count
            '    Debug.Print "------> Value: "; final_range.Cells(iRow, 19).Value; " -- Address: "; final_range.Cells(iRow, 1).Address
                Dim tar As Range
                Dim targetIndex As Long
                Dim channelResultsIndex As Long
                
                targetIndex = 2 ' Initialize target index
                channelResultsIndex = 2
                panelIndex = 1
                For Each tar In ThisWorkbook.Sheets("constant").Range("A1").CurrentRegion
                    Dim i As Long
                    'Debug.Print (tar.Value)
                    
                    For i = 1 To data.Rows.Count
                        If (data.Cells(i, 1) = final_range.Cells(iRow, 19).Value) And (data.Cells(i, 2) = tar.Value) Then ' checks if target name and sample name is the same
                            ' Assign value to corresponding column based on target index
                            If data.Cells(i, 4).Value = "" Then
                                final_range.Cells(iRow, 20 + channelResultsIndex).Value = "Undetermined"
                            Else
                                final_range.Cells(iRow, 20 + channelResultsIndex).Value = data.Cells(i, 4) ' Cts column
                            End If
                            
                            final_range.Cells(iRow, 21 + channelResultsIndex).Value = data.Cells(i, 5) ' Results column
                            If panelIndex Mod 4 = 1 Then
                            final_range.Cells(iRow, 19 + channelResultsIndex).Value = data.Cells(i, 3) ' Well column
                            
                            End If
                            
                        
                        
                        
                        End If

                    Next i
                    
                    ' Adjust target index and channelResultsIndex
                    If panelIndex Mod 4 <> 0 Then
                        targetIndex = targetIndex + 2 ' Increment target index
                        channelResultsIndex = channelResultsIndex + 2
                    Else
                        targetIndex = targetIndex + 3 ' Skip one column
                        channelResultsIndex = channelResultsIndex + 3
                    End If
                    panelIndex = panelIndex + 1
                    
                Next tar
                
            Next iRow
       
        
        
        book.Close SaveChanges:=False
    Next sItem
End With


TurnOnFuntionality

    

End Sub


Sub t()

Debug.Print Count
End Sub
