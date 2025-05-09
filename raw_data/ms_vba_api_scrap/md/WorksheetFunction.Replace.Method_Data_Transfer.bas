Attribute VB_Name = "Data_Transfer"
Option Explicit

Sub Fetch_Data_PPT_to_Excel()
'This procedure is extract data from PPT to Excel

    Dim Obj_App_PPT As Object: Dim Obj_Slide As Object: Dim Obj_Shape As Object: Dim Obj_FilePPt As Object
    Dim Source_Path As String: Dim str_Country As String: Dim ValidDate As Variant
    Dim LastRow As Integer: Dim Wrkbk As Workbook: Dim WrkSht As Worksheet
    Dim Entry_Exit As String: Dim int_Sln As Integer: Dim i As Integer: Dim InvalideSlideCount As Integer: Dim SlideNum As String
    Dim Height_Admission As String: Dim chk As Integer: Dim SlidesCnt As Integer
    Dim Vaccination_Req As String
    Dim Penalties As String
    Dim Quarantine_Isolation As String
    Dim Imapact_On_Exiting As String
    Dim Rng As Range: Dim Cell As Range: Dim Penalties_LastRow As Integer: Dim Rng_Col As Range: Dim Cell_Col As Range:
    Dim Destination_Path As String
    Dim FS As Object
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Check whether File Exist or not
    
    Source_Path = "C:\Users\shashidhara.kb\OneDrive - EY\Documents\PPT Automation\GS data mockup v1.pptx"
    
    If Dir(Source_Path) = "" Then
        MsgBox ("File Doesn't exist, Exiting Macro!")
        GoTo ExitHere
    End If
    
    'Add Workbook
    Set Wrkbk = Workbooks.Add
    ThisWorkbook.Sheets("Data").Copy After:=Wrkbk.Sheets(1)
    Set WrkSht = Wrkbk.Sheets("Data")
        WrkSht.Visible = True
    Wrkbk.Sheets("Sheet1").Delete
    
    'Assign object
    Set Obj_App_PPT = CreateObject("PowerPoint.Application")
        Obj_App_PPT.Visible = True
        
    Set Obj_FilePPt = Obj_App_PPT.Presentations.Open(Source_Path)
    ThisWorkbook.Sheets("Macro").Range("P6:P7").ClearContents
    SlidesCnt = Obj_FilePPt.Slides.Count
    'Loop through each slide
    For Each Obj_Slide In Obj_FilePPt.Slides
        
        'Check for valid slide
        If Obj_Slide.Shapes.Count <> 6 Then
            SlideNum = SlideNum & ", " & Obj_Slide.SlideNumber
            InvalideSlideCount = InvalideSlideCount + 1
            chk = 1
            GoTo StartHere
        End If
        
        'Loop through each shape
            i = 0
            For Each Obj_Shape In Obj_Slide.Shapes
            
                'TextBox
                If Obj_Shape.Type = 14 Then
                    'Country
                 
                    If i = 3 Then
                        str_Country = Trim(WorksheetFunction.Clean(Obj_Shape.TextFrame.TextRange.Text))
                    End If
                    
                    'Date
                    If i = 4 Then
                        ValidDate = Right(Obj_Shape.TextFrame.TextRange.Text, Len(Obj_Shape.TextFrame.TextRange.Text) - WorksheetFunction.Find(":", Obj_Shape.TextFrame.TextRange.Text, 15))
                    End If
                End If
                
                'Table
                If Obj_Shape.Type = 19 Then
                    'Clear Macro file previous data
                    With ThisWorkbook.Sheets("Table")
                        .Cells.Delete Shift:=xlUp
                    End With
                
                    Obj_Shape.Copy
                    ThisWorkbook.Sheets("Table").Activate
                    
                    ThisWorkbook.Sheets("Table").Range("A1").Select
                    ActiveSheet.Paste
                End If
               i = i + 1
            Next Obj_Shape
            
            '@@@@@@@@@@@@@@@@@@@ Sheet formate
            With ThisWorkbook.Sheets("Table")
            
                '********************************Entry and Exit restrictions********************
                LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
                    
                 Entry_Exit = ""
                If LastRow = 1 Then
                    Entry_Exit = ""
                Else
                    Set Rng = .Range("A2:A" & LastRow)
                    'Iteration Through each cell
                    For Each Cell In Rng
                        Entry_Exit = Entry_Exit & Trim(WorksheetFunction.Replace(Cell.Value, 1, 1, ""))
                    Next Cell
                End If
                
                '*********************Heightened admission requirements***************************
                LastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
                Height_Admission = ""
                If LastRow = 1 Then
                    Height_Admission = ""
                Else
                    Set Rng = .Range("C2:C" & LastRow)
                    'Iteration Through each cell
                    For Each Cell In Rng
                        Height_Admission = Height_Admission & Trim(WorksheetFunction.Replace(Cell.Value, 1, 1, ""))
                    Next Cell
                End If
                
                '************* Vaccination requirements & Penalties for non-compliance *********************
                'Vaccination requirements
                Vaccination_Req = ""
                LastRow = .Cells(.Rows.Count, "E").End(xlUp).Row
 
                Set Rng = .Range("E2:E" & LastRow)
                'Iteration Through each cell
                For Each Cell In Rng
                    If Cell.Interior.Color = RGB(255, 230, 0) Then
                        LastRow = Cell.Row - 1
                        Penalties_LastRow = Cell.Row + 1
                        Exit For
                    End If
                Next Cell
                
                Set Rng_Col = .Range("E2:E" & LastRow)
                For Each Cell_Col In Rng_Col
                    Vaccination_Req = Vaccination_Req & Trim(WorksheetFunction.Replace(Cell_Col.Value, 1, 1, ""))
                Next Cell_Col
            
            'Penalties for non-compliance
                Penalties = ""
                LastRow = .Cells(.Rows.Count, "E").End(xlUp).Row
                Set Rng = .Range("E" & Penalties_LastRow & ":E" & LastRow)

                    For Each Cell In Rng
                        Penalties = Penalties & Trim(WorksheetFunction.Replace(Cell.Value, 1, 1, ""))
                    Next Cell
                    
            '********************Quarantine / isolation requirements********************
                LastRow = .Cells(.Rows.Count, "G").End(xlUp).Row
                Quarantine_Isolation = ""
                If LastRow = 1 Then
                    Quarantine_Isolation = ""
                Else
                    Set Rng = .Range("G2:G" & LastRow)
                    'Iteration Through each cell
                    For Each Cell In Rng
                        Quarantine_Isolation = Quarantine_Isolation & Trim(WorksheetFunction.Replace(Cell.Value, 1, 1, ""))
                    Next Cell
                End If
                
                '********************Impact on existing visas, and new visa issuance********************
                LastRow = .Cells(.Rows.Count, "I").End(xlUp).Row
                Imapact_On_Exiting = ""
                If LastRow = 1 Then
                    Imapact_On_Exiting = ""
                Else
                    Set Rng = .Range("I2:I" & LastRow)
                    'Iteration Through each cell
                    For Each Cell In Rng
                        Imapact_On_Exiting = Imapact_On_Exiting & Trim(WorksheetFunction.Replace(Cell.Value, 1, 1, ""))
                    Next Cell
                End If
                
            End With
            
            'Add Data to Output sheet
            int_Sln = 0
            
            With WrkSht
                For i = 1 To 6
                    LastRow = .Cells(.Rows.Count, "D").End(xlUp).Row
                    .Range("A" & LastRow + 1).Value = str_Country
                    .Range("B" & LastRow + 1).Value = "Immigration"
                    .Range("C" & LastRow + 1).Value = "Immigration"
                    .Range("D" & LastRow + 1).Value = int_Sln + i
                
                    Select Case i
                        Case 1:
                            .Range("E" & LastRow + 1).Value = "Entry & exit restrictions"
                            .Range("G" & LastRow + 1).Value = Entry_Exit
                        Case 2:
                            .Range("E" & LastRow + 1).Value = "Heightened admission requirements"
                            .Range("G" & LastRow + 1).Value = Height_Admission
                        Case 3:
                            .Range("E" & LastRow + 1).Value = "Vaccination requirements & considerations"
                            .Range("G" & LastRow + 1).Value = Vaccination_Req
                        Case 4:
                            .Range("E" & LastRow + 1).Value = "Quarantine & isolation requirements"
                            .Range("G" & LastRow + 1).Value = Quarantine_Isolation
                        Case 5:
                            .Range("E" & LastRow + 1).Value = "Impact on existing visas and new visa issuance"
                            .Range("G" & LastRow + 1).Value = Imapact_On_Exiting
                        Case 6:
                            .Range("E" & LastRow + 1).Value = "Penalties for non-compliance"
                            .Range("G" & LastRow + 1).Value = Penalties
                    End Select
                    
                    .Range("H" & LastRow + 1).Value = ValidDate
                    .Range("I" & LastRow + 1).Value = "All"
                    .Range("J" & LastRow + 1).Value = "Manual"
                    .Range("K" & LastRow + 1).Value = "Country"

                Next i
            End With
StartHere:
    Next Obj_Slide
    
    'Formate Final Output
        With WrkSht
            .Activate
            .Columns("G:G").WrapText = True
            .Columns("G:G").VerticalAlignment = xlTop
            LastRow = .Cells(.Rows.Count, "D").End(xlUp).Row
            .ListObjects.Add(xlSrcRange, Range("$A$1:$K$" & LastRow), , xlYes).Name = "Table1"
            .Cells.VerticalAlignment = xlCenter
            .Columns("G:G").ColumnWidth = 60
            .Columns("E:E").ColumnWidth = 40
            .Cells.EntireRow.AutoFit
        End With
        
        If chk = 1 Then
            ThisWorkbook.Sheets("Macro").Range("P6").Value = SlidesCnt - InvalideSlideCount
            ThisWorkbook.Sheets("Macro").Range("P7").Value = InvalideSlideCount
        Else
            ThisWorkbook.Sheets("Macro").Range("P6").Value = SlidesCnt
            ThisWorkbook.Sheets("Macro").Range("P7").Value = "All are as per Standard"
        End If
        
        Set FS = CreateObject("Scripting.FileSystemObject")
        Source_Folder = ThisWorkbook.Path & "\" & "Output_File_" & Format(Now(), "DD-MMM-YYYY hh mm AMPM") & ".xlsx"
        Wrkbk.SaveAs (Source_Folder)
        
                   '***** Sharepoint ******
'        Root = "@SSL\DavWWWRoot"
'        Destination_Folder = "\\us.eyonespace.ey.com@SSL\DavWWWRoot\Sites\bcfaa042739b4ca7bed7085500291408\AI  R Share Drive Transition\CE File\"
'        FS.CopyFile (Source_Folder), Destination_Folder

            'Obj_FilePPt.Close
          Obj_App_PPT.Quit
          
          MsgBox "Completed"
ExitHere:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
