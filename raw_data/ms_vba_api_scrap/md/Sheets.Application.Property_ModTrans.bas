Attribute VB_Name = "ModTrans"
Public EX As Excel.Application
Public EXW As Excel.Workbook
'Dim RSHID As ADODB.Recordset
'Dim RsSECHEAD As ADODB.Recordset
'Dim RsSubsko As ADODB.Recordset

Public Constant_Count As Integer
Public Page_Count As Integer
Public Post_Count As Integer

Public Sub Setup_Excel(Optional SCHOOL As String = "INTERNATIONAL SCHOOL OF ASIA AND THE PACIFIC")
    'Set all
    'On Error Resume Next
    Dim FSO, x As String
    'Set Count to 0
    Constant_Count = 0
    Page_Count = 0
    'Set EX = GetObject(, "Excel.Application")
    'If Err.Number <> 0 Then
        Set EX = CreateObject("Excel.Application")
    'End If
    FRMTOR.CDL.DialogTitle = "Create TOR"
    FRMTOR.CDL.Filter = "Excel Tor Format(*.Xls)|*.xls"
    FRMTOR.CDL.ShowSave
    If FRMTOR.CDL.FileName = "" Then Exit Sub
    'select Send schoold
    If UCase(SCHOOL) = UCase("international school of asia and the pacific") Then
    x = App.Path & "\Templates\ISAPTOR.xls"
    Else
    x = App.Path & "\Templates\MCNPTOR.xls"
    End If
    Set FSO = CreateObject("Scripting.filesystemobject")
    If FSO.fileexists(FRMTOR.CDL.FileName) = True Then
    MsgBox "Can't Overwrite this file.", vbCritical, "ERROR"
    Set FSO = Nothing
    Exit Sub
    End If
    FSO.copyfile x, FRMTOR.CDL.FileName, True
    Set EXW = EX.Workbooks.Open(FRMTOR.CDL.FileName)
    Copy_Files
    Exit Sub

End Sub

Sub Write_Head(Page As Integer)
'Page Range
'0=Page 1
'1=Page 2
'2=Page 3
Dim RowCount As Integer
RowCount = 68
With FRMTOR
'First 2---------
        EXW.Worksheets.Application.Cells(6 + (RowCount * Page), 2) = .LBLNAME.Caption
        EXW.Worksheets.Application.Cells(6 + (RowCount * Page), 10) = .Tadd.Text
'----------------
        EXW.Worksheets.Application.Cells(7 + (RowCount * Page), 4) = .TAdmission.Text
        EXW.Worksheets.Application.Cells(7 + (RowCount * Page), 11) = .CCOURSE.Text
'----------------
        EXW.Worksheets.Application.Cells(8 + (RowCount * Page), 3) = .THS.Text
        EXW.Worksheets.Application.Cells(8 + (RowCount * Page), 12) = .tdesc.Text
'----------------
        EXW.Worksheets.Application.Cells(9 + (RowCount * Page), 2) = .CSCHOOL.Text
        EXW.Worksheets.Application.Cells(9 + (RowCount * Page), 13) = .TGrad.Text
'----------------
        EXW.Worksheets.Application.Cells(10 + (RowCount * Page), 8) = .TGEN.Text
        EXW.Worksheets.Application.Cells(10 + (RowCount * Page), 12) = .TSO.Text
'----------------
        EXW.Worksheets.Application.Cells(11 + (RowCount * Page), 6) = .TCred.Text
End With

End Sub

Public Sub Copy_Files()

    Dim msg As String
    
    msg = "Select SCHOOL, SCHOOLYEAR, SEMESTER, COURSE"
    msg = msg & " from GRADING_SYS Where IDNO='" & FRMTOR.TORID
    msg = msg & "' GROUP BY SCHOOL,SCHOOLYEAR, SEMESTER,COURSE"
    msg = msg & " ORDER BY SCHOOLYEAR, SEMESTER"

    With FRMTOR
        Dim strx As String, POSTX As Integer, sqlsub As String
        'Write Header
        Write_Head 0
        Set .RsSY = Nothing
        Set .RsSY = New ADODB.Recordset
        .RsSY.ActiveConnection = FrmInfoCNTR.ConX
        .RsSY.CursorLocation = adUseClient
        .RsSY.CursorType = adOpenDynamic
        .RsSY.LockType = adLockOptimistic
        .RsSY.Open msg
        
        POSTX = 14
        Do Until .RsSY.EOF
            
            If POSTX = 14 Then
                POSTX = POSTX
            Else
                POSTX = Constant_Count
            End If
            POSTX = POSTX + 2
            Post_Count = POSTX
            Select Case UCase(.RsSY.Fields("Semester").Value)
            Case UCase("1st")
                strx = "1st Semester "
            Case UCase("2nd")
                strx = "2nd Semester "
            Case UCase("Sum")
                strx = "Summer "
            End Select

            If POSTX + 2 >= 54 Or POSTX >= 54 Then
                If Page_Count = 0 Then
                    Page_Count = 1
                    Write_Head Page_Count
                End If
                'Set Count to 14
                Post_Count = POSTX - 38
                'Write The *Continued at Page 2*
                If POSTX = 54 Then
                'Write There
                    EXW.Worksheets.Application.Cells(54, 6) = "*********Continued at Page 2********"
                Else
                    EXW.Worksheets.Application.Cells(POSTX, 6) = "*********Continued at Page 2********"
                End If
            End If
            
            If POSTX + 2 >= 122 Or POSTX >= 122 Then
                If Page_Count = 1 Then
                    Page_Count = 2
                    Write_Head Page_Count
                End If
                'Set Count to 14
                Post_Count = POSTX - 76
                'Write The *Continued at Page 3*
                If POSTX = 122 Then
                'Write There
                    EXW.Worksheets.Application.Cells(122, 6) = "*********Continued at Page 3********"
                Else
                    EXW.Worksheets.Application.Cells(POSTX, 6) = "*********Continued at Page 3********"
                End If
            End If
            
            Write_Sems .RsSY, Page_Count, strx, Post_Count
        
            sqlsubs = "select * from GRADING_SYS where Student = '" & .LBLNAME.Caption & "'" & _
            " and semester = '" & .RsSY.Fields("Semester").Value & "' and SchoolYear = '" & .RsSY.Fields("SCHOOLYEAR").Value & _
            "' and School = '" & .RsSY.Fields("SCHooL").Value & "'"
            Set .RsGrades = Nothing
            Set .RsGrades = New ADODB.Recordset
            .RsGrades.ActiveConnection = FrmInfoCNTR.ConX
            .RsGrades.CursorLocation = adUseClient
            .RsGrades.CursorType = adOpenDynamic
            .RsGrades.LockType = adLockOptimistic
            .RsGrades.Open sqlsubs
            Do Until .RsGrades.EOF
                POSTX = Constant_Count
                POSTX = POSTX + 1
                Post_Count = POSTX
                If POSTX >= 54 Then
                    If Page_Count = 0 Then
                    Page_Count = 1
                    Write_Head Page_Count
                    'Set Count to 14
                    End If
                    Post_Count = POSTX - 38
                    If POSTX = 54 Then
                        'Write There
                        EXW.Worksheets.Application.Cells(54, 7) = "*********Continued at Page 2********"
                    Else
                        EXW.Worksheets.Application.Cells(POSTX, 7) = "*********Continued at Page 2********"
                    End If
                    
                End If
            
                If POSTX >= 122 Then
                    If Page_Count = 1 Then
                    Page_Count = 2
                    Write_Head Page_Count
                    'Set Count to 14
                    End If
                    Post_Count = POSTX - 78
                    If POSTX = 54 Then
                        'Write There
                        EXW.Worksheets.Application.Cells(122, 7) = "*********Continued at Page 2********"
                    Else
                        EXW.Worksheets.Application.Cells(POSTX, 7) = "*********Continued at Page 2********"
                    End If
                    
                End If
                Write_Subject .RsGrades, Page_Count, Post_Count
                .RsGrades.MoveNext
            Loop
            .RsSY.MoveNext
        Loop
        
    End With
    
    EXW.Close True
    Set EXW = Nothing
    Set EX = Nothing
End Sub


Sub Write_Sems(rs As ADODB.Recordset, Page As Integer, strx As String, POSTX As Integer)
Dim CONS As Integer
CONS = 68
EXW.Worksheets.Application.Cells(POSTX + (CONS * Page), 7) = rs.Fields("SCHOOL").Value
EXW.Worksheets.Application.Cells(POSTX + (CONS * Page), 3) = rs.Fields("SCHOOLYEAR").Value
EXW.Worksheets.Application.Cells(POSTX + 1 + (CONS * Page), 3) = strx & rs.Fields("COURSE").Value
Constant_Count = POSTX + 1 + (CONS * Page)
End Sub


Sub Write_Subject(rs As ADODB.Recordset, Page As Integer, POSTX As Integer)
Dim CONS As Integer
CONS = 68
EXW.Worksheets.Application.Cells(POSTX + (CONS * Page), 3) = rs.Fields("SUBJECT").Value
EXW.Worksheets.Application.Cells(POSTX + (CONS * Page), 7) = rs.Fields("SUBJECT_DESCRIPTION").Value
EXW.Worksheets.Application.Cells(POSTX + (CONS * Page), 13) = rs.Fields("REEXAM").Value
EXW.Worksheets.Application.Cells(POSTX + (CONS * Page), 14) = rs.Fields("REMARKS").Value
EXW.Worksheets.Application.Cells(POSTX + (CONS * Page), 15) = rs.Fields("UNITS").Value
Constant_Count = POSTX + (CONS * Page)
End Sub
