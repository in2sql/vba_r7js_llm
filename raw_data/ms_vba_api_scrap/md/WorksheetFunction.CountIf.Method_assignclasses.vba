' There must be 2 sheets in the workbook, 1 being classes, and 1 being signups
' In the classes sheet, Col A is the course name, Col B is the max class size
' In the signups sheet, Col B to D is the options selected, Col E is where the assigned class will be filled
' The macro is designed to be used after signups are collected from onedrive

Sub AssignClasses()
    Dim assigned As Boolean
    Dim options As Range
    Dim i, classCount, classMax, curRow As Integer
    Dim classes, signups As Worksheet
    Dim assignCol As String
    assignCol = "E"
    Set classes = ThisWorkbook.Sheets("Classes")
    Set signups = ThisWorkbook.Sheets("Signups")
    signups.Activate
    Range(Range(assignCol & "2"), Range(assignCol & "2").End(xlDown)).Delete
    signups.Range("A2").Select
    Do Until IsEmpty(ActiveCell)
        curRow = ActiveCell.Row
        Set options = Range("B" & curRow & ":D" & curRow)
        i = 1
        assigned = False
        Do
            classCount = Application.WorksheetFunction.CountIf(Range(assignCol & ":" & assignCol), options.Columns(i))
            classMax = Application.WorksheetFunction.VLookup(options.Columns(i), classes.Range("A2:B13").Value, 2, False)
            If classCount < classMax Then
                signups.Range(assignCol & curRow).Value = options.Columns(i)
                assigned = True
            End If
            i = i + 1
        Loop While (i < 4 And Not assigned)
        If Not assigned Then
            signups.Range("E" & curRow).Value = "N.A."
        End If
        ActiveCell.Offset(1, 0).Select
    Loop
End Sub
