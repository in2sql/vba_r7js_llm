# How to: Name a Worksheet By Using a Cell Value

## Business Description
This example shows how to name a worksheet by using the value in cell A1 on that sheet.

## Behavior
This example shows how to name a worksheet by using the value in cell A1 on that sheet. This example verifies that the value in cell A1 is a valid worksheet name, and if it is a valid name, renames the active worksheet to equal the value of cell A1 by using theNameproperty of theWorksheetobject.

## Example Usage
```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    'Specify the target cell whose entry shall be the sheet tab name.
    If Target.Address <> "$A$1" Then Exit Sub
        'If the target cell is empty (contents cleared) then do not change the shet name
    If IsEmpty(Target) Then Exit Sub

    'If the length of the target cell's entry is greater than 31 characters, disallow the entry.
    If Len(Target.Value) > 31 Then
        MsgBox "Worksheet tab names cannot be greater than 31 characters in length." & vbCrLf & _
        "You entered " & Target.Value & ", which has " & Len(Target.Value) & " characters.", , "Keep it under 31 characters"
        Application.EnableEvents = False
        Target.ClearContents
        Application.EnableEvents = True
        Exit Sub
    End If

    'Sheet tab names cannot contain the characters /, \, [, ], *, ?, or :.
    'Verify that none of these characters are present in the cell's entry.
    Dim IllegalCharacter(1 To 7) As String, i As Integer
    IllegalCharacter(1) = "/"
    IllegalCharacter(2) = "\"
    IllegalCharacter(3) = "["
    IllegalCharacter(4) = "]"
    IllegalCharacter(5) = "*"
    IllegalCharacter(6) = "?"
    IllegalCharacter(7) = ":"
    For i = 1 To 7
        If InStr(Target.Value, (IllegalCharacter(i))) > 0 Then
            MsgBox "You used a character that violates sheet naming rules." & vbCrLf & vbCrLf & _
            "Please re-enter a sheet name without the ''" & IllegalCharacter(i) & "'' character.", 48, "Not a possible sheet name !!"
            Application.EnableEvents = False
            Target.ClearContents
            Application.EnableEvents = True
            Exit Sub
        End If
    Next i

    'Verify that the proposed sheet name does not already exist in the workbook.
    Dim strSheetName As String, wks As Worksheet, bln As Boolean
    strSheetName = Trim(Target.Value)
    On Error Resume Next
    Set wks = ActiveWorkbook.Worksheets(strSheetName)
    On Error Resume Next
    If Not wks Is Nothing Then
        bln = True
    Else
        bln = False
        Err.Clear
    End If

    'If the worksheet name does not already exist, name the active sheet as the target cell value.
    'Otherwise, advise the user that duplicate sheet names are not allowed.
    If bln = False Then
        ActiveSheet.Name = strSheetName
    Else
        MsgBox "There is already a sheet named " & strSheetName & "." & vbCrLf & _
        "Please enter a unique name for this sheet."
        Application.EnableEvents = False
        Target.ClearContents
        Application.EnableEvents = True
    End If

End Sub
```