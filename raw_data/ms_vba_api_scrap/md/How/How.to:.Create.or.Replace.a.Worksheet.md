# How to: Create or Replace a Worksheet

## Business Description
The following examples show how to determine if a worksheet exists, and then how to create or replace the worksheet.

## Behavior
The following examples show how to determine if a worksheet exists, and then how to create or replace the worksheet.

## Example Usage
```vba
Sub TestSheetYesNo()
    Dim mySheetName As String, mySheetNameTest As String
    mySheetName = "Sheet4"
    
    On Error Resume Next
    mySheetNameTest = Worksheets(mySheetName).Name
    If Err.Number = 0 Then
        MsgBox "The sheet named ''" & mySheetName & "'' DOES exist in this workbook."
    Else
        Err.Clear
        MsgBox "The sheet named ''" & mySheetName & "'' does NOT exist in this workbook."
    End If
End Sub
```