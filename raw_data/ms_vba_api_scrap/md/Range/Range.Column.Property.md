# Range Column Property

## Business Description
Returns the number of the first column in the first area in the specified range. Read-only Long.

## Behavior
Returns the number of the first column in the first area in the specified range. Read-onlyLong.

## Example Usage
```vba
Sub Delete_Empty_Columns()
    'The range from which to delete the columns.
    Dim rnSelection As Range
    
    'Column and count variables used in the deletion process.
    Dim lnLastColumn As Long
    Dim lnColumnCount As Long
    Dim lnDeletedColumns As Long
    
    lnDeletedColumns = 0
    
    'Confirm that a range is selected, and that the range is contiguous.
    If TypeName(Selection) = "Range" Then
        If Selection.Areas.Count = 1 Then
            
            'Initialize the range to what the user has selected, and initialize the count for the upcoming FOR loop.
            Set rnSelection = Application.Selection
            lnLastColumn = rnSelection.Columns.Count
        
            'Start at the far-right column and work left: if the column is empty then
            'delete the column and increment the deleted column count.
            For lnColumnCount = lnLastColumn To 1 Step -1
                If Application.CountA(rnSelection.Columns(lnColumnCount)) = 0 Then
                    rnSelection.Columns(lnColumnCount).Delete
                    lnDeletedColumns = lnDeletedColumns + 1
                End If
            Next lnColumnCount
    
            rnSelection.Resize(lnLastColumn - lnDeletedColumns).Select
        Else
            MsgBox "Please select only one area.", vbInformation
        End If
    Else
        MsgBox "Please select a range.", vbInformation
    End If
    
    'Turn screen updating back on.
    Application.ScreenUpdating = True

End Sub
```