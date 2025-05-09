# How to: Prevent Duplicate Entries in a Range

## Business Description
The following code example verifies that a value entered in the range A1:B20 exists within that range on any of the worksheets in the current workbook and prevents duplicate entries if the value exists.

## Behavior
The following code example verifies that a value entered in the range A1:B20 exists within that range on any of the worksheets in the current workbook and prevents duplicate entries if the value exists.

## Example Usage
```vba
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    'Define your variables.
    Dim ws As Worksheet, EvalRange As Range
    
    'Set the range where you want to prevent duplicate entries.
    Set EvalRange = Range("A1:B20")
    
    'If the cell where value was entered is not in the defined range, if the value pasted is larger than a single cell,
    'or if no value was entered in the cell, then exit the macro.
    If Intersect(Target, EvalRange) Is Nothing Or Target.Cells.Count > 1 Then Exit Sub
    If IsEmpty(Target) Then Exit Sub
    
    'If the value entered already exists in the defined range on the current worksheet, throw an
    'error message and undo the entry.
    If WorksheetFunction.CountIf(EvalRange, Target.Value) > 1 Then
        MsgBox Target.Value & " already exists on this sheet."
        Application.EnableEvents = False
        Application.Undo
        Application.EnableEvents = True
    End If
    
    'Check the other worksheets in the workbook.
    For Each ws In Worksheets
        With ws
            If .Name <> Target.Parent.Name Then
                'If the value entered already exists in the defined range on the current worksheet, throw an
                'error message and undo the entry.
                If WorksheetFunction.CountIf(Sheets(.Name).Range("A1:B20"), Target.Value) > 0 Then
                    MsgBox Target.Value & " already exists on the sheet named " & .Name & ".", _
                    16, "No duplicates allowed in " & EvalRange.Address(0, 0) & "."
                    Application.EnableEvents = False
                    Application.Undo
                    Application.EnableEvents = True
                    Exit For
                End If
            End If
        End With
    Next ws

End Sub
```