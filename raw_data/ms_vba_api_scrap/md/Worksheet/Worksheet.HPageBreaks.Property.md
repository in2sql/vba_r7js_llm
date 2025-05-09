# Worksheet HPageBreaks Property

## Business Description
Returns an HPageBreaks collection that represents the horizontal page breaks on the sheet. Read-only.

## Behavior
Returns anHPageBreakscollection that represents the horizontal page breaks on the sheet. Read-only.

## Example Usage
```vba
Sub AddPageBreaks() 
    StartRow = 2 
    FinalRow = Range("A65536").End(xlUp).Row 
    LastVal = Cells(StartRow, 1).Value 
    For i = StartRow To FinalRow 
    ThisVal = Cells(i, 1).Value 
    If Not ThisVal = LastVal Then 
    ActiveSheet.HPageBreaks.Add before:=Cells(i, 1) 
    End If 
    LastVal = ThisVal 
    Next i 
End Sub
```