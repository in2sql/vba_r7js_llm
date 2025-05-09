# Worksheet Cells Property

## Business Description
Returns a Range object that represents all the cells on the worksheet (not just the cells that are currently in use).

## Behavior
Returns aRangeobject that represents all the cells on the worksheet (not just the cells that are currently in use).

## Example Usage
```vba
Option Explicit
Public blnToggle As Boolean

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim LastColumn As Long, keyColumn As Long, LastRow As Long
    Dim SortRange As Range
    LastColumn = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    keyColumn = Target.Column
    
    If keyColumn <= LastColumn Then
    
        Application.ScreenUpdating = False
        Cancel = True
        LastRow = Cells(Rows.Count, keyColumn).End(xlUp).Row
        Set SortRange = Target.CurrentRegion
        
        blnToggle = Not blnToggle
        If blnToggle = True Then
            SortRange.Sort Key1:=Cells(2, keyColumn), Order1:=xlAscending, Header:=xlYes
        Else
            SortRange.Sort Key1:=Cells(2, keyColumn), Order1:=xlDescending, Header:=xlYes
        End If
    
        Set SortRange = Nothing
        Application.ScreenUpdating = True
        
    End If
End Sub
```