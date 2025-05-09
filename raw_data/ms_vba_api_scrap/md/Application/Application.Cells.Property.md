# Application.Cells property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub ChangeInsertRows()
    Application.ScreenUpdating = False
    Dim xRow As Long
    
    For xRow = Application.Cells(Rows.Count, 1).End(xlUp).Row To 3 Step -1
        If Application.Cells(xRow, 1).Value <> Application.Cells(xRow - 1, 1).Value Then Rows(xRow).Resize(1).Insert
    Next xRow
    
    Application.ScreenUpdating = True
End Sub
```

## Remarks
Because the Item property is the default property for the Range object, you can specify the row and column index immediately after the Cells keyword. For more information, see the Item property and the examples for this topic.

## Example
```vba
Sub ChangeInsertRows()
    Application.ScreenUpdating = False
    Dim xRow As Long
    
    For xRow = Application.Cells(Rows.Count, 1).End(xlUp).Row To 3 Step -1
        If Application.Cells(xRow, 1).Value <> Application.Cells(xRow - 1, 1).Value Then Rows(xRow).Resize(1).Insert
    Next xRow
    
    Application.ScreenUpdating = True
End Sub
```

