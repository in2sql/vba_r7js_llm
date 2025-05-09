# Application.WorkbookPivotTableCloseConnection event (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Private Sub ConnectionApp_WorkbookPivotTableCloseConnection(ByVal wbOne As Workbook, Target As PivotTable) 
 
 MsgBox "The PivotTable connection has been closed." 
 
End Sub
```

## Parameters
- **Wb**: Required
- **Target**: Required

## Return Value
Nothing

## Example
```vba
Private Sub ConnectionApp_WorkbookPivotTableCloseConnection(ByVal wbOne As Workbook, Target As PivotTable) 
 
 MsgBox "The PivotTable connection has been closed." 
 
End Sub
```

