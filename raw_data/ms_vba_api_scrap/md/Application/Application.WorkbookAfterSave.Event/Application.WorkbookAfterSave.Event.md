# Application.WorkbookAfterSave event (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, _ 
 ByVal Success As Boolean) 
If Success Then 
 MsgBox ("The " & Wb.Name & " workbook was successfully saved.") 
End If 
End Sub
```

## Parameters
- **Wb**: Required
- **Success**: Required

## Return Value
Nothing

## Remarks
For information about how to use event procedures with the Application object, see Using events with the Application object.

## Example
```vba
Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, _ 
 ByVal Success As Boolean) 
If Success Then 
 MsgBox ("The " & Wb.Name & " workbook was successfully saved.") 
End If 
End Sub
```

