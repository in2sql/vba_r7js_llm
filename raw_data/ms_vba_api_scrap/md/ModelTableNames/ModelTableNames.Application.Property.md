# ModelTableNames Application Property

## Business Description
Returns an Application object that represents the Microsoft Excel application. Read-only.

## Behavior
Returns anApplicationobject that represents the Microsoft Excel application. Read-only.

## Example Usage
```vba
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If
```