# Application.Caller property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Select Case TypeName(Application.Caller) 
 Case "Range" 
 v = Application.Caller.Address 
 Case "String" 
 v = Application.Caller 
 Case "Error" 
 v = "Error" 
 Case Else 
 v = "unknown" 
End Select 
MsgBox "caller = " & v
```

## Parameters
- **Index**: Optional

## Remarks
This property returns information about how Visual Basic was called, as shown in the following table.

## Example
No VBA example available.
