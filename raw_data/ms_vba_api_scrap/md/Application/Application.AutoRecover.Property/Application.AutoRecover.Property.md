# Application.AutoRecover property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub UseAutoRecover() 
 
 Application.AutoRecover.Time = 5 
 
 MsgBox "The time that will elapse between each automatic " & _ 
 "save has been set to " & _ 
 Application.AutoRecover.Time & " minutes." 
 
End Sub
```

## Remarks
Valid time intervals are whole numbers from 1 to 120.

## Example
```vba
Sub UseAutoRecover() 
 
 Application.AutoRecover.Time = 5 
 
 MsgBox "The time that will elapse between each automatic " & _ 
 "save has been set to " & _ 
 Application.AutoRecover.Time & " minutes." 
 
End Sub
```

