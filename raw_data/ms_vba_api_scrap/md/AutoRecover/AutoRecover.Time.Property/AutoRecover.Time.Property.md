# AutoRecover.Time property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub SetTimeValue() 
 
 Application.AutoRecover.Time = 5 
 MsgBox "The AutoRecover time interval is set at " & _ 
 Application.AutoRecover.Time & " minutes." 
 
End Sub
```

## Remarks
Entering a decimal value will round to the nearest whole number. For example, entering a value of 5.5 is the equivalent of 6.

## Example
```vba
Sub SetTimeValue() 
 
 Application.AutoRecover.Time = 5 
 MsgBox "The AutoRecover time interval is set at " & _ 
 Application.AutoRecover.Time & " minutes." 
 
End Sub
```

