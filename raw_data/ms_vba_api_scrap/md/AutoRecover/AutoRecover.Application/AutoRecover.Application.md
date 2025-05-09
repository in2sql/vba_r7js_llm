# AutoRecover object (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Remarks
Properties for the AutoRecover object determine the path and time interval for backing up all files.

## Example
```vba
Sub SetPath() 
 
 Application.AutoRecover.Path = "C:\" 
 
End Sub
```

```vba
Sub SetTime() 
 
 Application.AutoRecover.Time = 5 
 
End Sub
```

