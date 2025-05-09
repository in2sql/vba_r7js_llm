# Application.Intersect method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Worksheets("Sheet1").Activate 
Set isect = Application.Intersect(Range("rg1"), Range("rg2")) 
If isect Is Nothing Then 
 MsgBox "Ranges don't intersect" 
Else 
 isect.Select 
End If
```

## Parameters
- **Arg1**: Required
- **Arg2**: Required
- **Arg3âArg30**: Optional

## Return Value
Range

## Example
```vba
Worksheets("Sheet1").Activate 
Set isect = Application.Intersect(Range("rg1"), Range("rg2")) 
If isect Is Nothing Then 
 MsgBox "Ranges don't intersect" 
Else 
 isect.Select 
End If
```

