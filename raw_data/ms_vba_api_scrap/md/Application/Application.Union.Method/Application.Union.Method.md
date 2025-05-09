# Application.Union method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Worksheets("Sheet1").Activate 
Set bigRange = Application.Union(Range("Range1"), Range("Range2")) 
bigRange.Formula = "=RAND()"
```

## Parameters
- **Arg1**: Required
- **Arg2**: Required
- **Arg3 â Arg30**: Optional

## Return Value
Range

## Example
```vba
Worksheets("Sheet1").Activate 
Set bigRange = Application.Union(Range("Range1"), Range("Range2")) 
bigRange.Formula = "=RAND()"
```

