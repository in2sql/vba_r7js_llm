# AddIn.progID property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
rw = 0 
For Each o in Worksheets(1).OLEObjects 
 With Worksheets(2) 
 rw = rw + 1 
 .cells(rw, 1).Value = o.ProgId 
 End With 
Next
```

## Example
No VBA example available.
