# Application.RegisteredFunctions property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
theArray = Application.RegisteredFunctions 
If IsNull(theArray) Then 
 MsgBox "No registered functions" 
Else 
 For i = LBound(theArray) To UBound(theArray) 
 For j = 1 To 3 
 Worksheets("Sheet1").Cells(i, j). _ 
 Formula = theArray(i, j) 
 Next j 
 Next i 
End If
```

## Parameters
- **Index1**: Optional
- **Index2**: Optional

## Remarks
If you don't specify the index arguments, this property returns an array that contains a list of all registered functions. Each row in the array contains information about a single function, as shown in the following table.

## Example
No VBA example available.
