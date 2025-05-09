# AutoCorrect.ReplacementList property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
repl = Application.AutoCorrect.ReplacementList 
For x = 1 To UBound(repl) 
 If repl(x, 1) = "Temperature" Then MsgBox repl(x, 2) 
Next
```

## Parameters
- **Index**: Optional

## Remarks
If Index is not specified, this method returns a two-dimensional array. Each row in the array contains one replacement, as shown in the following table.

## Example
No VBA example available.
