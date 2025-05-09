# Application.OnRepeat method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Application.OnRepeat "Repeat VB Procedure", _ 
 "Book1.xls!My_Repeat_Sub" 
Application.OnUndo "Undo VB Procedure", _ 
 "Book1.xls!My_Undo_Sub"
```

## Parameters
- **Text**: Required
- **Procedure**: Required

## Remarks
If a procedure doesn't use the OnRepeat method, the Repeat command repeats the procedure that was run most recently.

## Example
No VBA example available.
