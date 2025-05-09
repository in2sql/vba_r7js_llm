# Application.OnUndo method (Excel)

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
If a procedure doesn't use the OnUndo method, the Undo command is disabled.

## Example
No VBA example available.
