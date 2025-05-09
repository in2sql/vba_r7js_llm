# Application.PreviousSelections property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
On Error GoTo noSelections 
For i = LBound(Application.PreviousSelections) To _ 
 UBound(Application.PreviousSelections) 
 MsgBox Application.PreviousSelections(i).Address 
Next i 
Exit Sub 
On Error GoTo 0 
 
noSelections: 
 MsgBox "There are no previous selections"
```

## Parameters
- **Index**: Optional

## Remarks
Each time you go to a range or cell by using the Name box or the Go To command (Edit menu), or each time a macro calls the Goto method, the previous range is added to this array as element number 1, and the other items in the array are moved down.

## Example
```vba
On Error GoTo noSelections 
For i = LBound(Application.PreviousSelections) To _ 
 UBound(Application.PreviousSelections) 
 MsgBox Application.PreviousSelections(i).Address 
Next i 
Exit Sub 
On Error GoTo 0 
 
noSelections: 
 MsgBox "There are no previous selections"
```

