# Application.ThisCell property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Function UseThisCell() 
 MsgBox "The cell address is: " & _ 
 Application.ThisCell.Address 
End Function
```

## Remarks
Users should not access properties or methods on the Range object when inside the user-defined function. Users can cache the Range object for later use and perform additional actions when the recalculation is finished.

## Example
```vba
Function UseThisCell() 
 MsgBox "The cell address is: " & _ 
 Application.ThisCell.Address 
End Function
```

