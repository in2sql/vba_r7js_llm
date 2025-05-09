# Application.OLEDBErrors property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Set objEr = Application.OLEDBErrors.Item(1) 
MsgBox "The following error occurred:" & _ 
 objEr.ErrorString & " : " & objEr.SqlState
```

## Example
```vba
Set objEr = Application.OLEDBErrors.Item(1) 
MsgBox "The following error occurred:" & _ 
 objEr.ErrorString & " : " & objEr.SqlState
```

