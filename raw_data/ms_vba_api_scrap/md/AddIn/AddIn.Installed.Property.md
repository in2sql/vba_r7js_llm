# AddIn.Installed property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Set a = AddIns("Solver Add-In") 
If a.Installed = True Then 
 MsgBox "The Solver add-in is installed" 
Else 
 MsgBox "The Solver add-in is not installed" 
End If
```

## Remarks
Setting this property to True installs the add-in and calls its Auto_Add functions. Setting this property to False removes the add-in and calls its Auto_Remove functions.

## Example
```vba
Set a = AddIns("Solver Add-In") 
If a.Installed = True Then 
 MsgBox "The Solver add-in is installed" 
Else 
 MsgBox "The Solver add-in is not installed" 
End If
```

