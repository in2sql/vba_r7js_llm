# AddIns.Add method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub UseAddIn() 
 
 Set myAddIn = AddIns.Add(Filename:="A:\MYADDIN.XLA", _ 
 CopyFile:=True) 
 MsgBox myAddIn.Title & " has been added to the list" 
 
End Sub
```

## Parameters
- **FileName**: Required
- **CopyFile**: Optional

## Return Value
An AddIn object that represents the new add-in.

## Remarks
This method does not install the new add-in. You must set the Installed property to install the add-in.

## Example
```vba
Sub UseAddIn() 
 
 Set myAddIn = AddIns.Add(Filename:="A:\MYADDIN.XLA", _ 
 CopyFile:=True) 
 MsgBox myAddIn.Title & " has been added to the list" 
 
End Sub
```

