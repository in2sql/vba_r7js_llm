# AddIn object (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Remarks
The AddIn object is a member of the AddIns collection. The AddIns collection contains a list of all the add-ins available to Microsoft Excel, regardless of whether they're installed. This list corresponds to the list of add-ins displayed in the Add-Ins dialog box.

## Example
```vba
Set wb = Workbooks("myaddin.xla")
```

```vba
Set wb = Workbooks(AddIns("analysis toolpak").Name)
```

```vba
On Error Resume Next ' turn off error checking 
Set wbMyAddin = Workbooks(AddIns("My Addin").Name) 
lastError = Err 
On Error Goto 0 ' restore error checking 
If lastError <> 0 Then 
 ' the add-in workbook isn't currently open. Manually open it. 
 Set wbMyAddin = Workbooks.Open(AddIns("My Addin").FullName) 
End If
```

