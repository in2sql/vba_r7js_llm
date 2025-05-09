# AddIns object (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Remarks
This list corresponds to the list of add-ins displayed in the Add-Ins dialog box.

## Example
```vba
Sub DisplayAddIns() 
 Worksheets("Sheet1").Activate 
 rw = 1 
 For Each ad In Application.AddIns 
 Worksheets("Sheet1").Cells(rw, 1) = ad.Name 
 Worksheets("Sheet1").Cells(rw, 2) = ad.Installed 
 rw = rw + 1 
 Next 
End Sub
```

