# Application.FileDialog property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub UseFileDialogOpen() 
 
    Dim lngCount As Long 
 
    ' Open the file dialog 
    With Application.FileDialog(msoFileDialogOpen) 
        .AllowMultiSelect = True 
        .Show 
 
        ' Display paths of each file selected 
        For lngCount = 1 To .SelectedItems.Count 
            MsgBox .SelectedItems(lngCount) 
        Next lngCount 
 
    End With 
 
End Sub
```

## Parameters
- **fileDialogType**: Required

## Remarks
MsoFileDialogType can be one of these constants:

## Example
```vba
Sub UseFileDialogOpen() 
 
    Dim lngCount As Long 
 
    ' Open the file dialog 
    With Application.FileDialog(msoFileDialogOpen) 
        .AllowMultiSelect = True 
        .Show 
 
        ' Display paths of each file selected 
        For lngCount = 1 To .SelectedItems.Count 
            MsgBox .SelectedItems(lngCount) 
        Next lngCount 
 
    End With 
 
End Sub
```

