# Application.FileConverters property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
installedCvts = Application.FileConverters 
foundMultiplan = False 
If Not IsNull(installedCvts) Then 
 For arrayRow = 1 To UBound(installedCvts, 1) 
 If installedCvts(arrayRow, 1) Like "*Multiplan*" Then 
 foundMultiplan = True 
 Exit For 
 End If 
 Next arrayRow 
End If 
If foundMultiplan = True Then 
 MsgBox "Multiplan converter is installed" 
Else 
 MsgBox "Multiplan converter is not installed" 
End If
```

## Parameters
- **Index1**: Optional
- **Index2**: Optional

## Remarks
If you don't specify the index arguments, this property returns an array that contains information about all the installed file converters. Each row in the array contains information about a single file converter, as shown in the following table.

## Example
No VBA example available.
