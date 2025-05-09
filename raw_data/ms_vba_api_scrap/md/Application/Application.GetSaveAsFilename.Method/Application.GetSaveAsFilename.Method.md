# Application.GetSaveAsFilename method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
fileSaveName = Application.GetSaveAsFilename( _ 
 fileFilter:="Text Files (*.txt), *.txt") 
If fileSaveName <> False Then 
 MsgBox "Save as " & fileSaveName 
End If
```

## Parameters
- **InitialFilename**: Optional
- **FileFilter**: Optional
- **FilterIndex**: Optional
- **Title**: Optional
- **ButtonText**: Optional

## Return Value
Variant

## Remarks
This string passed in the FileFilter argument consists of pairs of file filter strings followed by the MS-DOS wildcard file filter specification, with each part and each pair separated by commas. Each separate pair is listed in the Files of type drop-down list box. For example, the following string specifies two file filtersâtext and addin:

## Example
No VBA example available.
