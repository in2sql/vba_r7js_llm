# FileExportConverter FileFormat Property

## Business Description
Returns an integer that identifies the file format associated with the specified FileExportConverter object. Read-only.

## Behavior
Returns an integer that identifies the file format associated with the specifiedFileExportConverterobject. Read-only.

## Example Usage
```vba
ActiveWorkbook.SaveAs _ 
 Filename:="C:\temp\myFile.xyz", _ 
 FileFormat:=Application.FileExportConverters(1).FileFormat, _ 
 CreateBackup:=False
```