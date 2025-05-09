# AddIn Object

## Business Description
Represents a single add-in, either installed or not installed.

## Behavior
Represents a single add-in, either installed or not installed.

## Example Usage
```vba
With Worksheets("sheet1") 
 .Rows(1).Font.Bold = True 
 .Range("a1:d1").Value = _ 
 Array("Name", "Full Name", "Title", "Installed") 
 For i = 1 ToAddIns.Count 
 .Cells(i + 1, 1) = AddIns(i).Name 
 .Cells(i + 1, 2) = AddIns(i).FullName 
 .Cells(i + 1, 3) = AddIns(i).Title 
 .Cells(i + 1, 4) = AddIns(i).Installed 
 Next 
 .Range("a1").CurrentRegion.Columns.AutoFit 
End With
```