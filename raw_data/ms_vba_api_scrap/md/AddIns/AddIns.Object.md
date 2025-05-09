# AddIns Object

## Business Description
A collection of AddIn objects that represents all the add-ins available to Microsoft Excel, regardless of whether they're installed.

## Behavior
A collection ofAddInobjects that represents all the add-ins available to Microsoft Excel, regardless of whether they're installed.

## Example Usage
```vba
Sub DisplayAddIns() 
 Worksheets("Sheet1").Activate 
 rw = 1 
 For Each ad In Application.AddInsWorksheets("Sheet1").Cells(rw, 1) = ad.Name 
 Worksheets("Sheet1").Cells(rw, 2) = ad.Installed 
 rw = rw + 1 
 Next 
End Sub
```