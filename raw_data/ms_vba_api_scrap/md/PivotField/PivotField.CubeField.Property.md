# PivotField CubeField Property

## Business Description
Returns the CubeField object from which the specified PivotTable field is descended. Read-only.

## Behavior
Returns theCubeFieldobject from which the specified PivotTable field is descended. Read-only.

## Example Usage
```vba
Sub UseCubeField() 
 
 Dim objNewSheet As Worksheet 
 Set objNewSheet = Worksheets.Add 
 objNewSheet.Activate 
 intRow = 1 
 
 For Each objPF in _ 
 Worksheets(1).PivotTables(1).PivotFields 
 If objPF.CubeField.CubeFieldType = xlHierarchy Then 
 objNewSheet.Cells(intRow, 1).Value = objPF.Name 
 intRow = intRow + 1 
 End If 
 Next objPF 
 
End Sub
```