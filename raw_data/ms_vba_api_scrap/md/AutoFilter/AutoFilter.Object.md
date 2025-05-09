# AutoFilter Object

## Business Description
Represents autofiltering for the specified worksheet.

## Behavior
Represents autofiltering for the specified worksheet.

## Example Usage
```vba
Dim w As Worksheet 
Dim filterArray() 
Dim currentFiltRange As String 
 
Sub ChangeFilters() 
 
Set w = Worksheets("Crew") 
With w.AutoFiltercurrentFiltRange = .Range.Address 
 With .Filters 
 ReDim filterArray(1 To .Count, 1 To 3) 
 For f = 1 To .Count 
 With .Item(f) 
 If .On Then 
 filterArray(f, 1) = .Criteria1 
 If .Operator Then 
 filterArray(f, 2) = .Operator 
 filterArray(f, 3) = .Criteria2 
 End If 
 End If 
 End With 
 Next 
 End With 
End With 
 
w.AutoFilterMode = False 
w.Range("A1").AutoFilter field:=1, Criteria1:="S" 
 
End Sub
```