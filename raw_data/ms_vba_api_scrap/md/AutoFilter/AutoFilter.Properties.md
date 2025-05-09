# AutoFilter object (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Example
```vba
Dim w As Worksheet 
Dim filterArray() 
Dim currentFiltRange As String 
 
Sub ChangeFilters() 
 
Set w = Worksheets("Crew") 
With w.AutoFilter 
 currentFiltRange = .Range.Address 
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

```vba
Sub RestoreFilters() 
Set w = Worksheets("Crew") 
w.AutoFilterMode = False 
For col = 1 To UBound(filterArray(), 1) 
 If Not IsEmpty(filterArray(col, 1)) Then 
 If filterArray(col, 2) Then 
 w.Range(currentFiltRange).AutoFilter field:=col, _ 
 Criteria1:=filterArray(col, 1), _ 
 Operator:=filterArray(col, 2), _ 
 Criteria2:=filterArray(col, 3) 
 Else 
 w.Range(currentFiltRange).AutoFilter field:=col, _ 
 Criteria1:=filterArray(col, 1) 
 End If 
 End If 
Next 
End Sub
```

