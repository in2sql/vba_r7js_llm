Worksheets("Sheet1").Activate 
Set isect = Application.Intersect(Range("rg1"), Range("rg2")) 
If isect Is Nothing Then 
 MsgBox "Ranges don't intersect" 
Else 
 isect.Select 
End If