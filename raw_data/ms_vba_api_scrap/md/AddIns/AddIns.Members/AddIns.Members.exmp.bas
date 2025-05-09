Sub DisplayAddIns() 
 Worksheets("Sheet1").Activate 
 rw = 1 
 For Each ad In Application.AddIns 
 Worksheets("Sheet1").Cells(rw, 1) = ad.Name 
 Worksheets("Sheet1").Cells(rw, 2) = ad.Installed 
 rw = rw + 1 
 Next 
End Sub