Set newSheet = Sheets.Add(Type:=xlWorksheet) 
For i = 1 To Sheets.Count 
 newSheet.Cells(i, 1).Value = Sheets(i).Name 
Next i