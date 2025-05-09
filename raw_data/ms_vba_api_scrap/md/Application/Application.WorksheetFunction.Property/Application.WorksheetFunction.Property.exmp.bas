Set myRange = Worksheets("Sheet1").Range("A1:C10") 
answer = Application.WorksheetFunction.Min(myRange) 
MsgBox answer