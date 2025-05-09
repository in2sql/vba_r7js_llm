[a1].Value = 25 
Evaluate("A1").Value = 25 
 
trigVariable = [SIN(45)] 
trigVariable = Evaluate("SIN(45)") 
 
Set firstCellInSheet = Workbooks("BOOK1.XLS").Sheets(4).[A1] 
Set firstCellInSheet = _ 
    Workbooks("BOOK1.XLS").Sheets(4).Evaluate("A1")