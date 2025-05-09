Worksheets("Sheet1").Activate 
SolverReset 
SolverOptions precision:=0.001 
SolverOK setCell:=Range("TotalProfit"), _ 
 maxMinVal:=1, _ 
 byChange:=Range("C4:E6") 
SolverAdd cellRef:=Range("F4:F6"), _ 
 relation:=1, _ 
 formulaText:=100 
SolverAdd cellRef:=Range("C4:E6"), _ 
 relation:=3, _ 
 formulaText:=0 
SolverAdd cellRef:=Range("C4:E6"), _ 
 relation:=4 
SolverSolve userFinish:=False 
SolverSave saveArea:=Range("A33")
