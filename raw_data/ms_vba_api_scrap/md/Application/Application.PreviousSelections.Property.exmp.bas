On Error GoTo noSelections 
For i = LBound(Application.PreviousSelections) To _ 
 UBound(Application.PreviousSelections) 
 MsgBox Application.PreviousSelections(i).Address 
Next i 
Exit Sub 
On Error GoTo 0 
 
noSelections: 
 MsgBox "There are no previous selections"