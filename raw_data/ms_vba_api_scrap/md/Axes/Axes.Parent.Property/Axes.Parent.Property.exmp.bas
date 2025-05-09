Sub DisplayParentName() 
 
 Set myAxis = Charts(1).Axes(xlValue) 
 MsgBox myAxis.Parent.Name 
 
End Sub