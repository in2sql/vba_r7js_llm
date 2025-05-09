Set rangeToUse = Selection 
If rangeToUse.Areas.Count = 1 Then 
 myOperation rangeToUse 
Else 
 For Each singleArea in rangeToUse.Areas 
 myOperation singleArea 
 Next 
End If