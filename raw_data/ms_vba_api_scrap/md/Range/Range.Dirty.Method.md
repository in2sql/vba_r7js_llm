# Range Dirty Method

## Business Description
Designates a range to be recalculated when the next recalculation occurs.

## Behavior
Designates a range to be recalculated when the next recalculation occurs.

## Example Usage
```vba
Sub UseDirtyMethod() 
 
 MsgBox "Two values and a formula will be entered." 
 Range("A1").Value = 1 
 Range("A2").Value = 2 
 Range("A3").Formula = "=A1+A2" 
 
 ' Save the changes made to the worksheet. 
 Application.DisplayAlerts = False 
 Application.Save 
 MsgBox "Changes saved." 
 
 ' Force a recalculation of range A3. 
 Application.Range("A3").DirtyMsgBox "Try to close the file without saving and a dialog box will appear." 
 
End Sub
```