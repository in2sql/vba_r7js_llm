Sub ChangeCursor() 
 
 Application.Cursor = xlIBeam 
 For x = 1 To 1000 
 For y = 1 to 1000 
 Next y 
 Next x 
 Application.Cursor = xlDefault 
 
End Sub