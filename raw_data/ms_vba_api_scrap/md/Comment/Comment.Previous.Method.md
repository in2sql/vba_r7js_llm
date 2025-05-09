# Comment Previous Method

## Business Description
Returns a Comment object that represents the previous comment.

## Behavior
Returns aCommentobject that represents the previous comment.

## Example Usage
```vba
'Sets up the comments 
For xNum = 1 To 10 
 Range("A" & xNum).AddComment 
 Range("A" & xNum).Comment.Text Text:="Comment " & xNum 
Next 
 
MsgBox "Comments created... A1:A10" 
 
'Deletes every second comment in the A1:A10 range 
For yNum = 10 To 1 Step -2 
 Range("A" & yNum).Comment.Previous.Shape.Select True 
 Selection.Delete 
Next 
 
MsgBox "Deleted every second comment"
```