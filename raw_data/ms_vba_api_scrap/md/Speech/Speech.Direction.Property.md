# Speech Direction Property

## Business Description
Returns or sets the order in which the cells will be spoken. The value of the Direction property is an XlSpeakDirection constant. Read/write.

## Behavior
Returns or sets the order in which the cells will be spoken.  The value of theDirectionproperty is anXlSpeakDirectionconstant. Read/write.

## Example Usage
```vba
Sub CheckSpeechDirection() 
 
 ' Notify user of speech direction. 
 If Application.Speech.Direction= xlSpeakByColumns Then 
 MsgBox "The speech direction is set to speak by columns." 
 Else 
 MsgBox "The speech direction is set to speak by rows." 
 End If 
 
End Sub
```