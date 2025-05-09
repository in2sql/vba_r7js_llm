# Speech SpeakCellOnEnter Property

## Business Description
Microsoft Excel supports a mode where the active cell will be spoken when the ENTER key is pressed or when the active cell is finished being edited. Setting the SpeakCellOnEnter property to True will turn this mode on. False turns this mode off.

## Behavior
Microsoft Excel supports a mode where the active cell will be spoken when the ENTER key is pressed or when the active cell is finished being edited.  Setting theSpeakCellOnEnterproperty toTruewill turn this mode on.Falseturns this mode off. Read/writeBoolean.

## Example Usage
```vba
Sub SpeechCheck() 
 
 ' Determine mode setting and notify user. 
 If Application.Speech.SpeakCellOnEnter= True Then 
 MsgBox "The Speak On Enter mode is turned on. " & _ 
 "The active cell will be spoken when the ENTER "& _ 
 "key is pressed or it is done being edited." 
 Else 
 MsgBox "The Speaker On Enter mode is turned off." 
 End If 
 
End Sub
```