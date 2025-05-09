# How to: Automatically Dismiss a Message Box

## Business Description
This example shows how to automatically dismiss a message box after a specified period of time. This example displays a message box and then automatically dismisses it after 10 seconds.

## Behavior
This example shows how to automatically dismiss a message box after a specified period of time. This example displays a message box and then automatically dismisses it after 10 seconds.

## Example Usage
```vba
Sub MessageBoxTimer()
    Dim AckTime As Integer, InfoBox As Object
    Set InfoBox = CreateObject("WScript.Shell")
    'Set the message box to close after 10 seconds
    AckTime = 10
    Select Case InfoBox.Popup("Click OK (this window closes automatically after 10 seconds).", _
    AckTime, "This is your Message Box", 0)
        Case 1, -1
            Exit Sub
    End Select
End Sub
```