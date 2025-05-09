Attribute VB_Name = "modConnect"
Option Explicit

Public VBInstance As VBIDE.VBE
Public Sub CloseAll()
  Dim X As Long
    
    X = VBInstance.ActiveVBProject.VBE.CodePanes.Count
    If X > 0 Then
        Do
            VBInstance.ActiveVBProject.VBE.CodePanes(1).Window.Close
            X = VBInstance.ActiveVBProject.VBE.CodePanes.Count
        Loop Until X = 0
    End If
    
End Sub

