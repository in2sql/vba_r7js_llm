Attribute VB_Name = "modViewMaximise"
Option Explicit

Private Const SM_XVIRTUALSCREEN = 76
Private Const SM_YVIRTUALSCREEN = 77
Private Const SM_CXVIRTUALSCREEN = 78
Private Const SM_CYVIRTUALSCREEN = 79
Private Const SM_CMONITORS = 80

Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Sub FillVirtualScreen()
    Static Win_W As Long
    Static Win_H As Long
    Static Win_L As Long
    Static Win_T As Long
    Static Win_State As XlWindowState
    Static Filled As Boolean
    
    Dim monitors As Long
    Dim w As Long
    Dim H As Long
    
    monitors = GetSystemMetrics(SM_CMONITORS)
    If monitors = 1 Then
        MsgBox "This option requires 2 or more monitors."
        Exit Sub
    End If
    
    w = GetSystemMetrics(SM_CXVIRTUALSCREEN)
    H = GetSystemMetrics(SM_CYVIRTUALSCREEN)
    
    With Application
        '@ check if possible with primary monitor on the right
        ' if not, inform user
        
        If Filled Then
            ' Restore to initial position
            .Left = Win_L
            .Top = Win_T
            .Width = Win_W
            .Height = Win_H
            .WindowState = Win_State
            
            Filled = False
        Else
            ' Save initial state
            ' Preserve proprer dimensions by
            ' saving them in normal state
            Win_State = .WindowState
            .WindowState = xlNormal
            
            Win_L = .Left
            Win_T = .Top
            Win_W = .Width
            Win_H = .Height
            
            
            ' Maximize across screens
            .WindowState = xlNormal
            .Left = 0
            .Top = 0
            .Width = w
            .Height = H
            
            Filled = True
        End If
    End With
End Sub
