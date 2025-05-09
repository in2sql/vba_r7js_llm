VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Chess by MAN Soft"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSaveSettings 
         Caption         =   "&Save Settings"
         Shortcut        =   ^S
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuEngineThinking 
         Caption         =   "Engine Thinking"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuMoveList 
         Caption         =   "Move List"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuCaptPieces 
         Caption         =   "Captured Pieces"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuChessClock 
         Caption         =   "Chess Clock"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu setings 
         Caption         =   "Settings"
      End
      Begin VB.Menu MoveClock 
         Caption         =   "Move clock"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    Dim strVariable As String
    Dim bolValue As Boolean
    
    'Open the INI file "sah.ini" and
    'read the settings information
    'previously stored there...
    
    Open App.Path & "\sah.ini" For Input As #1
    
    Line Input #1, strVariable
    Call subProcessData(strVariable, intBoardPosX, intBoardPosY, bolValue)
    
    Line Input #1, strVariable
    Call subProcessData(strVariable, intMiniBoardPosX, intMiniBoardPosY, bolValue)
    mnuEngineThinking.Checked = bolValue
    
    Line Input #1, strVariable
    Call subProcessData(strVariable, intMoveListPosX, intMoveListPosY, bolValue)
    mnuMoveList.Checked = bolValue
    
    Line Input #1, strVariable
    Call subProcessData(strVariable, intClockPosX, intClockPosY, bolValue)
    mnuChessClock.Checked = bolValue
    
    Line Input #1, strVariable
    Call subProcessData(strVariable, intCapturPiecesPosX, intCapturPiecesPosY, bolValue)
    mnuCaptPieces.Checked = bolValue
    
    Line Input #1, strVariable
    frmSettings.srbGameLevel.Value = CInt(Mid(strVariable, 17))
    
    Line Input #1, strVariable
    frmSettings.srbPlayingStyle.Value = CInt(Mid(strVariable, 12))
    
    Close
    
    'Setup the chess boards to
    'begin the game...
    Call subResetChessBoard
    Call subResetMiniChessBoard
    
    'Open the various windows and
    'put them on their right place...
    Call subSetMoveListWindow
    Call subSetCapturPiecesWindow
    Call subSetClockWindow
    
    frmChessBoard.SetFocus
    
    'Set each of the windows visible
    'or not according to the saved
    'settings...
    If (mnuEngineThinking.Checked = False) Then
        frmMiniChessBoard.Hide
    End If
    
    If (mnuMoveList.Checked = False) Then
        frmMoveList.Hide
    End If
    
    If (mnuChessClock.Checked = False) Then
        frmChessClock.Hide
    End If
    
    If (mnuCaptPieces.Checked = False) Then
        frmCapturedPieces.Hide
    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'Call agame_over_Click
    End
    
End Sub

Private Sub mnuNewGame_Click()
    'Restart the move count...
    intMoveCount = 1
    
    'Reset the captured pieces count...
    aryCaptPiecesCount(0) = 0
    aryCaptPiecesCount(1) = 0
    
    'aryCaptPiecesCount(2) = 0
    
    'Reset chess board to its initial position...
    Call subResetChessBoard
    Call subResetMiniChessBoard
    
    'Put the following windows to
    'their original position.
    Call subSetMoveListWindow
    Call subSetCapturPiecesWindow
    Call subSetClockWindow
    
    'frmChessBoard.SetFocus
    
    'Check the menu to see if there is
    'a window that was set invisible...
    If (mnuEngineThinking.Checked = False) Then
        frmMiniChessBoard.Hide
        
    End If
    
    If (mnuMoveList.Checked = False) Then
        frmMoveList.Hide
        
    End If
    
    If (mnuChessClock.Checked = False) Then
        frmChessClock.Hide
        
    End If
    
    If (mnuCaptPieces.Checked = False) Then
        frmCapturedPieces.Hide
        
    End If
    
    
    frmChessBoard.SetFocus
    'Empty the Move List...
    frmMoveList.txtMoveList.Text = ""
    
    'Erase the captured pieces...
    frmCapturedPieces.Cls
    
    'Stop the Timer that tracks the
    'duration of the game...
    frmChessClock.subStopClock
    
    'Reset the display of the Chess Clock...
    frmChessClock.subResetClocks
    
    'Tell the Chess Engine to stop...
    bolChessEngineTurn = False

End Sub

Private Sub mnuExit_Click()
    End
    
End Sub

Private Sub mnuSaveSettings_Click()
    
    'Save settings from game...
    Open App.Path & "\sah.ini" For Output As #1
    
    'najprej polozaji oken
    Print #1, "Chess Board:" & frmChessBoard.Left & "," & frmChessBoard.Top
    Print #1, "PC:" & frmMiniChessBoard.Left & "," & frmMiniChessBoard.Top & "," & mnuEngineThinking.Checked
    Print #1, "Move List:" & frmMoveList.Left & "," & frmMoveList.Top & "," & mnuMoveList.Checked
    Print #1, "Chess Clock:" & frmChessClock.Left & "," & frmChessClock.Top & "," & mnuChessClock.Checked
    Print #1, "Zajete:" & frmCapturedPieces.Left & "," & frmCapturedPieces.Top & "," & mnuCaptPieces.Checked
    Print #1, "Število nivojev:" & frmSettings.srbGameLevel.Value
    Print #1, "Naèin igre:" & frmSettings.srbPlayingStyle.Value
    
    Close
    
End Sub

Private Sub MoveClock_Click()
    'This menu button was created
    'just to bring the chess clock
    'into view. The author created
    'the game and on a large computer
    'screen. While setting the clock
    'window, he left it very low on
    'the bottom. When other users
    'open his project on a smaller
    'screen, the clock ends up out
    'of view! That's why he decided
    'to create this button!!
    frmChessClock.Top = 100
    frmChessClock.Left = 100
    frmChessClock.Visible = True
    frmChessClock.SetFocus
    
End Sub

Private Sub setings_Click()
    'This command will open the
    'Settings window in "Modal" state...
    frmSettings.Show 1
    
End Sub

Private Sub mnuCaptPieces_Click()
    'It will show or hide the
    '"Captured Pieces" window
    'according to the menu set...
    If (mnuCaptPieces.Checked = True) Then
        frmCapturedPieces.Hide
        mnuCaptPieces.Checked = False
        frmChessBoard.SetFocus
        
    Else
        frmCapturedPieces.Show
        mnuCaptPieces.Checked = True
        frmChessBoard.SetFocus
        
    End If
    
End Sub

Private Sub mnuChessClock_Click()
    'It will show or hide the
    '"Chess Clock" window
    'according to the menu set...
    If (mnuChessClock.Checked = True) Then
        frmChessClock.Hide
        mnuChessClock.Checked = False
        frmChessBoard.SetFocus
        
    Else
        frmChessClock.Show
        mnuChessClock.Checked = True
        frmChessBoard.SetFocus
        
    End If
    
End Sub

Private Sub mnuMoveList_Click()
    'It will show or hide the
    '"Move List" window
    'according to the menu set...
    If (mnuMoveList.Checked = True) Then
        frmMoveList.Hide
        mnuMoveList.Checked = False
        frmChessBoard.SetFocus
    Else
        frmMoveList.Show
        mnuMoveList.Checked = True
        frmChessBoard.SetFocus
    End If

End Sub

Private Sub mnuEngineThinking_Click()
    'It will show or hide the
    '"Engine Thinking" window
    'according to the menu set...
    If mnuEngineThinking.Checked = True Then
        frmMiniChessBoard.Hide
        mnuEngineThinking.Checked = False
        frmChessBoard.SetFocus
        
    Else
        frmMiniChessBoard.Show
        mnuEngineThinking.Checked = True
        frmChessBoard.SetFocus
        
    End If
    
End Sub
