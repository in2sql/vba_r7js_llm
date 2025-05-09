VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Inside CodeSettings class module
Option Explicit

Private pCalcState As Long
Private pEventState As Boolean
Private pPageBreakState As Boolean
Private pSkip As Boolean
Private wasProtected As Boolean

Public Property Get CalcState() As Long
    CalcState = pCalcState
End Property

Public Property Get EventState() As Boolean
    EventState = pEventState
End Property

Public Property Get PageBreakState() As Boolean
    PageBreakState = pPageBreakState
End Property

Public Property Get Skip() As Boolean
    Skip = pSkip
End Property

Public Property Let Skip(value As Boolean)
    pSkip = value
End Property

Public Sub Code_Begin()
If Not pSkip Then
    On Error Resume Next
    Application.ScreenUpdating = False
    
    pEventState = Application.EnableEvents
    Application.EnableEvents = False
    
    pCalcState = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    pPageBreakState = ActiveSheet.DisplayPageBreaks
    ActiveSheet.DisplayPageBreaks = False
    Application.IgnoreRemoteRequests = True
End If
End Sub

Public Sub Code_End()
On Error Resume Next
If Not pSkip Then
    ActiveSheet.DisplayPageBreaks = pPageBreakState
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ' Call RefreshHeading() here if it's implemented as a method
    'ActiveSheet.EnableSelection = xlUnlockedCells
    ' Implement protection logic here if necessary
    Application.IgnoreRemoteRequests = False
Else
    pSkip = False
End If
End Sub

Public Sub protectWS(ws As Worksheet)
If wasProtected Then
    ws.Protect Password:=pwd, UserInterfaceOnly:=True, AllowFormattingCells:=True, _
                AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowInsertingColumns:=True, _
                AllowInsertingRows:=True, AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
                AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
End If
End Sub

Public Sub unprotectWs(ws As Worksheet)
If ws.ProtectContents Then
wasProtected = True
ws.Unprotect pwd
End If
End Sub
