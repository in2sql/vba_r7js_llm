VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAS400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

''''''''''''''''''''''''''''''''''''''''''''''''''
' Additional documentation for the automation objects can be found here: _
    https://www.ibm.com/support/knowledgecenter/SSEQ5Y_6.0.0/com.ibm.pcomm.doc/books/html/host_access08.htm

'       Created By:     Den Temple | dentemple.io
'       Creation Date:  May 8, 2015
'       Last Modified:  Nov 11, 2016
''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Error Constants
    Const CUSTOM_ERROR_CODE As Long = vbObjectError + 65535     'Maximum possible value
    Const ERROR_SOURCE As String = "cAS400.Property"
    Const ASSERT_ERROR_DESCRIPTION As String = "Either the AS400 is not active, or " & _
        "the AS400 attempted an action prior to the completion of a previous operation."
    Const SESSION_ERROR_DESCRIPTION As String = "The AS400 session name must be a single " & _
        "alphabetical character."
        
' AS400 magic values
    Const SESSION_NAME_LENGTH As Integer = 1
    Const SEARCH_FORWARD As Long = 1
    Const SEARCH_BACKWARD As Long = 2
    Const MINIMUM_MILLISECONDS As Long = 1

' iSeries automation objects (properties)
    Private pConnectionList As Object
    Private pConnectionManager As Object
    Private pOperatorInformation As Object
    Private pFieldList As Object
    Private pSessionInformation As Object
    Private pWindowMetrics As Object
    Private pPresentationSpace As Object


''''''''''''''''''''''''''''''''''''''''''''''''''
' I. Initialization & custom procedures
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitializeApp(ByVal sessionName As String)

' Sets the initial connection to the iSeries/AS400
' Input: sessionName (string). The iseries provides a single alphabetical _
    character to each application window, which can be called by name. _
    By default, the very first window is given the letter "A".
    
    If Len(sessionName) = SESSION_NAME_LENGTH Then
    
        ' Must be set independently
        Set pConnectionManager = CreateObject("PCOMM.autECLConnMgr")
        Set pSessionInformation = CreateObject("PCOMM.autECLSession")
        
        ' Note: autECLSession session must be set prior to sub-classing _
            any additional classes under it.
        pSessionInformation.SetConnectionByName (sessionName)
        
        ' Can be connected to a higher object for convenience
        Set pConnectionList = pConnectionManager.autECLConnList
        Set pOperatorInformation = pSessionInformation.autECLOIA
        Set pPresentationSpace = pSessionInformation.autECLPS
        Set pWindowMetrics = pSessionInformation.autECLWinMetrics
        
        ' Recommended by the documentation, but not required
        pConnectionList.Refresh
        
    Else
    
        Err.Raise _
            CUSTOM_ERROR_CODE, _
            ERROR_SOURCE, _
            SESSION_ERROR_DESCRIPTION
    
    End If
    
    
    
End Sub

Public Sub AssertReady()

    If Me.IsReady = False Then
    
        Dim inhibitedCode As Long
        inhibitedCode = Me.ReturnInhibitedCode
        Err.Raise _
            CUSTOM_ERROR_CODE, _
            ERROR_SOURCE, _
            ASSERT_ERROR_DESCRIPTION & vbNewLine & vbNewLine & _
                "Inhibited Code: " & CStr(inhibitedCode)
                
    End If
    
End Sub

Public Sub RunComponentTests()
    
    Debug.Print "== Run Test =="
    Debug.Print ""
    
    With Me

        Debug.Print "== App Status =="
        Debug.Print "Is Started: " & .IsStarted
        Debug.Print "Is Connected: " & .IsConnected
        Debug.Print "Is Enabled: " & .IsEnabled
        Debug.Print "Is Ready: " & .IsReady
        Debug.Print ""
        Debug.Print "== autECLPS Class =="
        Debug.Print "Rows & Cols Count: (" & .RowsCount & ", " & .ColumnsCount & ")"
        Debug.Print "Cursor Position: (" & .CursorPositionRow & ", " & .CursorPositionColumn & ")"
        Debug.Print "Current Session: " & .CurrentConnectionName
            
    End With
    
    Debug.Print ""
    Debug.Print "== End Test =="

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
' II. Direct automation object wrappers
''''''''''''''''''''''''''''''''''''''''''''''''''
' The following properties wrap only the automation objects. These _
    wrappers will be used internally by the class.
' These wrappers are public and can also be called outside the class. _
    This can be useful for experimentation and debugging purposes.
' EXAMPLE:
'   >>> Dim App As cAS400: Set App = New cAS400: App.InitializeAS400 ("A")
'   >>> App.PresentationSpaceObject.GetText 1,1,50
''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ConnectionManagerObject() As Object
    Set ConnectionManagerObject = pConnectionManager
End Property

Public Property Get SessionInformationObject() As Object
    Set SessionInformationObject = pSessionInformation
End Property

Public Property Get ConnectionListObject() As Object
    Set ConnectionListObject = pConnectionList
End Property

Public Property Get OperatorInformationObject() As Object
    Set OperatorInformationObject = pOperatorInformation
End Property

Public Property Get PresentationSpaceObject() As Object
    Set PresentationSpaceObject = pPresentationSpace
End Property

Public Property Get WindowMetricsObject() As Object
    Set WindowMetricsObject = pWindowMetrics
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''
' III. autECLConnList Class
' autECLConnList contains information about all started connections.
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshConnections()
    Me.ConnectionListObject.Refresh
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
' IV. autECLPS Class
' autECLPS performs operations on a presentation space.
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get RowsCount() As Long
    AssertReady
    RefreshConnections
    RowsCount = Me.PresentationSpaceObject.NumRows
End Property

Public Property Get ColumnsCount() As Long
    AssertReady
    RefreshConnections
    ColumnsCount = Me.PresentationSpaceObject.NumCols
End Property

Public Property Get CursorPositionRow() As Long
    AssertReady
    RefreshConnections
    CursorPositionRow = Me.PresentationSpaceObject.CursorPosRow
End Property

Public Property Get CursorPositionColumn() As Long
    AssertReady
    RefreshConnections
    CursorPositionColumn = Me.PresentationSpaceObject.CursorPosCol
End Property

Public Property Get CurrentConnectionName() As String
    AssertReady
    CurrentConnectionName = Me.PresentationSpaceObject.Name
End Property

Public Property Get IsStarted() As Boolean
    IsStarted = PresentationSpaceObject.Started
End Property

Public Property Get IsConnected() As Boolean
    IsConnected = Me.PresentationSpaceObject.CommStarted
End Property

Public Property Get IsEnabled() As Boolean
    IsEnabled = Me.PresentationSpaceObject.APIEnabled
End Property

Public Property Get IsReady() As Boolean
    ' Implicitely includes the .Started, .CommStarted, and .APIEnabled properties
    IsReady = Me.PresentationSpaceObject.Ready
End Property

Public Sub SetCursorPosition( _
    ByVal targetRow As Long, _
    ByVal targetCol As Long)
    
    AssertReady
    Me.PresentationSpaceObject targetRow, targetCol
    
End Sub

Public Sub SendKeys( _
    ByVal targetText As String, _
    Optional ByVal targetRow As Long = -1, _
    Optional ByVal targetCol As Long = -1, _
    Optional ByVal addWait As Long = 0)
    
    AssertReady
    
    If targetRow > 0 And targetCol > 0 Then
        Me.PresentationSpaceObject.SendKeys targetText, targetRow, targetCol
    Else
        Me.PresentationSpaceObject.SendKeys targetText
    End If
    
    If addWait > 0 Then
        Me.Wait addWait
    End If
    
End Sub

Public Property Get SearchText( _
    ByVal targetText As String, _
    Optional ByVal searchDirection As Long = SEARCH_FORWARD, _
    Optional ByVal targetRow As Long = -1, _
    Optional ByVal targetCol As Long = -1 _
    ) As Boolean
    
    AssertReady
    If startRow > 0 And startCol > 0 Then
        SearchText = Me.PresentationSpaceObject.SearchText( _
            targetText, searchDirection, targetRow, targetCol)
    Else
        SearchText = Me.PresentationSpaceObject.SearchText( _
            targetText, searchDirection)
    End If
    
End Property

Public Property Get GetText( _
    ByVal targetRow As Long, _
    ByVal targetCol As Long, _
    ByVal numberChars As Long)
    
    AssertReady
    GetText = Me.PresentationSpaceObject.GetText(targetRow, targetCol, numberChars)
    
End Property

Public Sub SetText( _
    ByVal targetText As String, _
    ByVal targetRow As Long, _
    ByVal targetCol As Long)
    
    AssertReady
    Me.PresentationSpaceObject.SetText targetText, targetRow, targetCol
    
End Sub

Public Sub Wait(Optional ByVal addMilliseconds As Long = MINIMUM_MILLISECONDS)
    
    Me.PresentationSpaceObject.Wait addMilliseconds
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''
' V. autECLOIA Class
' The autECLOIA object retrieves status from the Host Operator Information Area.

''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get ReturnInhibitedCode() As Long
    ReturnInhibitedCode = Me.OperatorInformationObject.InputInhibited
End Property
                                                                                                                    
Public Function WaitForInput() As Boolean
    'Note: The AS/400 returns False if the timeout is exceeded
    WaitForInput = Me.OperatorInformationObject.WaitForInputReady
End Function

Public Function WaitForApp() As Boolean
    'Note: The AS/400 returns False if the timeout is exceeded
    WaitForApp = Me.OperatorInformationObject.WaitForAppAvailable
End Function

Public Sub Pause(Optional addWait As Long = 1)
    
    Me.WaitForApp
    Me.WaitForInput
    Me.Wait addWait

End Sub
