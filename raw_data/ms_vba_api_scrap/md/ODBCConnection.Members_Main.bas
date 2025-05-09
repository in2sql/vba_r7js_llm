Attribute VB_Name = "Main"
Option Explicit

Sub WaitforWorksheetQueries()
    Sheet1.Range("A2").Clear
    
    Dim aListObject As ListObject
    For Each aListObject In Sheet1.ListObjects
        If Not aListObject.QueryTable Is Nothing Then
            Debug.Print aListObject.QueryTable.WorkbookConnection.Name
            aListObject.QueryTable.Refresh
        End If
    Next aListObject
    
    Application.CalculateUntilAsyncQueriesDone
    
    MsgBox "Done. Code execution did wait for this query to complete."
    
End Sub

Sub WaitforQueries()
    Sheet1.Range("A2").Clear
    
    ActiveWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    
    MsgBox "Done. Code execution did wait for all queries to complete."
    
End Sub

Sub DontWaitforQueries()
    Sheet1.Range("A2").Clear
    
    ActiveWorkbook.RefreshAll
    
    MsgBox "Code excecution didn't wait."
    
End Sub

Sub HandcraftedWaitforQueries()
    Sheet1.Range("A2").Clear
    
    RefreshAllWait
    
    MsgBox "Done. Code execution did wait for queries to complete."
    
End Sub

Sub RefreshAllWait(Optional TypeOfQuery As XlConnectionType)
   
    Dim qry As WorkbookConnection
    Dim thisConnectionType As Object
    Dim DefaultRefresh As Boolean
    
    Debug.Print VBA.String(200, vbNewLine)
    
    For Each qry In ActiveWorkbook.Connections
        
        If qry.Type = xlConnectionTypeODBC Then
            ' PowerQuery
            Set thisConnectionType = qry.ODBCConnection
        ElseIf qry.Type = xlConnectionTypeOLEDB Then
            ' PowerPivot and other connections
            Set thisConnectionType = qry.OLEDBConnection
        End If
        
        With thisConnectionType
            Debug.Print "Query " & qry.Name
            
            DefaultRefresh = .BackgroundQuery
            .BackgroundQuery = False
            .Refresh
            .BackgroundQuery = DefaultRefresh
        End With
        
    Next qry
    
End Sub
