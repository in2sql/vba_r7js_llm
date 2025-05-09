Attribute VB_Name = "modMisc"
                                                                                                                                          ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::                                   ' _
    ''''''''''''''''''''''s'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''                                   ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||             MISCELLANEOUS             ||||||||||||||||||||||||||||||||||                                   ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                   ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                   ' _
                                                                                                                                          ' _
    AUTHOR:   Kallun Willock                                                                                                              ' _
    PURPOSE:  Collection of misc. procedures                                                                                              ' _
                                                                                                                                          ' _
    VERSION:  1.0        25/05/2021                                                                                                       ' _
                                                                                                                                          ' _
    NOTES:    [�]                                                                                                                         ' _
                                                                                                                                          ' _
              -  [�]:      [�]                                                                                                            ' _
                                                                                                                                          ' _
    TODO:     [�]                                                                                                                         ' _
                                                                                                                                          ' _
    ...................................................................................................                                   ' _
    :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Sub INCREMENT(ByRef ItemValue As Long, Optional Amount As Long = 1)
    ItemValue = ItemValue + Amount
End Sub

Sub DECREMENT(ByRef ItemValue As Long, Optional Amount As Long = 1)
    ItemValue = ItemValue - Amount
End Sub

Function ISITODD(Target) As Boolean
    ISITODD = Application.WorksheetFunction.IsOdd(Target)
End Function

Function ISITEVEN(Target) As Boolean
    ISITEVEN = Application.WorksheetFunction.IsEven(Target)
End Function

Function HOWMANY(Source As String, TargetText As String) As Long
    If InStr(Source, TargetText) = 0 Then HOWMANY = 0: Exit Function
    FinalLength = Len(Source) - Len(Replace(Source, TargetText, ""))
    HOWMANY = FinalLength / Len(TargetText)
End Function

'  Procedures:   Information re: files and folders

Function GETFILENAME(FILENAME As String) As String
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GETFILENAME = FSO.GETFILENAME(FILENAME)
    Set FSO = Nothing
End Function

Function GETEXTENSION(FILENAME As String) As String
    On Error Resume Next
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GETEXTENSION = FSO.GetExtensionName(FILENAME)
    Set FSO = Nothing
End Function

Function GETPATH(FILENAME As String) As String
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GETPATH = Replace(FSO.GetAbsolutePathName(FILENAME), FSO.GETFILENAME(FILENAME), "", , , vbTextCompare)
    Set FSO = Nothing
End Function

'  Accss-specific procedures
Function PathSeparator() As String
    PathSeparator = Replace(Replace(Application.CurrentDb.Name, Application.CodeProject.Path, vbNullString), Application.CodeProject.Name, vbNullString)
End Function 
