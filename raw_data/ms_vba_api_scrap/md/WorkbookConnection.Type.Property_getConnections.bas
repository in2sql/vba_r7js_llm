Attribute VB_Name = "getConnections"
Option Explicit
Option Compare Text
Option Base 0
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       6th April 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================

Function DecodeCommandType(Incoming As Excel.XlCmdType) As String
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       6th April 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================

Select Case Incoming
    Case Excel.XlCmdType.xlCmdCube
        DecodeCommandType = "xlCmdCube"
    Case Excel.XlCmdType.xlCmdDAX
        DecodeCommandType = "xlCmdDAX"
    Case Excel.XlCmdType.xlCmdDefault
        DecodeCommandType = "xlCmdDefault"
    Case Excel.XlCmdType.xlCmdExcel
        DecodeCommandType = "xlCmdExcel"
    Case Excel.XlCmdType.xlCmdList
        DecodeCommandType = "xlCmdList"
    Case Excel.XlCmdType.xlCmdSql
        DecodeCommandType = "xlCmdSql"
    Case Excel.XlCmdType.xlCmdTable
        DecodeCommandType = "xlCmdTable"
    Case Excel.XlCmdType.xlCmdTableCollection
        DecodeCommandType = "xlCmdTableCollection"
    Case Else
        DecodeCommandType = "[UNKNOWN]"
End Select


End Function

Function DecodeConnectionType(Incoming As Excel.XlConnectionType) As String
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       6th April 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================

Select Case Incoming
    Case Excel.XlConnectionType.xlConnectionTypeDATAFEED
        DecodeConnectionType = "DATAFEED"
    Case Excel.XlConnectionType.xlConnectionTypeMODEL
        DecodeConnectionType = "MODEL"
    Case Excel.XlConnectionType.xlConnectionTypeNOSOURCE
        DecodeConnectionType = "NOSOURCE"
    Case Excel.XlConnectionType.xlConnectionTypeODBC
        DecodeConnectionType = "ODBC"
    Case Excel.XlConnectionType.xlConnectionTypeOLEDB
        DecodeConnectionType = "OLEDB"
    Case Excel.XlConnectionType.xlConnectionTypeTEXT
        DecodeConnectionType = "TEXT"
    Case Excel.XlConnectionType.xlConnectionTypeWEB
        DecodeConnectionType = "WEB"
    Case Excel.XlConnectionType.xlConnectionTypeWORKSHEET
        DecodeConnectionType = "WORKSHEET"
    Case Excel.XlConnectionType.xlConnectionTypeXMLMAP
        DecodeConnectionType = ""
    Case Else
        DecodeConnectionType = "[UNKNOWN]"
End Select

End Function


Function EntryPointGetConnections()
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       6th April 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
Dim oWorkbook As Excel.Workbook
Dim oWorkbookConnection As Excel.WorkbookConnection
Dim oODBCConnection As Excel.ODBCConnection
Dim oOLEDBConnection As Excel.OLEDBConnection
Dim strConnectionSpecificString As String

Set oWorkbook = ThisWorkbook

For Each oWorkbookConnection In oWorkbook.Connections
    If oWorkbookConnection.Type = xlConnectionTypeODBC Then
        Set oODBCConnection = oWorkbookConnection.ODBCConnection
        With oODBCConnection
            strConnectionSpecificString = "[" & DecodeConnectionType(oWorkbookConnection.Type) & "] Command=[" & .CommandText & "] Connection=[" & .Connection & "] SourceConnectionFile=[" & .SourceConnectionFile & "] CommandType=[" & DecodeConnectionType(.CommandType) & "]"
        End With
    ElseIf oWorkbookConnection.Type = xlConnectionTypeOLEDB Then
        Set oOLEDBConnection = oWorkbookConnection.OLEDBConnection
        With oOLEDBConnection
            strConnectionSpecificString = "[" & DecodeConnectionType(oWorkbookConnection.Type) & "] Command=[" & .CommandText & "] Connection=[" & .Connection & "] SourceConnectionFile=[" & .SourceConnectionFile & "] CommandType=[" & DecodeConnectionType(.CommandType) & "]"
        End With
    ElseIf oWorkbookConnection.Type = xlConnectionTypeTEXT Then
        With oWorkbookConnection
            strConnectionSpecificString = "[" & DecodeConnectionType(oWorkbookConnection.Type) & "] Command=[" & .TextConnection.Connection & "]"
        End With
    End If
    Debug.Print strConnectionSpecificString & "[" & oWorkbookConnection.Description & "] Type=[" & DecodeCommandType(oWorkbookConnection.Type) & "]"
Next
End Function



