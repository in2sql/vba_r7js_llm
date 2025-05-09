Attribute VB_Name = "getModels"
Option Explicit
Option Compare Text
Option Base 0
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       18th April 2015
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================

Function EntryPointGetConnections()
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       18th April 2015
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
'Constants
Const cstrTableName As String = "Table="
Const cstrRecordCount As String = "RecordCount="
Const cstrConnectionText As String = "ConnectionText="
Const cstrConnectionType As String = "ConnectionType="
Const cstrColumnName As String = "ColumnName="
Const cstrDataType As String = "DataType="
Const cstrPrimaryKeyTableName As String = "PrimaryKeyTableName="
Const cstrPrimaryKeyColumnName As String = "PrimaryKeyColumnName"
Const cstrForeignKeyTableName As String = "ForeignKeyTableName"
Const cstrForeignKeyColumnName As String = "ForeignKeyColumnName"
'Variables
Dim oWorkbook As Excel.Workbook
Dim strConnectionType As String
Dim oModel As Excel.Model
Dim oTable As Excel.ModelTable
Dim oTableColumn As Excel.ModelTableColumn
Dim oRelationship As Excel.ModelRelationship


Set oWorkbook = ThisWorkbook
Debug.Print "Relationships--------------------------------------"

For Each oRelationship In oWorkbook.Model.ModelRelationships
    With oRelationship
        Debug.Print .PrimaryKeyTable.Name, .PrimaryKeyColumn.Name, .ForeignKeyTable.Name, .ForeignKeyColumn.Name
    End With
Next


Debug.Print "---------------------------------------------------"
Debug.Print "Tables---------------------------------------------"
For Each oTable In oWorkbook.Model.ModelTables
    With oTable
        strConnectionType = DecodeConnectionType(.SourceWorkbookConnection.WorksheetDataConnection.CommandType)
        Debug.Print cstrTableName & .Name, cstrRecordCount & .RecordCount, cstrConnectionText & .SourceWorkbookConnection.WorksheetDataConnection.CommandText, cstrConnectionType & strConnectionType
        strConnectionType = ""
    End With
    For Each oTableColumn In oTable.ModelTableColumns
        With oTableColumn
            Debug.Print cstrColumnName & .Name, cstrDataType & DecodeParameter(.DataType)
        End With
    Next
    Debug.Print "---------------------------------------------------"
Next

End Function

Function DecodeParameter(Incoming As Excel.XlParameterDataType) As String
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       18th April 2015
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================

Select Case Incoming
    Case Excel.XlParameterDataType.xlParamTypeBigInt
        DecodeParameter = "xlParamTypeBigInt"
    Case Excel.XlParameterDataType.xlParamTypeBinary
        DecodeParameter = "xlParamTypeBinary"
    Case Excel.XlParameterDataType.xlParamTypeBit
        DecodeParameter = "xlParamTypeBit"
    Case Excel.XlParameterDataType.xlParamTypeChar
        DecodeParameter = "xlParamTypeChar"
    Case Excel.XlParameterDataType.xlParamTypeDate
        DecodeParameter = "xlParamTypeDate"
    Case Excel.XlParameterDataType.xlParamTypeDecimal
        DecodeParameter = "xlParamTypeDecimal"
    Case Excel.XlParameterDataType.xlParamTypeDouble
        DecodeParameter = "xlParamTypeDouble"
    Case Excel.XlParameterDataType.xlParamTypeFloat
        DecodeParameter = "xlParamTypeFloat"
    Case Excel.XlParameterDataType.xlParamTypeInteger
        DecodeParameter = "xlParamTypeInteger"
    Case Excel.XlParameterDataType.xlParamTypeLongVarBinary
        DecodeParameter = "xlParamTypeLongVarBinary"
    Case Excel.XlParameterDataType.xlParamTypeLongVarChar
        DecodeParameter = "xlParamTypeLongVarChar"
    Case Excel.XlParameterDataType.xlParamTypeNumeric
        DecodeParameter = "xlParamTypeNumeric"
    Case Excel.XlParameterDataType.xlParamTypeReal
        DecodeParameter = "xlParamTypeReal"
    Case Excel.XlParameterDataType.xlParamTypeSmallInt
        DecodeParameter = "xlParamTypeSmallInt"
    Case Excel.XlParameterDataType.xlParamTypeTime
        DecodeParameter = "xlParamTypeTime"
    Case Excel.XlParameterDataType.xlParamTypeTimestamp
        DecodeParameter = "xlParamTypeTimestamp"
    Case Excel.XlParameterDataType.xlParamTypeTinyInt
        DecodeParameter = "xlParamTypeTinyInt"
    Case Excel.XlParameterDataType.xlParamTypeUnknown
        DecodeParameter = "xlParamTypeUnknown"
    Case Excel.XlParameterDataType.xlParamTypeVarBinary
        DecodeParameter = "xlParamTypeVarBinary"
    Case Excel.XlParameterDataType.xlParamTypeVarChar
        DecodeParameter = "xlParamTypeVarChar"
    Case Excel.XlParameterDataType.xlParamTypeWChar
        DecodeParameter = "xlParamTypeWChar"
    Case Else
        DecodeParameter = "[UNKNOWN]"
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

