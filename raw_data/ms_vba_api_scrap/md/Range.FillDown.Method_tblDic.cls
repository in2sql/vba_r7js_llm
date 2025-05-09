VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tblDic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'TblDic Class Module

Private lstObj As ListObject
Private tblDic As Object
Private columnFormulas As Object
Private rowSearchCriteria As Object
Private tblDicEmpty As Boolean
Private currentRowDic As RowDictionaryWrapper
Private keyCols As Collection
Private lastDelimiter As String
Private dumpedFormulas As Object
Private lastRowKey As Long
Private pValuesOnly As Boolean
Private updateColumnSpecified As Boolean
Private updateColumnName As String

' Regular Class Initialization
Private Sub Class_Initialize()
    Set keyCols = New Collection
    Set dumpedFormulas = CreateObject("Scripting.Dictionary")
    lastDelimiter = "|"
    Set tblDic = CreateObject("Scripting.Dictionary")
    Set columnFormulas = CreateObject("Scripting.Dictionary")
    Set rowSearchCriteria = CreateObject("Scripting.Dictionary")
    tblDicEmpty = True
    lastRowKey = 0
End Sub
' Call this method before buildDic to specify a single column update
Public Sub SpecifyUpdateColumn(ColumnName As String, Optional keyColumn As String = "")
    updateColumnSpecified = True
    updateColumnName = ColumnName
    If keyColumn <> "" Then
        keyColumnName = keyColumn
    End If
End Sub
' Initialization Method acting as a Constructor
Public Sub Init(tbl As ListObject)
''    Set lstObj = tbl
''    buildDic ' Automatically build the dictionary after setting the table
End Sub
Public Property Get rowDic() As RowDictionaryWrapper
    Set rowDic = currentRowDic
End Property
'Public Property Let rowDic(value As String)
'If tblDic.Exists(value) Then
'    Set currentRowDic = tblDic(value)
'End If
'End Property
'Public Property Let rowDic(updatedRowDic As RowDictionaryWrapper)
'    Set rowDic = updatedRowDic
'End Property
' Returns a ListColumn object given its name
Public Property Get lo() As ListObject
     Set lo = lstObj
End Property


Public Function getLC(ColName As String) As ListColumn
    On Error Resume Next ' In case the column name does not exist
    Set getLC = lstObj.ListColumns(ColName)
    On Error GoTo 0 ' Turn back on default error handling

    If getLC Is Nothing Then
        Err.Raise vbObjectError + 515, "TblDic::getLC", "Column '" & ColName & "' not found."
    End If
End Function

Public Sub AddKeyCol(ColumnName As String, Optional DELIMITER As String = "|")
    keyCols.Add ColumnName
    lastDelimiter = DELIMITER
End Sub

Public Sub ClearKeyCols()
    Set keyCols = New Collection
    lastDelimiter = "|"
End Sub
Public Function GetKeyCols() As Collection
    Set GetKeyCols = keyCols
End Function

Public Function GetLastDelimiter() As String
    GetLastDelimiter = lastDelimiter
End Function
Private Function GenerateRowKey(row As Variant) As String
    Dim key As String
    Dim ColName As Variant
    
    ' Handle case when row is a ListRow and key columns are specified
    If TypeName(row) = "ListRow" And keyCols.Count > 0 Then
        For Each ColName In keyCols
            If Len(key) > 0 Then key = key & lastDelimiter
            key = key & row.Range(1, lstObj.ListColumns(ColName).Index).value
        Next ColName
    ' Handle case when row is a RowDictionaryWrapper and key columns are specified
    ElseIf TypeName(row) = "RowDictionaryWrapper" And keyCols.Count > 0 Then
            For Each ColName In keyCols
                If Len(key) > 0 Then key = key & lastDelimiter
                ' Ensure RowDictionaryWrapper has a way to get values by column name
                key = key & row.GetValue(CStr(ColName))
            Next ColName
    Else
        ' Increment lastRowKey and use it as the new key
        lastRowKey = lastRowKey + 1
        key = CStr(lastRowKey)
    End If
    
    GenerateRowKey = key
End Function


Public Sub buildDic(tbl As ListObject, Optional fromScratch As Boolean = False, Optional valuesOnly As Boolean = False)
    Set lstObj = tbl
    Dim rowKey As String
    Dim lr As listRow
    Dim rowCount As Long
    Dim counter As Long

    pValuesOnly = valuesOnly
    
    If fromScratch Then
        If tbl.ListRows.Count > 0 Then
            tbl.DataBodyRange.Delete
        End If
    End If
    
    ' Ensure columnFormulas is ready to capture new formulas
    columnFormulas.RemoveAll
    CaptureColumnFormulas tbl
    
    tblDic.RemoveAll
    tblDicEmpty = True
    rowCount = lstObj.ListRows.Count

    If rowCount = 0 Then
        ' Add a RowDictionaryWrapper with columnFormulas if no rows exist
        Dim emptyWrapper As New RowDictionaryWrapper
        Call emptyWrapper.InitializeWithFormulas(columnFormulas)
        tblDic.Add "TBL_FORMULAS", emptyWrapper
    Else
        ' When building tblDic from existing rows
        counter = 0
        For Each lr In lstObj.ListRows
            counter = counter + 1
            Dim rowWrapper As New RowDictionaryWrapper
            Set rowWrapper = New RowDictionaryWrapper
            rowKey = GenerateRowKey(lr)
            
            ' Populate rowWrapper with cell data or formulas
            Call rowWrapper.InitializeFromExisting(lr, valuesOnly)
            ' Check if this is the last row
            If counter = rowCount Then
                rowWrapper.isLastItem = True
            End If
            If rowKey = "" Then
                tblDic.Add CStr(counter), rowWrapper
            Else
                tblDic.Add rowKey, rowWrapper
            End If
        Next lr
    End If
End Sub
Public Sub RebuildDic(Optional valuesOnly As Boolean = False)
    ' Assume lstObj is already set to the correct ListObject
    Dim rowKey As String
    Dim lr As listRow
    Dim counter As Long
    Dim rowCount As Long
    
    counter = 0
    rowCount = lstObj.ListRows.Count
    
    pValuesOnly = valuesOnly
    
    ' Clear the existing tblDic without touching columnFormulas or other settings
    tblDic.RemoveAll
    tblDicEmpty = (rowCount = 0)
    
    ' Rebuild the dictionary from the sorted ListObject's rows
    For Each lr In lstObj.ListRows
        counter = counter + 1
        Dim rowWrapper As New RowDictionaryWrapper
        Set rowWrapper = New RowDictionaryWrapper
        rowKey = GenerateRowKey(lr)
        
        ' Initialize rowWrapper with the current row's data or formulas
        Call rowWrapper.InitializeFromExisting(lr, pValuesOnly)
        
        ' Mark the last item, useful for certain operations that depend on row order
        If counter = rowCount Then rowWrapper.isLastItem = True
        
        ' Add or replace the rowWrapper in tblDic with the new key
        tst = rowWrapper.GetValue("TICKER")
        tblDic.Add rowKey, rowWrapper
    Next lr
End Sub

' Returns the internal dictionary object to allow access to its keys and other properties/methods
Public Function getTblDic() As Object
    Set getTblDic = tblDic
End Function
Public Sub CaptureColumnFormulas(tbl As ListObject)
    Dim tempRow As listRow
    Dim col As ListColumn
    Dim cell As Range
    Dim delTempRow As Boolean
    Dim backupFormDic As Object
    Set backupFormDic = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    Dim backupTbl As ListObject
    Dim lr As listRow
    Dim indexKey As String
    Dim formValue As String

    ' Check for the backup formulas sheet and table
    On Error Resume Next ' Use this to avoid error if sheet or table doesn't exist
    Set ws = b_FM
    Set backupTbl = ws.ListObjects("formula_tbl")
    On Error GoTo 0 ' Turn back on default error handling

    ' If backup table exists, populate the backupFormDic
    If Not backupTbl Is Nothing Then
        For Each lr In backupTbl.ListRows
            indexKey = lr.Range(1, 1).value ' Assuming INDEX is the first column
            formValue = lr.Range(1, 4).value ' Assuming FORM is the third column
            backupFormDic(indexKey) = formValue
        Next lr
    End If

    If tbl.ListRows.Count = 0 Then
        Set tempRow = tbl.ListRows.Add
        delTempRow = True
    Else
        Set tempRow = tbl.ListRows(1)
    End If

    For Each col In tbl.ListColumns
        Set cell = tempRow.Range.Cells(1, col.Index)
        If cell.HasFormula And Not dumpedFormulas.Exists(col.name) Then
            columnFormulas(col.name) = cell.Formula
        Else
            indexKey = tbl.name & "|" & col.name ' Construct the index key for backup lookup
            If backupFormDic.Exists(indexKey) Then
                columnFormulas(col.name) = "=" & backupFormDic(indexKey) ' Use backup formula if available
            Else
                columnFormulas(col.name) = vbNullString
            End If
        End If
    Next col

    If delTempRow Then tempRow.Delete
End Sub

Public Sub DumpFormulas(ColumnName As String)
    ' Add columnName to a collection to keep track of columns with dumped formulas
    On Error Resume Next ' Prevent error if item already exists
    dumpedFormulas.Add ColumnName, ColumnName
    On Error GoTo 0 ' Turn back on default error handling
End Sub
' Clears any existing search criteria
Public Sub ClearSearchCriteria()
    rowSearchCriteria.RemoveAll
End Sub

' Adds search criteria for finding rows
Public Sub AddRowSearch(ColumnName As String, value As Variant, Optional comparison As String = "=")
    Dim criteria As Object
    Set criteria = CreateObject("Scripting.Dictionary")
    criteria.Add "columnName", ColumnName
    criteria.Add "value", value
    criteria.Add "comparison", comparison
    rowSearchCriteria.Add criteria, ColumnName
End Sub

Public Function FindRowKeyBySearch() As String

    Dim matchedKey As Variant

    ' Searching for a matching row
    For Each key In tblDic.Keys
        If MatchesCriteria(tblDic(key).GetDictionary()) Then
            FindRowKeyBySearch = CStr(key)
            Exit For
        End If
    Next key

End Function
Public Function FindOrAddRowByKey(key As String) As Boolean
    ' Check if a row with the provided key already exists in the dictionary
    If tblDic.Exists(key) Then
        ' If exists, set currentRowDic to the existing RowDictionaryWrapper
        Set currentRowDic = tblDic(key)
        FindOrAddRowByKey = False ' Indicates the row already exists, no new row added
    Else
        ' If the key does not exist, create a new RowDictionaryWrapper
        Set currentRowDic = New RowDictionaryWrapper
        ' Initialize the new RowDictionaryWrapper with column formulas
        Call currentRowDic.InitializeWithFormulas(columnFormulas)
        
        ' Add the new RowDictionaryWrapper to tblDic with the provided key
        tblDic.Add key, currentRowDic
        FindOrAddRowByKey = True ' Indicates a new row has been added to tblDic
    End If
End Function
Public Function FindByKey(key As String) As Boolean
    ' Check if a row with the provided key already exists in the dictionary
    If tblDic.Exists(key) Then
        ' If exists, set currentRowDic to the existing RowDictionaryWrapper
        Set currentRowDic = tblDic(key)
        FindByKey = True ' Indicates rowDic found
    Else

        FindByKey = False ' Indicates rowDic not found
    End If
End Function
Public Sub UpdateRowDicValues(newRowDic As RowDictionaryWrapper)
    Dim ColName As Variant
    
    ' Ensure currentRowDic is initialized
    If currentRowDic Is Nothing Then
        Set currentRowDic = New RowDictionaryWrapper
        ' Optionally, initialize currentRowDic with default values/formulas
    End If
    
    ' Iterate through updated columns in newRowDic
    For Each ColName In newRowDic.UpdatedColumns
        ' Update or add the value from newRowDic to currentRowDic
        currentRowDic.SetValue CStr(ColName), newRowDic.GetValue(CStr(ColName))
    Next ColName
End Sub


' Inside TblDic Class Module
' In TblDic class
Public Function AddRow(key As String) As Boolean

    Set currentRowDic = New RowDictionaryWrapper
    Call currentRowDic.InitializeWithFormulas(columnFormulas) ' Initialize with column formulas
    
    FindOrAddRow = True ' Indicates a new rowDic needs to be added to tblDic

End Function

Public Sub AddRowWrapper(rowWrapper As RowDictionaryWrapper)
    ' Generate the key based on the content of rowWrapper
    Dim key As String
    key = GenerateRowKey(rowWrapper) ' Use Option 2's approach here

    If Not tblDic.Exists(key) Then
        tblDic.Add key, rowWrapper
    End If
End Sub

'''Public Sub AddRowWrapper(rowWrapper As RowDictionaryWrapper)
'''Dim key As Variant
'''key = GenerateRowKey
'''    If Not tblDic.Exists(key) Then
'''        tblDic.Add key, rowWrapper
'''    End If
'''End Sub



Private Function MatchesCriteria(ByRef rowDic As Object) As Boolean
    Dim criteria As Object
    MatchesCriteria = True ' Assume match until proven otherwise
    
    For Each criteria In rowSearchCriteria
        Dim ColumnName As String: ColumnName = criteria("columnName")
        Dim expectedValue As Variant: expectedValue = criteria("value")
        Dim comparison As String: comparison = criteria("comparison")
        
        Dim actualValue As Variant: actualValue = rowDic(ColumnName)
        
        ' Ensure numeric comparison for certain operators
        If IsNumeric(actualValue) And IsNumeric(expectedValue) Then
            Select Case comparison
                Case "="
                    If actualValue <> expectedValue Then MatchesCriteria = False
                Case ">"
                    If Not (actualValue > expectedValue) Then MatchesCriteria = False
                Case "<"
                    If Not (actualValue < expectedValue) Then MatchesCriteria = False
                Case ">="
                    If Not (actualValue >= expectedValue) Then MatchesCriteria = False
                Case "<="
                    If Not (actualValue <= expectedValue) Then MatchesCriteria = False
            End Select
        ElseIf comparison = "=" Then ' Fallback to string comparison for "="
            If CStr(actualValue) <> CStr(expectedValue) Then MatchesCriteria = False
        Else
            ' For non-numeric values, other comparisons may not make logical sense
            ' Consider how to handle or raise an error
            Err.Raise vbObjectError + 514, "MatchesCriteria", _
                "Non-numeric values encountered for comparison: " & comparison
        End If
        
        If Not MatchesCriteria Then Exit Function
    Next criteria
End Function


Private Function CreateRowDicFromFormulas() As Object
    Dim newRowDic As Object
    Set newRowDic = CreateObject("Scripting.Dictionary")
    Dim ColName As Variant
    
    For Each ColName In columnFormulas.Keys
        newRowDic(ColName) = columnFormulas(ColName)
    Next ColName
    
    ' Initialize other columns to default values if necessary
    ' For Each colName in lstObj.ListColumns...
    '     If Not newRowDic.Exists(colName) Then newRowDic(colName) = "DefaultValue"
    
    Set CreateRowDicFromFormulas = newRowDic
End Function


' Sets or updates the value for a specified column in the focused dictionary (row) within tblDic
' Sets the value for a specified column in every RowDictionaryWrapper within tblDic
Public Sub SetValue(ColumnName As String, value As Variant)
    Dim key As Variant
    Dim rowWrapper As RowDictionaryWrapper
    
    ' Iterate through all keys in tblDic
    For Each key In tblDic.Keys
        ' Retrieve the RowDictionaryWrapper by key
        Set rowWrapper = tblDic(key)
        
        ' Set the value for the specified column in the RowDictionaryWrapper
        rowWrapper.SetValue ColumnName, value
    Next key
End Sub
' Retrieves the value for a specified column in the row identified by the key
Public Function GetValue(key As String, ColumnName As String) As Variant
    ' Check if the row with the provided key exists
    If tblDic.Exists(key) Then
        ' Retrieve the RowDictionaryWrapper associated with the given key
        Dim rowWrapper As RowDictionaryWrapper
        Set rowWrapper = tblDic(key)
        
        ' Retrieve and return the value for the specified column name
        If rowWrapper.Exists(ColumnName) Then
            GetValue = rowWrapper.GetValue(ColumnName)
        Else
            ' Return vbNullString or another default value if the column does not exist
            GetValue = vbNullString
        End If
    Else
        ' Handle the case where the key does not exist
        Err.Raise vbObjectError + 514, "GetValue", "Key not found in dictionary."
    End If
End Function

Public Sub SetValueByRecType(lookUpColumn As String, lookupValue As Variant, targetColumn As String, targetValue As Variant)
    Dim key As Variant
    Dim rowWrapper As RowDictionaryWrapper

    For Each key In tblDic.Keys
        Set rowWrapper = tblDic(key)
        If rowWrapper.GetDictionary().Exists(lookUpColumn) Then
            If rowWrapper.GetDictionary()(lookUpColumn) = lookupValue Then
                rowWrapper.SetValue targetColumn, targetValue
            End If
        End If
    Next key
End Sub

Public Sub AddToRowValue(key As Variant, ColumnName As String, value As Variant)
If value <> 0 Then
    If tblDic.Exists(key) Then
        ' Retrieve the RowDictionaryWrapper associated with the given key
        Dim rowWrapper As RowDictionaryWrapper
        Set rowWrapper = tblDic(key)
        
        ' Check if the columnName exists in the row's dictionary
        If rowWrapper.GetDictionary().Exists(ColumnName) Then
            ' Assuming the value is numeric. Perform addition
            Dim currentValue As Variant
            currentValue = rowWrapper.GetDictionary()(ColumnName)
            If Not IsNumeric(currentValue) Then currentValue = 0
            
            If IsNumeric(value) Then
                currentValue = currentValue + value
                ' Update the value by adding the provided value to the current value
                rowWrapper.GetDictionary()(ColumnName) = currentValue
            Else
                Err.Raise vbObjectError + 1, "AddToRowValue", "Value is not numeric."
            End If
        Else
            ' Optional: Handle the case where the column does not exist in the row
            ' For example, you might want to initialize the value for this column
            ' rowWrapper.GetDictionary()(columnName) = value
        End If
    Else
        ' Handle case where the rowKey does not exist; could raise an error or simply ignore
        Err.Raise vbObjectError + 2, "AddToRowValue", "Key does not exist."
    End If
End If
End Sub

Public Sub UpdateOrAddRow(key As Variant)
    Dim rowDic As RowDictionaryWrapper
    
    ' Check if the key exists in tblDic, throw an error if not
    If tblDic.Exists(key) Then
        Set rowDic = tblDic(key)
    Else
        Err.Raise vbObjectError + 513, "UpdateOrAddRow", "Key does not exist in tblDic."
    End If
    
    ' Convert rowDic to an array for updating or adding a row
    Dim rowData As Variant
    rowData = rowDic.ToArray(lstObj)  ' Ensure ToArray method matches ListObject structure
    
    ' If pValuesOnly is True, replace values with formulas from columnFormulas where applicable
    If pValuesOnly Then
        Dim colIndex As Long
        For colIndex = 1 To UBound(rowData, 2)
            Dim ColName As String: ColName = lstObj.ListColumns(colIndex).name
            If columnFormulas.Exists(ColName) And columnFormulas(ColName) <> vbNullString Then
                rowData(1, colIndex) = columnFormulas(ColName)
            End If
        Next colIndex
    End If
    
    Dim found As Boolean: found = False
    Dim listRow As listRow
    Dim targetRange As Range
    
    ' Attempt to find and update the row in the ListObject
    For Each listRow In lstObj.ListRows
        If listRowMatchesKey(listRow, key) Then
            ' Efficiently update the entire row at once, considering the pValuesOnly flag
            Set targetRange = listRow.Range.Cells(1, 1).Resize(1, UBound(rowData, 2))
            targetRange.value = rowData
            found = True
            Exit For
        End If
    Next listRow

    ' If the row wasn't found, add a new row with the rowData
    If Not found Then
        Set listRow = lstObj.ListRows.Add
        Set targetRange = listRow.Range.Cells(1, 1).Resize(1, UBound(rowData, 2))
        targetRange.value = rowData
    End If
End Sub
Public Function UpdateListRow(key As Variant) As Boolean ' TRUE indicates key was found
    Dim updatedRowDic As RowDictionaryWrapper
    
    ' Check if the key exists in tblDic, throw an error if not
    If tblDic.Exists(key) Then
        Set updatedRowDic = tblDic(CStr(key))
    Else
        Err.Raise vbObjectError + 513, "UpdateOrAddRow", "Key does not exist in tblDic."
    End If
    
    ' Convert rowDic to an array for updating or adding a row
    Dim rowData As Variant
    rowData = updatedRowDic.ToArray(lstObj)  ' Ensure ToArray method matches ListObject structure
    
    ' If pValuesOnly is True, replace values with formulas from columnFormulas where applicable
    If pValuesOnly Then
        Dim colIndex As Long
        For colIndex = 1 To UBound(rowData, 2)
            Dim ColName As String: ColName = lstObj.ListColumns(colIndex).name
            If columnFormulas.Exists(ColName) And columnFormulas(ColName) <> vbNullString Then
                rowData(1, colIndex) = columnFormulas(ColName)
            End If
        Next colIndex
    End If
    
    Dim found As Boolean: found = False
    Dim listRow As listRow
    Dim targetRange As Range
    
    ' Attempt to find and update the row in the ListObject
    For Each listRow In lstObj.ListRows
        If listRowMatchesKey(listRow, key) Then
            ' Efficiently update the entire row at once, considering the pValuesOnly flag
            Set targetRange = listRow.Range.Cells(1, 1).Resize(1, UBound(rowData, 2))
            targetRange.value = rowData
            found = True
            Exit For
        End If
    Next listRow
    
    UpdateListRow = found

End Function
Public Sub updateTbl()
    If tblDic.Count = 1 Then Exit Sub ' if only 1: the one is TBL_FORMULAS

    Dim Data() As Variant
    Dim actualCount As Long
    actualCount = tblDic.Count - IIf(tblDic.Exists("TBL_FORMULAS"), 1, 0) ' Subtract 1 if TBL_FORMULAS exists
    
    ReDim Data(1 To actualCount, 1 To lstObj.ListColumns.Count)

    Dim rowIndex As Long: rowIndex = 1
    Dim colIndex As Long
    Dim key As Variant
    Dim rowWrapper As RowDictionaryWrapper
    Dim cName As Variant
    
    For Each key In tblDic.Keys
        If CStr(key) <> "TBL_FORMULAS" Then
            Set rowWrapper = tblDic(key)

            For colIndex = 1 To lstObj.ListColumns.Count
                Dim ColName As String: ColName = lstObj.ListColumns(colIndex).name
                
                ' Check against columnFormulas and write the formula instead of the value where applicable
                If pValuesOnly And columnFormulas.Exists(ColName) And columnFormulas(ColName) <> "" Then
                    Data(rowIndex, colIndex) = "" ' Use formula from columnFormulas
                ElseIf rowWrapper.GetDictionary().Exists(ColName) Then
                    Data(rowIndex, colIndex) = rowWrapper.GetDictionary()(ColName) ' Use value from rowWrapper
                Else
                    Data(rowIndex, colIndex) = "" ' Default value if column name not found
                End If
            Next colIndex

            rowIndex = rowIndex + 1
        End If
    Next key

    ' Update the Excel table with the data array
    If lstObj.ListRows.Count > 0 Then lstObj.DataBodyRange.Delete
    lstObj.Range(2, 1).Resize(UBound(Data, 1), UBound(Data, 2)).value = Data

    ' Remove the extra blank row if it exists
    If lstObj.ListRows.Count > actualCount Then
        lstObj.ListRows(lstObj.ListRows.Count).Delete
    End If
    
    ' Now loop through columnFormulas and apply formulas to the first row
    For Each cName In columnFormulas.Keys
        If columnFormulas(cName) <> vbNullString Then
            colIndex = lstObj.ListColumns(cName).Index
            lstObj.DataBodyRange.Cells(1, colIndex).Formula = columnFormulas(cName)
            ' Use Excel's fill feature to apply the formula down the column if the rest of the cells are blank
            'lstObj.ListColumns(cName).DataBodyRange.FillDown
        End If
    Next cName
End Sub



Private Function listRowMatchesKey(listRow As listRow, key As Variant) As Boolean
    ' Generate a key for the ListRow based on the same logic as when adding to tblDic
    Dim generatedKey As String
    generatedKey = GenerateRowKey(listRow) ' Assume GenerateRowKey is adjusted to handle both listRow and rowWrapper
    
    listRowMatchesKey = (generatedKey = key)
End Function



