VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "ExcelFilter"
'----------------------------------------------------------------------------------------------------------
Private Const FILTER_OPERATOR As String = "operator"
Private Const CRITERIA_1 As String = "criteria1"
Private Const CRITERIA_2 As String = "criteria2"
Private Const FROM_SHEET As String = "fromSheet"
Private Const IS_DATE As String = "isDate"
'----------------------------------------------------------------------------------------------------------
Private pSheet As Excel.Worksheet
Private pHasFilter As Boolean
Private pIsActive As Boolean
Private pRange As Excel.Range
Private pFilters As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------



'[SETTERS]

Public Function setSheet(value As Excel.Worksheet) As ExcelFilter
    Set setSheet = Me
    Set pSheet = value
End Function

Public Sub readFromSheet(Optional wks As Excel.Worksheet)
    Dim filter As Excel.filter
    Dim Counter As Long
    '------------------------------------------------------------------------------------------------------
    
    If Not wks Is Nothing Then Set pSheet = wks
    
    If f.sheets.IsValid(pSheet) Then
        With pSheet
            '[Meta properties]
            pHasFilter = .AutoFilterMode
            pIsActive = .FilterMode
            
            '[Filters properties]
            If pIsActive Then
                With .AutoFilter
                    Set pRange = .Range
                    Set pFilters = f.dictionaries.Create(False)
                    For Each filter In .filters
                        Counter = Counter + 1
                        If filter.On Then
                            Call pFilters.Add(Counter, getFilterPropsDictionary(filter, Counter))
                        End If
                    Next filter
                End With
            End If
            
        End With
    End If
    
End Sub

Private Function getFilterPropsDictionary(filter As Excel.filter, index As Long) As Scripting.Dictionary
    Dim filterActive As Boolean
    Dim dict As Scripting.Dictionary
    Dim criteria As Scripting.Dictionary
    Dim key As Variant
    Dim criteriaValue As Variant
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    filterActive = filter.On
    '###if not errHandler
    
    If filterActive Then
        Set getFilterPropsDictionary = f.dictionaries.Create(False)
        
        With filter
            Call getFilterPropsDictionary.Add(FILTER_OPERATOR, f.Math.coalesce(.operator, xlFilterValues))
            '[Criteria1]
            Set criteria = f.dictionaries.Create(False)
            With criteria
                On Error Resume Next
                Call .Add(CRITERIA_1, filter.criteria1)
                Call .Add(CRITERIA_2, filter.criteria2)
                '###if not errHandler
                
                If f.dictionaries.isNonEmptyDictionary(criteria) Then
                    For Each key In .Keys
                        If Not f.Variables.isUndefined(.item(key)) Then
                            Call f.Variables.assign(criteriaValue, getFilterCriteriaValue(.item(key), filter))
                            If Not f.Variables.isUndefined(criteriaValue) Then
                                Call getFilterPropsDictionary.Add(key, f.dictionaries.createWithItems(False, KeyValue(VALUE_____, criteriaValue)))
                            End If
                        End If
                    Next key
                Else
                    Call f.Variables.assign(criteriaValue, getCriteriaValueFromSheet(filter, index))
                    If Not f.Variables.isUndefined(criteriaValue) Then
                        Call getFilterPropsDictionary.Add(FROM_SHEET, criteriaValue)
                    End If
                End If
                
            End With
        End With
    End If
    
End Function

Private Function getFilterCriteriaValue(value As Variant, filter As Excel.filter) As Variant
    Select Case filter.operator
        Case xlFilterValues:
            getFilterCriteriaValue = value
        Case xlFilterCellColor:
            getFilterCriteriaValue = value.color
        Case Else:
            getFilterCriteriaValue = value
    End Select
End Function

Private Function getCriteriaValueFromSheet(filter As Excel.filter, index As Long) As Scripting.Dictionary
    Dim Parent As AutoFilter
    Dim rng As Excel.Range
    Dim visibleCells As Excel.Range
    Dim cell As Excel.Range
    Dim dict As Scripting.Dictionary
    Dim isDateValue As Boolean
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    Set Parent = filter.Parent
    
    Set getCriteriaValueFromSheet = f.dictionaries.Create(False)
    
    If Not Parent Is Nothing Then
        Set rng = Parent.Range
        If f.ranges.IsValid(rng) Then
            If rng.columns.count >= index Then
                On Error Resume Next
                Set visibleCells = rng.columns(index).SpecialCells(xlCellTypeVisible)
                '### errHandler
            End If
        End If
        
        Set dict = f.dictionaries.Create(False)
        If f.ranges.IsValid(visibleCells) Then
            For Each cell In visibleCells
                If cell.row <> rng.row Then
                    isDateValue = VBA.IIf(isDateValue, isDateValue, VBA.IsDate(cell.value))
                    If Not dict.exists(cell.value2) Then
                        Call dict.Add(cell.value2, vbNullString)
                    End If
                End If
            Next cell
        End If
        
    End If
    
    If f.dictionaries.isNonEmptyDictionary(dict) Then
        Call getCriteriaValueFromSheet.Add(VALUE_____, dict.Keys)
        Call getCriteriaValueFromSheet.Add(IS_DATE, isDateValue)
    End If
    
End Function

Public Sub removeFilter()
    If f.sheets.IsValid(pSheet) Then
        If pSheet.FilterMode Then
            pSheet.AutoFilterMode = False
        End If
    End If
End Sub

Public Sub applyFilter()
    Dim index As Variant
    Dim dictCriteria As Scripting.Dictionary
    Dim operator As XlAutoFilterOperator
    Dim criteria1 As Scripting.Dictionary
    Dim criteria2 As Scripting.Dictionary
    Dim fromSheet As Scripting.Dictionary
    Dim IsDate As Boolean
    '------------------------------------------------------------------------------------------------------
    
    If f.sheets.IsValid(pSheet) Then
        If pHasFilter Then
            If pIsActive Then
                pSheet.AutoFilterMode = False
                Call pRange.AutoFilter
                'Call pRange.AutoFilter
                If f.dictionaries.isNonEmptyDictionary(pFilters) Then
                    For Each index In pFilters.Keys
                        If index > 0 And index < pRange.columns.count Then
                            Set dictCriteria = pFilters.item(index)
                            operator = f.dictionaries.getItem(dictCriteria, FILTER_OPERATOR)
                            Set criteria1 = f.dictionaries.getObject(dictCriteria, CRITERIA_1)
                            Set criteria2 = f.dictionaries.getObject(dictCriteria, CRITERIA_2)
                            Set fromSheet = f.dictionaries.getObject(dictCriteria, FROM_SHEET)
                            IsDate = f.dictionaries.getItem(fromSheet, IS_DATE)
                            
                            If Not criteria2 Is Nothing Then
                                Call pRange.AutoFilter(Field:=index, criteria1:=getCriteriaValueFromJson(f.dictionaries.getItem(dictCriteria, CRITERIA_1)), operator:=operator, criteria2:=getCriteriaValueFromJson(f.dictionaries.getItem(dictCriteria, CRITERIA_2)))
                            ElseIf Not criteria1 Is Nothing Then
                                Call pRange.AutoFilter(Field:=index, criteria1:=getCriteriaValueFromJson(f.dictionaries.getItem(dictCriteria, CRITERIA_1)), operator:=operator)
                            ElseIf Not fromSheet Is Nothing Then
                                'Call pRange.AutoFilter(Field:=index, criteria2:=getCriteriaValueFromJson(f.dictionaries.getItem(dictCriteria, CRITERIA_1), isDate), operator:=operator)
                                Call pRange.AutoFilter(Field:=index, criteria2:=getCriteriaValueFromJson(f.dictionaries.getItem(dictCriteria, fromSheet), IsDate), operator:=operator)
                            End If
                        End If
                    Next index
                    
                End If
                
            End If
        End If
    End If
    
End Sub

Public Function getCriteriaValueFromJson(dict As Scripting.Dictionary, Optional IsDate As Boolean = False) As Variant
    If f.dictionaries.isNonEmptyDictionary(dict) Then
        If Not dict.exists(VALUE_____) Then
        ElseIf IsDate Then
            getCriteriaValueFromJson = createDatesFilterArray(dict.item(VALUE_____))
        ElseIf VBA.IsObject(dict.item(VALUE_____)) Then
            Set getCriteriaValueFromJson = dict.item(VALUE_____)
        Else
            getCriteriaValueFromJson = dict.item(VALUE_____)
            If f.Variables.isString(getCriteriaValueFromJson) Then
                If f.strings.startsWith(VBA.CStr(getCriteriaValueFromJson), "=") Then
                    getCriteriaValueFromJson = VBA.mid$(getCriteriaValueFromJson, 2)
                End If
            End If
        End If
    End If
End Function

Public Function createDatesFilterArray(arr As Variant) As Variant
    Dim i As Long
    Dim col As VBA.Collection
    Dim d As Date
    '------------------------------------------------------------------------------------------------------
    
    If f.arrays.countDimensions(arr) = 1 Then
        Set col = New VBA.Collection
        For i = LBound(arr) To UBound(arr)
            d = arr(i)
            Call col.Add(2)
            Call col.Add(VBA.replace(VBA.format(d, "m/d/yyyy"), "-", "/"))
        Next i
    End If
    
    createDatesFilterArray = f.Collections.toArray(col)
    
End Function
