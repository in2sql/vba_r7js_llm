VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelRelinker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const TAG_PATTERN As String = "[{({0})}]"
Private Const SHEET_FORMULA_SEPARATOR As String = "?!?!?!?"
'----------------------------------------------------------------------------------------------------------
Private pState As ExcelState
Private pBook As Excel.Workbook
Private pSheets As VBA.Collection
Private pOldPathToKey As Scripting.Dictionary
Private pKeyToNewPath As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------

Private Sub reset()
    Set pBook = Nothing
    Set pSheets = New VBA.Collection
    Set pOldPathToKey = f.dictionaries.Create
    Set pKeyToNewPath = f.dictionaries.Create
End Sub


Public Sub updateLinks(wkb As Excel.Workbook, dictLinks As Scripting.Dictionary, Optional sheetNames As Variant)
    Dim state As ExcelState:        state = f.Excel.getDefaultExcelState
    Dim sheets As VBA.Collection
    Dim groups As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    Call f.Excel.freeze(pState)
    
    Set pBook = wkb
    Set pSheets = f.Books.getSheets(pBook, sheetNames)
    
    Call createSubDictionaries(dictLinks)
    Set groups = getCellGroupsByFormula(pOldPathToKey)
    If f.dictionaries.isNonEmptyDictionary(groups) Then
        Call replacePathInFormulas(groups, pKeyToNewPath)
    End If
    
    Call f.Excel.unfreeze(pState)
    
End Sub




Private Sub createSubDictionaries(dictLinks As Scripting.Dictionary)
    Dim key As Variant
    Dim index As Long
    Dim tag As String
    Dim originalLinkXlsNotation As String
    Dim newLinkXlsNotation As String
    '------------------------------------------------------------------------------------------------------
    
    Set pOldPathToKey = f.dictionaries.Create(False)
    Set pKeyToNewPath = f.dictionaries.Create(False)
    
    For Each key In dictLinks.Keys
        index = index + 1
        tag = f.strings.format(TAG_PATTERN, index)
        originalLinkXlsNotation = f.files.adjustPathForExcelFormula(VBA.CStr(key))
        newLinkXlsNotation = f.files.adjustPathForExcelFormula(VBA.CStr(dictLinks.item(key)))
        Call f.dictionaries.addItem(pOldPathToKey, originalLinkXlsNotation, tag)
        Call f.dictionaries.addItem(pKeyToNewPath, tag, newLinkXlsNotation)
    Next key
    
End Sub


Private Function getCellGroupsByFormula(Optional ByVal filters As Scripting.Dictionary) As Scripting.Dictionary
    Const FORMULA_FILTER_PATTERN As String = "\\\[(.*)\]$"
    '------------------------------------------------------------------------------------------------------
    Dim adjustedFilters As Scripting.Dictionary
    Dim key As Variant
    Dim filter As String
    '------------------------------------------------------------------------------------------------------
    Dim wks As Excel.Worksheet
    Dim rng As Excel.Range
    Dim formulas As Variant
    Dim formula As String
    Dim i As Long
    Dim j As Long
    '------------------------------------------------------------------------------------------------------
    Dim groupKey As String
    Dim groupCells As VBA.Collection
    Dim qualified As Boolean
    '------------------------------------------------------------------------------------------------------
    Dim rngFormulas As Excel.Range
    Dim area As Excel.Range
    '------------------------------------------------------------------------------------------------------

    Set getCellGroupsByFormula = f.dictionaries.Create(False)
    
    'Modify formula filter to include differences between UNC and share drive path.
    Set adjustedFilters = f.dictionaries.Create(False)
    If Not filters Is Nothing Then
        For Each key In filters.Keys
            filter = f.regex.getFirstGroupMatch(VBA.CStr(key), FORMULA_FILTER_PATTERN)
            Call f.dictionaries.addItem(adjustedFilters, filter, filters.item(key))
        Next key
    End If
    
    
    For Each wks In pSheets
        Set rng = wks.Range(wks.cells(1, 1), wks.cells.SpecialCells(xlCellTypeLastCell))
        If f.ranges.IsValid(rng) Then formulas = rng.FormulaR1C1
        
        If f.arrays.isDefined(formulas) Then
            For i = LBound(formulas, 1) To UBound(formulas, 1)
                For j = LBound(formulas, 2) To UBound(formulas, 2)
                    formula = formulas(i, j)
                    
                    If VBA.Left$(formula, 1) <> "=" Then
                        ' not formula, skip this cell
                    ElseIf isFormulaMatchedWithFilters(formula, adjustedFilters) Then
                        formula = getFormulaShortNotation(formula, adjustedFilters)
                        groupKey = wks.name & SHEET_FORMULA_SEPARATOR & formula
                        
                        Set groupCells = f.dictionaries.getObject(getCellGroupsByFormula, groupKey)
                        If groupCells Is Nothing Then
                            Set groupCells = New VBA.Collection
                            Call getCellGroupsByFormula.Add(groupKey, groupCells)
                        End If
                        Call groupCells.Add(wks.cells(i, j))
                    End If
                    
                Next j
            Next i
        End If
        
    Next wks
    
    
    
    
End Function


Private Function isFormulaMatchedWithFilters(formula As String, filters As Scripting.Dictionary) As Boolean
    Dim filter As Variant
    '------------------------------------------------------------------------------------------------------
    
    If filter.count = 0 Then
        isFormulaMatchedWithFilters = True
    Else
        For Each filter In filters.Keys
            If VBA.InStr(1, formula, filter, vbTextCompare) Then
                isFormulaMatchedWithFilters = True
                Exit For
            End If
        Next filter
    End If
    
End Function


Private Function getFormulaShortNotation(formula As String, dictLinks As Scripting.Dictionary) As String
    Const REPLACEMENT_REGEX As String = "(?:[A-Z]:|\\\\)\\\\[^']*\\\\\[{0}\]"
    '------------------------------------------------------------------------------------------------------
    Dim key As Variant
    Dim searchText As String
    '------------------------------------------------------------------------------------------------------
    
    getFormulaShortNotation = formula
    
    For Each key In dictLinks.Keys
        searchText = f.strings.format(REPLACEMENT_REGEX, key)
        getFormulaShortNotation = f.regex.replace(getFormulaShortNotation, searchText, dictLinks.item(key))
    Next key
    
End Function

Private Sub replacePathInFormulas(groups As Scripting.Dictionary, dictKeyToNewPath As Scripting.Dictionary)
    Dim key As Variant
    Dim newFormula As String
    Dim cells As VBA.Collection
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    For Each key In groups.Keys
        Set cells = groups.item(key)
        newFormula = getNewFormula(VBA.CStr(key), dictKeyToNewPath)
        Set rng = f.ranges.join(cells)
        rng.FormulaR1C1 = newFormula
    Next key
    
End Sub

Private Function getNewFormula(formula As String, dictKeyToNewPath As Scripting.Dictionary) As String
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    getNewFormula = f.strings.substring(formula, SHEET_FORMULA_SEPARATOR, VBA.vbNullString, False)
    For Each key In dictKeyToNewPath.Keys
        getNewFormula = VBA.replace(getNewFormula, key, dictKeyToNewPath.item(key))
    Next key
End Function
