VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepoGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "RepoGenerator"
'----------------------------------------------------------------------------------------------------------
Private Const EN_REPO_NAME As String = "RANGE_REPO_NAME"
Private Const EN_SAVE_BUTTON As String = "RANGE_SAVE_BUTTON"
Private Const PROPERTIES_CLASS_NAME As String = "CProperties"
'[Excel names] --------------------------------------------------------------------------------------------
Private Const EN_IS_REPO_GENERATOR_TECH_BOOK As String = "IsRepoGeneratorTechBook"
'[Captions]
Private Const EDIT_SHEET_NAME As String = "VBE.AddRepo.EditSheetName"
'[Files]
Private Const TEMPLATE_ENTITY_FILE_NAME As String = "EntityClass.txt"
Private Const TEMPLATE_REPO_FILE_NAME As String = "RepoClass.txt"
Private Const TEMPLATE_SETTERS_GETTERS_PATTERNS As String = "SetterGetterPatterns.txt"
Private Const TEMPLATE_MSG_FILE As String = "RepoMsg.txt"
Private Const TEMPLATE_CONFIG_FILE As String = "RepoConfig.txt"
Private Const TEMPLATE_DB_SERVICE As String = "RepoDbService.txt"
'[Tags]
Private Const REMOVE_TAG_ALIASES As String = "aliases"
Private Const REMOVE_TAG_EDITABLE_NAME As String = "hasEditableName"
Private Const REMOVE_TAG_HAS_DATE_PROPERTY As String = "hasDateProperty"
'----------------------------------------------------------------------------------------------------------
Private pPatterns As Scripting.Dictionary
'----------------------------------------------------------------------------------------------------------
Private pInitProject As VBIDE.VBProject
Private pCoreProject As VBIDE.VBProject
Private pDbProject As VBIDE.VBProject
Private pObjectsProject As VBIDE.VBProject
'----------------------------------------------------------------------------------------------------------
Private pEditBook As Excel.Workbook
Private WithEvents pEditSheet As Excel.Worksheet
Attribute pEditSheet.VB_VarHelpID = -1
Private pDefinition As ERepoDefinition
'----------------------------------------------------------------------------------------------------------


Private Sub Class_Initialize()
    Call loadSetterGetterPatterns
End Sub

Private Sub loadSetterGetterPatterns()
    'Const REGEX_PATTERN As String = "\[([\w:]*)\]\s((?:.|\s)*?End (?:Function|Sub))"
    Const REGEX_PATTERN As String = "\[([\w:]*)\]\s((?:.|\s)*?)\[\/\1"
    '------------------------------------------------------------------------------------------------------
    Dim filePath As String
    Dim content As String
    Dim matches As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    Dim match As Variant
    Dim tag As String
    Dim pattern As String
    '------------------------------------------------------------------------------------------------------
    
    Set pPatterns = F.dictionaries.Create(False)
    
    filePath = Paths.getTemplateFilePath(TEMPLATE_SETTERS_GETTERS_PATTERNS)
    If F.Files.FileExists(filePath) Then
        content = F.TextFiles.readTextFile(filePath)
        Set matches = F.regex.getMatchArraysCollection(content, REGEX_PATTERN)
        For Each match In matches
            tag = match(LBound(match))
            pattern = match(LBound(match) + 1)
            Call F.dictionaries.addItem(pPatterns, tag, pattern)
        Next match
    End If
    
End Sub



Public Sub run()
    Set pInitProject = Fn.getActiveProject
    Call setupVbProjects(pInitProject)
    Call setupEditBook
End Sub

Private Sub setupEditBook()
    Dim wkb As Excel.Workbook
    Dim wks As Excel.Worksheet
    '------------------------------------------------------------------------------------------------------
    
    If F.Books.IsValid(pEditBook) Then
        ' no action needed
    Else
        For Each wkb In Excel.Workbooks
            If F.ExcelNames.getValue(wkb, EN_IS_REPO_GENERATOR_TECH_BOOK) Then
                Set pEditBook = wkb
                Exit For
            End If
        Next wkb
        
        If pEditBook Is Nothing Then
            Set pEditBook = F.Books.addNew(1, Excel.Application, TEMP_SHEET_NAME & "_")
            Call F.ExcelNames.assignValue(pEditBook, EN_IS_REPO_GENERATOR_TECH_BOOK, 1, True)
        End If
        
    End If
    
    
    '[Prepare sheets]
    Set wks = pEditBook.Worksheets.Add
    wks.name = F.Sheets.getUniqueName(pEditBook, TEMP_SHEET_NAME)
    
    Call F.Sheets.deleteAllExcept(pEditBook, TEMP_SHEET_NAME, False)
    
    Call wksRepoTemplate.Copy(Before:=pEditBook.Worksheets(1))
    Set pEditSheet = pEditBook.Worksheets(1)
    pEditSheet.name = msg.getText(EDIT_SHEET_NAME)
    
    Call F.Sheets.Delete(pEditBook, TEMP_SHEET_NAME)
    
End Sub

Private Sub setupVbProjects(project As VBIDE.VBProject)
    Const REGEX_PATTERN As String = "^(\w+?)(?:_|$)"
    Const CORE_FILE_NAME_PATTERN As String = "{0}_Core"
    Const OBJECTS_FILE_NAME_PATTERN As String = "{0}_Objects"
    Const DB_FILE_NAME_PATTERN As String = "{0}_Db"
    '------------------------------------------------------------------------------------------------------
    Dim projectPrefix As String
    Dim coreName As String
    '------------------------------------------------------------------------------------------------------
    projectPrefix = F.regex.getFirstGroupMatch(project.name, REGEX_PATTERN)
    Set pCoreProject = F.Developer.getVbProjectByProjectName(F.Strings.Format(CORE_FILE_NAME_PATTERN, projectPrefix))
    Set pObjectsProject = F.Developer.getVbProjectByProjectName(F.Strings.Format(CORE_FILE_NAME_PATTERN, projectPrefix))
    Set pDbProject = F.Developer.getVbProjectByProjectName(F.Strings.Format(DB_FILE_NAME_PATTERN, projectPrefix))
End Sub





Public Function getPattern(tag As String) As String
    getPattern = F.dictionaries.GetItem(pPatterns, tag)
End Function

Private Function getProjectFolderPath() As String
    If Not pInitProject Is Nothing Then
        getProjectFolderPath = F.Files.getFileFolderPath(pInitProject.Filename)
    End If
End Function


Private Sub createRepo()
    Call setupForInsert
    Call setupRepoDefinition
    
    Call insertDataTypes
    Call insertProps
    Call insertEntityClass
    Call insertRepoClass
    Call insertConstructor
    Call insertMsgFile
    Call insertConfigFile
    Call insertDbService
    
End Sub

Private Sub setupForInsert()
    Const REGEX_PATTERN As String = "^(\w+?)(?:_|$)"
    Const CORE_FILE_NAME_PATTERN As String = "{0}_Core"
    Const OBJECTS_FILE_NAME_PATTERN As String = "{0}_Objects"
    Const DB_FILE_NAME_PATTERN As String = "{0}_Db"
    '------------------------------------------------------------------------------------------------------
    Dim projectPrefix As String
    Dim coreName As String
    Dim objectsName As String
    Dim dbName As String
    '------------------------------------------------------------------------------------------------------
    
    projectPrefix = F.regex.getFirstGroupMatch(pInitProject.name, REGEX_PATTERN)
    coreName = F.Strings.Format(CORE_FILE_NAME_PATTERN, projectPrefix)
    objectsName = F.Strings.Format(OBJECTS_FILE_NAME_PATTERN, projectPrefix)
    dbName = F.Strings.Format(DB_FILE_NAME_PATTERN, projectPrefix)
    
    Set pCoreProject = F.Developer.getVbProjectByProjectName(coreName)
    Set pObjectsProject = F.Developer.getVbProjectByProjectName(objectsName)
    Set pDbProject = F.Developer.getVbProjectByProjectName(dbName)
    
End Sub

Private Sub setupRepoDefinition()
    Set pDefinition = New ERepoDefinition
    Call pDefinition.readFromSheet(pEditSheet)
End Sub

Private Sub pEditSheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    Set rng = F.Sheets.getNamedRange(pEditSheet, EN_SAVE_BUTTON)
    If Target.Address = rng.Address Then
        Call createRepo
        Cancel = True
    End If
    
End Sub




Private Sub insertDataTypes()
    Dim col As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    
    Set col = New VBA.Collection
    
    With pDefinition
        Call col.Add(VBA.Array(.getName, .getReadTable, .getWriteTable))
        If .isEntityAliasable Then
            Call col.Add(VBA.Array(.getAliasName, .getAliasReadTable, .getAliasWriteTable))
        End If
    End With
    
    If F.Collections.isNonEmptyCollection(col) Then
        Call DataTypesGenerator.setDisplayAlerts(False).addDataTypes(pCoreProject, col, pDefinition.getName)
    End If
    
End Sub



Private Sub insertProps()
    Dim dictProps As Scripting.Dictionary
    Dim key As Variant
    Dim propName As String
    Dim dbField As String
    '------------------------------------------------------------------------------------------------------
    Dim colInputArrays As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    
    Set dictProps = pDefinition.getPropsDictionary
    Set colInputArrays = New VBA.Collection
    
    For Each key In dictProps.Keys
        propName = key
        dbField = dictProps.item(key)
        If F.Strings.compareStrings(propName, dbField) Then dbField = Empty
        Call colInputArrays.Add(VBA.Array(propName, dbField))
    Next key
    
    If F.Collections.isNonEmptyCollection(colInputArrays) Then
        Call PropsGenerator. _
                        setDisplayAlerts(False). _
                        addProps(pCoreProject, colInputArrays, pDefinition.getName)
    End If

End Sub


Private Sub insertEntityClass()
    Dim templatePath As String
    Dim template As String
    Dim replacements As Scripting.Dictionary
    Dim content As String
    '------------------------------------------------------------------------------------------------------
    Dim cl As VBIDE.VBComponent
    '------------------------------------------------------------------------------------------------------
    
    templatePath = Paths.getTemplateFilePath(TEMPLATE_ENTITY_FILE_NAME)
    template = removeRedundantSections(F.TextFiles.readTextFile(templatePath))
    
    Set replacements = createReplacementsDictionary_EntityClass
    content = F.Strings.formatByTags(template, replacements)
            
    If Not pObjectsProject Is Nothing Then
        Set cl = pObjectsProject.VBComponents.Add(vbext_ct_ClassModule)
        Call F.Developer.changeClassInstancing(cl, True)
        cl.name = pDefinition.getEntityClassName
        Call cl.CodeModule.AddFromString(content)
    Else
        Debug.Print "[!!!] Error while trying to insert entity class - [project]_objects.xlam file not found"
    End If
            
End Sub

Private Function createReplacementsDictionary_EntityClass() As Scripting.Dictionary
    Dim entityName As String
    '------------------------------------------------------------------------------------------------------

    entityName = pDefinition.getEntityName

    Set createReplacementsDictionary_EntityClass = F.dictionaries.Create(False)
    With createReplacementsDictionary_EntityClass
        Call .Add("ClassName", "E" & F.Strings.convertLetterCasing(entityName, LetterCasing_Sentence))
        Call .Add("EntityNameCapitalLetters", VBA.UCase$(entityName))
        Call .Add("EntityNameLowerCase", F.Strings.convertLetterCasing(entityName, LetterCasing_StartWithLower))
        Call .Add("EntityNameSentenceCase", F.Strings.convertLetterCasing(entityName, LetterCasing_Sentence))
        Call .Add("NameProperty", pDefinition.getNameProperty)
        Call .Add("DataTypeName", pDefinition.getName)
        Call .Add("RepoIdField", pDefinition.getIdField)
        Call .Add("ClassPropertiesDeclarations", getClassPropertiesDeclaration)
        Call .Add("PropertiesSetters", getPropertiesSetters)
        Call .Add("PropertiesDtoLoading", getPropertiesDtoLoading)
        Call .Add("SetValueMethodCases", getSetValueMethodCases)
        Call .Add("PropertiesGetters", getPropertiesGetters)
        Call .Add("GetValueMethodCases", getGetValueMethodCases)
        Call .Add("CloneProperties", getCloneProperties)
        Call .Add("InsertDtoEntries", getInsertDtoEntries)
        Call .Add("EditLinesDefinitionsDeclarations", getEditLinesDefinitionsDeclarations)
        Call .Add("EditLinesDefinitionsAssignment", getEditLinesDefinitionsAssignment)
        Call .Add("EditLinesDefinitions", getEditLinesDefinitions)
        Call .Add("ValidationFunctions", getValidationFunctions)
    End With
    
End Function

Private Function removeRedundantSections(template As String) As String
    Const REMOVE_REGEX_PATTERN As String = "\n?<{0}>(?:.|\n)*?<\/{0}>\n?"
    Const KEEP_REGEX_PATTERN As String = "\n?[ \t]*<\/?{0}>[ \t]*\n?"
    '------------------------------------------------------------------------------------------------------
    Dim regexPattern As String
    '------------------------------------------------------------------------------------------------------
    
    removeRedundantSections = template
    
    '[Aliases]
    regexPattern = VBA.Replace( _
                        VBA.IIf(pDefinition.isEntityAliasable, KEEP_REGEX_PATTERN, REMOVE_REGEX_PATTERN), _
                        "{0}", REMOVE_TAG_ALIASES)
    removeRedundantSections = F.regex.Replace(removeRedundantSections, regexPattern, VBA.vbNullString)
    
    
    '[Name property editable]
    regexPattern = VBA.Replace( _
                        VBA.IIf(pDefinition.hasNameProperty, KEEP_REGEX_PATTERN, REMOVE_REGEX_PATTERN), _
                        "{0}", REMOVE_TAG_EDITABLE_NAME)
    removeRedundantSections = F.regex.Replace(removeRedundantSections, regexPattern, VBA.vbNullString)
    
    
    '[Date property]
    regexPattern = VBA.Replace( _
                        VBA.IIf(pDefinition.hasDateProperty, KEEP_REGEX_PATTERN, REMOVE_REGEX_PATTERN), _
                        "{0}", REMOVE_TAG_HAS_DATE_PROPERTY)
    removeRedundantSections = F.regex.Replace(removeRedundantSections, regexPattern, VBA.vbNullString)

End Function

Private Function getClassPropertiesDeclaration() As String
    Dim lines As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    Set lines = F.Collections.Map(pDefinition.getProps, "$.getDeclarationLine")
    getClassPropertiesDeclaration = F.Collections.toString(lines, , VBA.vbCrLf)
End Function

Private Function getPropertiesSetters() As String
    Dim lines As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    Set lines = F.Collections.Map(pDefinition.getProps, "$.getSetterMethod")
    getPropertiesSetters = F.Collections.toString(lines, , VBA.vbCrLf)
End Function

Private Function getPropertiesDtoLoading() As String
    Dim lines As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    Set lines = F.Collections.Map(pDefinition.getProps, "$.getDtoLoadingLine")
    getPropertiesDtoLoading = F.Collections.toString(lines, , VBA.vbCrLf)
End Function

Private Function getSetValueMethodCases() As String
    Dim lines As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    Set lines = F.Collections.Map(pDefinition.getProps, "$.getSetValueMethodCases")
    getSetValueMethodCases = F.Collections.toString(lines, , VBA.vbCrLf)
End Function

Private Function getPropertiesGetters() As String
    Dim lines As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    Set lines = F.Collections.Map(pDefinition.getProps, "$.getGetterMethod")
    getPropertiesGetters = F.Collections.toString(lines, , VBA.vbCrLf)
End Function

Private Function getGetValueMethodCases() As String
    Dim lines As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    Set lines = F.Collections.Map(pDefinition.getProps, "$.getGetValueMethodCases")
    getGetValueMethodCases = F.Collections.toString(lines, , VBA.vbCrLf)
End Function

Private Function getCloneProperties() As String
    Dim lines As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    Set lines = F.Collections.Map(pDefinition.getProps, "$.getCloneLine")
    getCloneProperties = F.Collections.toString(lines, , VBA.vbCrLf)
End Function

Private Function getInsertDtoEntries() As String
    Dim props As VBA.Collection
    Dim lines As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    Set props = F.Collections.Filter(pDefinition.getProps, "$.isInsertedToDb")
    Set lines = F.Collections.Map(props, "$.getInsertDtoLine")
    getInsertDtoEntries = F.Collections.toString(lines, , VBA.vbCrLf)
End Function

Private Function getEditLinesDefinitionsDeclarations() As String
    Dim props As VBA.Collection
    Dim lines As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    Set props = F.Collections.Filter(pDefinition.getProps, "$.isLinkedToRepo")
    If F.Collections.isNonEmptyCollection(props) Then
        Set lines = F.Collections.Map(props, "$.getEditLinesDefinitionsDeclaration")
        getEditLinesDefinitionsDeclarations = F.Collections.toString(lines, , VBA.vbCrLf) & _
                                                VBA.vbCrLf & VBA.vbTab & _
                                                "'" & VBA.String(METHOD_LEVEL_SEPARATOR_LENGTH, "-")
    End If
End Function

Private Function getEditLinesDefinitionsAssignment() As String
    Dim props As VBA.Collection
    Dim lines As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    Set props = F.Collections.Filter(pDefinition.getProps, "$.isLinkedToRepo")
    Set lines = F.Collections.Map(props, "$.getEditLinesDefinitionsAssignment")
    getEditLinesDefinitionsAssignment = F.Collections.toString(lines, , VBA.vbCrLf)
End Function

Private Function getEditLinesDefinitions() As String
    Dim props As VBA.Collection
    Dim lines As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    Set props = F.Collections.Filter(pDefinition.getProps, "$.isEditable")
    Set lines = F.Collections.Map(props, "$.getEditLineDefinition")
    getEditLinesDefinitions = F.Collections.toString(lines, , VBA.vbCrLf)
End Function

Private Function getValidationFunctions() As String
    Dim dict As Scripting.Dictionary
    Dim prop As EProperty
    '------------------------------------------------------------------------------------------------------
    Set dict = F.dictionaries.Create(False)
    For Each prop In pDefinition.getProps
        Call prop.appendValidationFunction(dict)
    Next prop
    getValidationFunctions = F.dictionaries.toString(dict, DictPart_ValueOnly, StringifyMode_Normal, VBA.vbCrLf)
End Function








Private Sub insertRepoClass()
    Dim templatePath As String
    Dim template As String
    Dim replacements As Scripting.Dictionary
    Dim content As String
    '------------------------------------------------------------------------------------------------------
    Dim cl As VBIDE.VBComponent
    '------------------------------------------------------------------------------------------------------
    
    templatePath = Paths.getTemplateFilePath(TEMPLATE_REPO_FILE_NAME)
    template = removeRedundantSections(F.TextFiles.readTextFile(templatePath))
    
    Set replacements = createReplacementsDictionary_RepoClass
    content = F.Strings.formatByTags(template, replacements)

    If Not pObjectsProject Is Nothing Then
        Set cl = pObjectsProject.VBComponents.Add(vbext_ct_ClassModule)
        Call F.Developer.changeClassInstancing(cl, True)
        cl.name = pDefinition.getRepoClassName
        Call cl.CodeModule.AddFromString(content)
    Else
        Debug.Print "[!!!] Error while trying to insert repo class - [project]_objects.xlam file not found"
    End If

End Sub

Private Function createReplacementsDictionary_RepoClass() As Scripting.Dictionary
    Dim repoName As String
    Dim entityName As String
    '------------------------------------------------------------------------------------------------------

    repoName = pDefinition.getName
    entityName = pDefinition.getEntityName

    Set createReplacementsDictionary_RepoClass = F.dictionaries.Create(False)
    With createReplacementsDictionary_RepoClass
        Call .Add("RepoName", pDefinition.getName)
        Call .Add("RepoNameSentenceCase", F.Strings.convertLetterCasing(pDefinition.getName, LetterCasing_Sentence))
        Call .Add("RepoNameLowerCase", F.Strings.convertLetterCasing(pDefinition.getName, LetterCasing_StartWithLower))
        Call .Add("EntityNameLowerCase", F.Strings.convertLetterCasing(entityName, LetterCasing_StartWithLower))
        Call .Add("EntityNameSentenceCase", F.Strings.convertLetterCasing(entityName, LetterCasing_Sentence))
    End With
    
End Function



Private Sub insertConstructor()
    Dim pattern As String
    Dim methodBody As String
    Dim component As VBIDE.VBComponent
    '------------------------------------------------------------------------------------------------------
    
    pattern = getPattern("constructor")
    methodBody = F.Strings.Format(pattern, F.Strings.convertLetterCasing(pDefinition.getName, LetterCasing_Sentence))
    
    If Not pObjectsProject Is Nothing Then
        Set component = F.Developer.getVbComponent(pObjectsProject, "Entities")
    End If
    
    If component Is Nothing Then
        Debug.Print "[!!!] Error while trying to add repo class constructor - class Entities not found in [project]_objects.xlam"
    Else
        Call component.CodeModule.AddFromString(VBA.vbCrLf & methodBody)
    End If
    
End Sub


Private Sub insertMsgFile()
    Dim filePath As String
    Dim template As String
    Dim content As String
    Dim replacements As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Dim destinationFilePath As String
    '------------------------------------------------------------------------------------------------------
    
    filePath = Paths.getTemplateFilePath(TEMPLATE_MSG_FILE)
    If F.Files.FileExists(filePath) Then
        template = removeRedundantSections(F.TextFiles.readTextFile(filePath))
    End If
    
    If VBA.Len(template) Then
        Set replacements = F.dictionaries.Create(False)
        With replacements
            Call .Add("REPO_NAME", F.Strings.convertLetterCasing(pDefinition.getName, LetterCasing_Sentence))
            Call .Add("REPO_NAME_LOW", F.Strings.convertLetterCasing(pDefinition.getName, LetterCasing_Lower))
            Call .Add("ENTITY_NAME", F.Strings.convertLetterCasing(pDefinition.getEntityName, LetterCasing_Sentence))
            Call .Add("ENTITY_NAME_SENTENCE", F.Strings.convertLetterCasing(pDefinition.getEntityName, LetterCasing_Sentence))
            Call .Add("ENTITY_NAME_LOW", F.Strings.convertLetterCasing(pDefinition.getEntityName, LetterCasing_Lower))
        End With
        content = VBA.Replace(F.Strings.formatByTags(template, replacements), NEW_LINE_REPLACEMENT_TAG, NEW_LINE_SYMBOL)
    End If
    
    If VBA.Len(content) Then
        destinationFilePath = F.Files.concatenatePath(getProjectFolderPath, _
                                    Paths.FolderNameResources, Paths.FolderNameMsgs, Paths.FolderNameProject, _
                                    F.Strings.convertLetterCasing(pDefinition.getName, LetterCasing_Lower) & EXTENSION_JSON)
        Call F.TextFiles.printToTextFile(content, destinationFilePath, False)
    End If
    
    
End Sub

Private Sub insertConfigFile()
    Dim filePath As String
    Dim template As String
    Dim content As String
    Dim replacements As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Dim destinationFilePath As String
    '------------------------------------------------------------------------------------------------------
    
    filePath = Paths.getTemplateFilePath(TEMPLATE_CONFIG_FILE)
    If F.Files.FileExists(filePath) Then
        template = removeRedundantSections(F.TextFiles.readTextFile(filePath))
    End If
    
    If VBA.Len(template) Then
        Set replacements = F.dictionaries.Create(False)
        With replacements
            Call .Add("REPO_NAME", F.Strings.convertLetterCasing(pDefinition.getName, LetterCasing_Sentence))
            Call .Add("ENTITY_NAME", F.Strings.convertLetterCasing(pDefinition.getEntityName, LetterCasing_Sentence))
        End With
        
        content = F.Strings.formatByTags(template, replacements)
        
    End If
    
    
    
    If VBA.Len(content) Then
        destinationFilePath = F.Files.concatenatePath(getProjectFolderPath, _
                                    Paths.FolderNameResources, Paths.FolderNameConfig, "objects", _
                                    F.Strings.convertLetterCasing(pDefinition.getName, LetterCasing_Lower) & EXTENSION_JSON)
        Call F.TextFiles.printToTextFile(content, destinationFilePath, False)
    End If
    
    
End Sub

Private Sub insertDbService()
    Dim filePath As String
    Dim template As String
    Dim content As String
    Dim replacements As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Dim dbClass As VBIDE.VBComponent
    Dim serviceClass As VBIDE.VBComponent
    '------------------------------------------------------------------------------------------------------
    
    filePath = Paths.getTemplateFilePath(TEMPLATE_DB_SERVICE)
    If F.Files.FileExists(filePath) Then
        template = removeRedundantSections(F.TextFiles.readTextFile(filePath))
    End If
    
    If VBA.Len(template) Then
        Set replacements = F.dictionaries.Create(False)
        With replacements
            Call .Add("REPO_NAME", F.Strings.convertLetterCasing(pDefinition.getName, LetterCasing_Sentence))
            Call .Add("ENTITY_NAME", F.Strings.convertLetterCasing(pDefinition.getEntityName, LetterCasing_Sentence))
        End With
        
        content = F.Strings.formatByTags(template, replacements)
        
    End If
    
    
    If VBA.Len(content) = 0 Then
        ' empty - skip
    ElseIf pDbProject Is Nothing Then
        ' DB file not found
    Else
    
        Set serviceClass = pDbProject.VBComponents.Add(vbext_ct_ClassModule)
        Call F.Developer.changeClassInstancing(serviceClass, True)
        With serviceClass
            .name = "Service_" & F.Strings.convertLetterCasing(pDefinition.getName, LetterCasing_Sentence)
            .CodeModule.AddFromString (content)
        End With
    
        Set dbClass = F.Developer.getVbComponent(pDbProject, "DbRepository")
        If Not dbClass Is Nothing Then
            Call dbClass.CodeModule.AddFromString(VBA.vbCrLf & getDbServiceCreateMethod)
        End If
        
    End If
    
End Sub

Private Function getDbServiceCreateMethod() As String
    Dim pattern As String
    '------------------------------------------------------------------------------------------------------
    pattern = getPattern("DbServiceConstructor")
    getDbServiceCreateMethod = F.Strings.Format(pattern, _
                                    F.Strings.convertLetterCasing(pDefinition.getName, LetterCasing_Sentence))
End Function
