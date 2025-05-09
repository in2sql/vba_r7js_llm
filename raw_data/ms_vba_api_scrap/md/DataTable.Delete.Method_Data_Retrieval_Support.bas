Attribute VB_Name = "Data_Retrieval_Support"

Option Explicit
Public Enum SocrataStatus
    NoNewData = 1
    Failure = 2
    NewDataQueried = 3
End Enum
Public Sub Retrieve_Historical_Workbooks(ByRef Path_CLCTN As Collection, ByVal ICE_Contracts As Boolean, ByVal CFTC_Contracts As Boolean, _
                                               ByVal Mac_User As Boolean, _
                                               ByVal eReport As ReportEnum, _
                                               ByVal downloadFuturesAndOptions As Boolean, _
                                            Optional ByVal CFTC_Start_Date As Date, _
                                            Optional ByVal CFTC_End_Date As Date, _
                                            Optional ByVal ICE_Start_Date As Date, _
                                            Optional ByVal ICE_End_Date As Date, _
                                            Optional ByVal Historical_Archive_Download As Boolean = False)
'===================================================================================================================
    'Summary: Downloads CFTC .zip files.
    'Inputs: Path_CLTCN - Collection to store file paths to extracted CoT data.
    '        ICE_Contracts - True if ICE data should be downloaded.
    '        CFTC_Contracts - True if CFTC data should be downloaded.
    '        Mac_User - True if script is being run on a MAC.
    '        reportType - Type of report to download.
    '        downloadFuturesAndOptions - True if futures + options should be retrieved else futures only.
    '        CFTC_Start_Date - Min cftc date.
    '        CFTC_End_Date - Max cftc date.
    '        Historical_Archive_Download - If true then download all data available.
'===================================================================================================================
    Dim fileNameWithinZip$, Path_Separator$, AnnualOF_FilePath$, Destination_Folder$, zipFileNameAndPath$, _
    fullFileName$, multiYearFileExtractedFromZip$, Partial_Url$, url$, multiYearZipFileFullName$, combinedOrFutures$, Multi_Year_URL$
    
    Dim Queried_Date As Long, Download_Year As Long, Final_Year As Long, multiYearName$, reportInitial$
    
    Const TXT$ = ".txt", ZIP$ = ".zip", CSV$ = ".csv"
    
    Const mainFolderName$ = "COT_Historical_MoshiM"
    
    On Error GoTo Failed_To_Download
    
    reportInitial = ConvertReportTypeEnum(eReport)
    
    #If Not Mac Then
        
        Path_Separator = Application.PathSeparator
        
        Destination_Folder = Environ$("TEMP") & Path_Separator & mainFolderName & Path_Separator & reportInitial & Path_Separator & IIf(downloadFuturesAndOptions = True, "Combined", "Futures Only")
        
        If Not (FileOrFolderExists(Destination_Folder) Or ICE_Contracts) Then
            CreateFolderRecursive Destination_Folder
        End If
        
    #Else
        '/Users/rondebruin/Library/Containers/com.microsoft.Excel/Data

'        This setion is for if files are downloaded and stored on client computer.
'        As of May 2024 MAc users only need this sub for getting urls to ice data.
'        Path_Separator = "/"
'        Destination_Folder = BasicMacAvailablePathMac & Path_Separator & mainFolderName & Path_Separator & IIf(downloadFuturesAndOptions = True, "Combined", "Futures Only") 'Keep variable as an empty string.User will decide path
'        If Not FileOrFolderExists(Destination_Folder) Then
'            Call CreateRootDirectories(Destination_Folder)
'        End If
        
    #End If
    
    With Path_CLCTN
    
        #If Not Mac Then
        
            If CFTC_Contracts Then
            
                If Not downloadFuturesAndOptions Then  'IF Futures Only Workbook
                
                    combinedOrFutures = "_Futures_Only"
                    
                    Select Case eReport
                        Case eLegacy
                            fileNameWithinZip = "annual" & TXT
                            Partial_Url = "https://www.cftc.gov/files/dea/history/deacot"
                            Multi_Year_URL = "https://www.cftc.gov/files/dea/history/deacot1986_2016" & ZIP
                            multiYearName = "FUT86_16"
                        Case eDisaggregated
                            fileNameWithinZip = "f_year" & TXT
                            Partial_Url = "https://www.cftc.gov/files/dea/history/fut_disagg_txt_"
                            Multi_Year_URL = "https://www.cftc.gov/files/dea/history/fut_disagg_txt_hist_2006_2016" & ZIP
                            multiYearName = "F_DisAgg06_16"
                        Case eTFF
                            fileNameWithinZip = "FinFutYY" & TXT
                            Partial_Url = "https://www.cftc.gov/files/dea/history/fut_fin_txt_"
                            Multi_Year_URL = "https://www.cftc.gov/files/dea/history/fin_fut_txt_2006_2016" & ZIP
                            multiYearName = "F_TFF_2006_2016"
                    End Select
                
                Else 'Combined Contracts
                
                    combinedOrFutures = "_Combined"
                    
                    Select Case eReport
                        Case eLegacy
                            fileNameWithinZip = "annualof.txt"
                            Partial_Url = "https://www.cftc.gov/files/dea/history/deahistfo" 'TXT URL
                            Multi_Year_URL = "https://www.cftc.gov/files/dea/history/deahistfo_1995_2016" & ZIP
                            multiYearName = "Com95_16"
                        Case eDisaggregated
                            fileNameWithinZip = "c_year" & TXT
                            Partial_Url = "https://www.cftc.gov/files/dea/history/com_disagg_txt_"
                            'https://www.cftc.gov/files/dea/history/com_disagg_txt_hist_2006_2016.zip
                            Multi_Year_URL = "https://www.cftc.gov/files/dea/history/com_disagg_txt_hist_2006_2016" & ZIP
                            multiYearName = "C_DisAgg06_16"
                        Case eTFF
                            fileNameWithinZip = "FinComYY" & TXT
                            'https://www.cftc.gov/files/dea/history/com_fin_txt_2014.zip
                            Partial_Url = "https://www.cftc.gov/files/dea/history/com_fin_txt_"
                            Multi_Year_URL = "https://www.cftc.gov/files/dea/history/fin_com_txt_2006_2016" & ZIP
                            multiYearName = "C_TFF_2006_2016"
                    End Select
                
                End If
                
                If Year(CFTC_Start_Date) <= 2016 Then 'All report types have a compiled file for data before 2016
                    Historical_Archive_Download = True
                    CFTC_Start_Date = DateSerial(2017, 1, 1) 'So we can start dates in 2017 instead
                End If
                
                multiYearZipFileFullName = Destination_Folder & Path_Separator & reportInitial & "_COT_MultiYear_Archive" & combinedOrFutures & ZIP
                
                AnnualOF_FilePath = Destination_Folder & Path_Separator & fileNameWithinZip
        
                Download_Year = Year(CFTC_Start_Date)
                Final_Year = Year(CFTC_End_Date)
                Queried_Date = CFTC_End_Date
                
                '-1 is for if historical archive download needs to be executed
                For Download_Year = Download_Year - 1 To Final_Year
                        
                    If Not Historical_Archive_Download Then 'if not doing a download where multi year files are needed ie 2006-2016
                    
                        If Download_Year = Year(CFTC_Start_Date) - 1 Then
                            GoTo Skip_Download_Loop 'if on first loop
                        Else
                            url = Partial_Url & Download_Year & ZIP 'Declare URL of Zip file
                        End If
                        
                    ElseIf Historical_Archive_Download Then
                        url = Multi_Year_URL
                    End If

                    If Historical_Archive_Download Then
                        fullFileName = Destination_Folder & Path_Separator & reportInitial & "_" & multiYearName & combinedOrFutures & TXT
                    ElseIf Final_Year = Download_Year Then
                        fullFileName = Destination_Folder & Path_Separator & reportInitial & "_Weekly_" & CLng(Queried_Date) & "_" & Download_Year & combinedOrFutures & TXT
                    Else
                        fullFileName = Destination_Folder & Path_Separator & reportInitial & "_" & Download_Year & combinedOrFutures & TXT
                    End If
                    
                    If Not FileOrFolderExists(fullFileName) Then   'If wanted workbook doesn't exist
                        
                        If Historical_Archive_Download Then
                            zipFileNameAndPath = multiYearZipFileFullName
                        Else
                            zipFileNameAndPath = Replace$(fullFileName, TXT, ZIP)
                        End If
                        
                        If Not FileOrFolderExists(zipFileNameAndPath) Then
                            #If Mac Then
                                Call DownloadFileMAC(url, zipFileNameAndPath)
                            #Else
                                Call DownloadFile(url, zipFileNameAndPath)
                            #End If
                        End If

                        If Not Historical_Archive_Download Then
                        
                            If FileOrFolderExists(AnnualOF_FilePath) Then Kill AnnualOF_FilePath    'If file within Zip folder exists within file directory then kill it
                        
                            #If Mac Then
                                Call UnzipFile(zipFileNameAndPath, Destination_Folder, fileNameWithinZip)
                            #Else
                                Call entUnZip1File(zipFileNameAndPath, Destination_Folder, fileNameWithinZip) 'Unzip specified file
                            #End If
                            
                            Name AnnualOF_FilePath As fullFileName
                            
                        ElseIf Historical_Archive_Download Then
                        
                            multiYearFileExtractedFromZip = Destination_Folder & Path_Separator & multiYearName & TXT
                            
                            If FileOrFolderExists(multiYearFileExtractedFromZip) Then Kill multiYearFileExtractedFromZip
    
                            #If Mac Then
                                Call UnzipFile(zipFileNameAndPath, Destination_Folder, multiYearName & TXT)
                            #Else
                                Call entUnZip1File(zipFileNameAndPath, Destination_Folder, multiYearName & TXT) 'Unzip specified file
                            #End If
                            
                            Name multiYearFileExtractedFromZip As fullFileName
                            
                        End If
                            
                    End If
                    
                    .Add fullFileName, fullFileName
        
Skip_Download_Loop:
                    Historical_Archive_Download = False
        
                Next Download_Year
                
            End If
        
        #End If
        
        If ICE_Contracts Then
            
            If Year(ICE_Start_Date) < 2011 Then
                ICE_Start_Date = #1/1/2011#
            End If
            
            Final_Year = Year(ICE_End_Date)
            
            For Download_Year = Year(ICE_Start_Date) To Final_Year
            
                url = "https://www.theice.com/publicdocs/futures/COTHist" & Download_Year & ".csv"
                
                Select Case Download_Year
                    Case Final_Year
                        fullFileName = Destination_Folder & Path_Separator & "ICE_Weekly_" & CLng(ICE_End_Date) & "_" & Download_Year & ".csv"
                    Case Else
                        fullFileName = Destination_Folder & Path_Separator & "ICE_" & Download_Year & ".csv"
                End Select
                
                .Add url, url
    
            Next Download_Year
            
        End If
        
    End With
    
    Exit Sub
    
Failed_To_Download:
    Call PropagateError(Err, "Retrieve_Historical_Workbooks")
End Sub
Public Function IsWorkbookOutdated(Optional workbookPath$) As Boolean

'===================================================================================================================
    'Summary: Tests if a given file has been updated with the most recent data available.
    'Inputs: workbookPath - File path  of file to test.
    'Returns: True if data doesn't need updating; else, false.
'===================================================================================================================
    Dim Last_Release As Date

    On Error GoTo Default_True
    
    Last_Release = CFTC_Release_Dates(True, True) 'Returns Local date and time for the most recent release
    
    If LenB(workbookPath) <> 0 And CDbl(Last_Release) <> 0 Then
        IsWorkbookOutdated = (FileDateTime(workbookPath) < Last_Release)
    Else
       IsWorkbookOutdated = True
    End If
    
    Exit Function
    
Default_True:
    IsWorkbookOutdated = True
    
End Function

Public Function HTTP_Weekly_Data(previousUpdateDate As Date, reportType As ReportEnum, retrieveCombinedData As Boolean, ByRef useApi As Boolean, ByRef columnMap As Collection, Optional suppressMessages As Boolean = False, _
                                Optional testAllMethods As Boolean = False, Optional DebugActive As Boolean = False) As Variant
'===================================================================================================================
    'Summary: Uses multiple methods of data retrieval from the CFTC.
    'Inputs: previousUpdateDate - Date converted to long for which data was last updated to.
    '        reportType - One of L,D,T to represent what type of report to retrieve.
    '        retrieveCombinedData - true if futures + options data should be retrieved; else, futures only data will be retrieved.
    '        useApi - If true then the function will attempt to retrieve data via API.
    '        suppressMessages - true if error messages should be repressed.
    '        columnMap - An empty collection that wil store FieldInfo instances for each column found within the output.
    'Returns: An array of weekly data if ap method fails; else, all data since last_update.
'===================================================================================================================
    Dim PowerQuery_Available As Boolean, Power_Query_Failed As Boolean, _
    Text_Method_Failed As Boolean, Query_Table_Method_Failed As Boolean, _
    MAC_OS As Boolean, dataRetrieved As Boolean, successCount As Long, tempData() As Variant, attemptCount As Long
    
    Dim retrievalTimer As TimedTask, savedState As Boolean, apiStatusCode As SocrataStatus
        
    Const PowerQTask$ = "Power Query Retrieval", _
    QueryTask$ = "QueryTable Retrieval", HTTPTask$ = "HTTP Retrieval", _
    ApiTask = "Socrata API", ProcedureName = "HTTP_Weekly_Data"
    
    #If Mac Then
        MAC_OS = True
        PowerQuery_Available = False 'Use standalone QueryTable rather than QueryTable wrapped in listobject
    #Else
        PowerQuery_Available = IsPowerQueryAvailable()
    #End If
    
Retrieval_Process:
    If testAllMethods Then
        Set retrievalTimer = New TimedTask
        retrievalTimer.Start "Time Retrieval Methods."
    End If
    
    savedState = ThisWorkbook.Saved
    
    If useApi Then
        If testAllMethods Then
            If MsgBox("Test Socrata API Method", vbYesNo) <> vbYes Then GoTo QueryTable_Method
            attemptCount = attemptCount + 1
            retrievalTimer.StartSubTask ApiTask
        End If

        On Error GoTo Catch_SocrataRetrievalFailed

        If TryGetCftcWithSocrataAPI(tempData, reportType, retrieveCombinedData, apiStatusCode, debugModeActive:=(testAllMethods Or DebugActive), fieldInfoByEditedName:=columnMap, greaterThanDate:=previousUpdateDate) Then
            On Error GoTo 0
            If IsArrayAllocated(tempData) Then
                HTTP_Weekly_Data = tempData
                Erase tempData
                dataRetrieved = True
            End If
        ElseIf apiStatusCode = SocrataStatus.NoNewData Then
            If Not testAllMethods Then On Error GoTo 0
            Err.Raise RetrievalErr.SocrataSuccessNoNewData, ProcedureName, "No new data could be retrieved from Socrata's API."
        End If
        
        If testAllMethods Then
            retrievalTimer.StopSubTask ApiTask
            successCount = successCount + 1
        End If
    End If
    
QueryTable_Method:
    If dataRetrieved = False Or testAllMethods Then
        If testAllMethods Then
            If MsgBox("Test Querytable Method", vbYesNo) <> vbYes Then GoTo PowerQuery_Method
            attemptCount = attemptCount + 1
            retrievalTimer.StartSubTask QueryTask
        End If
        
        On Error GoTo QueryTable_Failed
            
        HTTP_Weekly_Data = CFTC_Data_QueryTable_Method(reportType:=reportType, retrieveCombinedData:=retrieveCombinedData)
        
        If testAllMethods Then
            retrievalTimer.StopSubTask QueryTask
            successCount = successCount + 1
        End If
        
        dataRetrieved = True
    End If
    
PowerQuery_Method:

    If Not MAC_OS Then
        If (Not dataRetrieved And PowerQuery_Available) Or testAllMethods Then
            If testAllMethods Then
                If MsgBox("Test PowerQuery Method", vbYesNo) <> vbYes Then GoTo TXT_Method
                attemptCount = attemptCount + 1
                retrievalTimer.StartSubTask PowerQTask
            End If
            
            On Error GoTo PowerQuery_Failed
                
            HTTP_Weekly_Data = CFTC_Data_PowerQuery_Method(reportType, retrieveCombinedData)
                
            If testAllMethods Then
                retrievalTimer.StopSubTask PowerQTask
                successCount = successCount + 1
            End If
            dataRetrieved = True
        End If
TXT_Method:
        If testAllMethods Or Not dataRetrieved Then
            If testAllMethods Then
                If MsgBox("Test Txt Method", vbYesNo) <> vbYes Then GoTo Finally
                attemptCount = attemptCount + 1
                retrievalTimer.StartSubTask HTTPTask
            End If
            
            On Error GoTo TXT_Failed
                
            HTTP_Weekly_Data = CFTC_Data_Text_Method(previousUpdateDate, reportType:=reportType, retrieveCombinedData:=retrieveCombinedData)
                
            If testAllMethods Then
                retrievalTimer.StopSubTask HTTPTask
                successCount = successCount + 1
            End If
        
            dataRetrieved = True
        End If
    End If
Finally:
    On Error GoTo Catch_GeneralError
    
    ThisWorkbook.Saved = savedState
    
    If testAllMethods Then retrievalTimer.DPrint
    
    If dataRetrieved And columnMap Is Nothing Then
        Set columnMap = GetExpectedLocalFieldInfo(reportType, False, False, False, False)
    End If
    
    On Error GoTo 0
    If Not dataRetrieved Then
        Err.Raise RetrievalErr.RetrievalFailed, ProcedureName, "All retrieval methods have failed."
    ElseIf testAllMethods And successCount <> attemptCount Then
        Err.Raise RetrievalErr.RetrievalFailed, ProcedureName, successCount & " of " & attemptCount & " retrieval methods have failed."
    End If
    
    Exit Function

Catch_GeneralError:
    PropagateError Err, ProcedureName
PowerQuery_Failed:

    If testAllMethods Then
        DisplayErr Err, ProcedureName
        retrievalTimer.StopSubTask PowerQTask
    End If
    Resume TXT_Method
    
TXT_Failed:

    If testAllMethods Then
        DisplayErr Err, ProcedureName
        retrievalTimer.StopSubTask HTTPTask
    End If
    Resume Finally
    
QueryTable_Failed:

    If testAllMethods Then
        DisplayErr Err, ProcedureName
        retrievalTimer.StopSubTask QueryTask
    End If
    
    If Not MAC_OS Then
        Resume PowerQuery_Method
    Else
        Resume Finally
    End If

Catch_SocrataRetrievalFailed:

    If testAllMethods Then
        DisplayErr Err, ProcedureName
        retrievalTimer.StopSubTask ApiTask
    End If
    
    useApi = False
    Resume QueryTable_Method

End Function

Private Function SocrataRetrievalPowerQuery(eReport As ReportEnum, getFuturesAndOptions As Boolean, statusCode As SocrataStatus, apiUrl$, queryReturnLimit&, _
        Optional debugModeActive As Boolean = False, _
        Optional executionTimer As TimedTask) As Collection

    Dim loopCount As Long

    On Error GoTo Finally

    Dim savedState As Boolean, enableTimers As Boolean, queryTimer As TimedTask, processTimer As TimedTask, _
    eventState As Boolean, workbookQueryAssignmentTimer As TimedTask

    With Application
        eventState = .EnableEvents: .EnableEvents = False
    End With

    enableTimers = Not executionTimer Is Nothing

    If enableTimers Then
        With executionTimer
            Set workbookQueryAssignmentTimer = .SubTask("Create PowerQuery Objects.")
            Set queryTimer = .SubTask("Query Socrata with PowerQuery.")
            Set processTimer = .SubTask("Gather Data.")
        End With
    End If

    Const queryName$ = "Socrata API"
    
    ' Using Object allows you to avoid compilation errors for older version of excel.
    Dim apiQuery As Object, apiUrlParam As Object, _
    wb As Workbook, mCode$, socrataDataTable As ListObject, QT As QueryTable, editQueryProperties As Boolean
    
    ' Excel won't flag compilation errors if using a workbook object rather than ThisWorkbook.
    Set wb = ThisWorkbook: savedState = ThisWorkbook.Saved
    
    If enableTimers Then workbookQueryAssignmentTimer.Start
    On Error GoTo Catch_Queries_Unavailable
    With wb.Queries
    
        On Error Resume Next
        Set apiUrlParam = .item("Socrata_API_URL")
        If Err.Number <> 0 Then
           mCode = """" & apiUrl & """ meta [IsParameterQuery=true, Type=""Text"", IsParameterQueryRequired=true]"
            Set apiUrlParam = .Add("Socrata_API_URL", mCode, "Socrata API URL parameter.")
            Err.Clear
        End If
        
        Set apiQuery = .item(queryName)
        If Err.Number <> 0 Then
            mCode = "let Source = Csv.Document(Web.Contents(Socrata_API_URL),[Delimiter="","", Encoding=65001, QuoteStyle=QuoteStyle.None])," & _
                vbNewLine & Space$(4) & "#""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])," & _
                vbNewLine & Space$(4) & "#""Removed Columns"" = Table.Buffer(Table.RemoveColumns(#""Promoted Headers"",{""id"", ""yyyy_report_week_ww"", ""contract_market_name"", ""cftc_region_code"", ""cftc_commodity_code"", ""commodity_name"", ""commodity"", ""commodity_subgroup_name"", ""commodity_group_name"", ""futonly_or_combined"", ""cftc_subgroup_code""}, MissingField.Ignore))," & _
                vbNewLine & Space$(4) & "Table_Headers = List.Buffer(Table.ColumnNames(#""Removed Columns""))," & _
                vbNewLine & Space$(4) & "Numeric_KeyWords = List.Buffer({""open"", ""trader"", ""pos"", ""pct"", ""conc"",""change_in"",""yyyy""})," & _
                vbNewLine & Space$(4) & "IsNumeric = (name as text) => List.MatchesAny(Numeric_KeyWords, each Text.Contains(name ,_))," & _
                vbNewLine & Space$(4) & "Numeric_Fields = List.Select(Table_Headers, each IsNumeric(_))," & _
                vbNewLine & Space$(4) & "columnTransformations = List.Transform(Numeric_Fields, each {_,  if Text.StartsWith(_,""conc"") or Text.StartsWith(_,""pct"") then Number.FromText else if Text.StartsWith(_,""report"") then DateTime.FromText else Int32.From})," & _
                vbNewLine & Space$(4) & "TransformedTable = Table.TransformColumns(#""Removed Columns"", columnTransformations)," & _
                vbNewLine & Space$(4) & "#""Replaced Value"" = Table.ReplaceValue(TransformedTable,0,null,Replacer.ReplaceValue,Table_Headers)" & _
            vbNewLine & "in" & _
                vbNewLine & Space$(4) & "#""Replaced Value"""

            Set apiQuery = .Add(queryName, mCode, "Queries a Socrata API.")
            Err.Clear
        End If
        
    End With
    
    On Error GoTo Finally
    
    With QueryT
        On Error Resume Next
        Set socrataDataTable = .ListObjects(queryName)
        
        If Err.Number <> 0 Then
            On Error GoTo Finally
            Set QT = .ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location= """ & queryName & """;Extended Properties=""""" _
                , Destination:=.Range("$N$56")).QueryTable
                editQueryProperties = True
        Else
            On Error GoTo Finally
            Set QT = socrataDataTable.QueryTable
        End If
        
    End With
    
    If editQueryProperties Then
        With QT
            .CommandType = xlCmdSql
            .CommandText = Array("SELECT * FROM [" & queryName & "]")
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = False
            .AdjustColumnWidth = False
            .RefreshPeriod = 0
            .PreserveColumnInfo = True
        End With
    End If
    
    If enableTimers Then workbookQueryAssignmentTimer.EndTask
    
    loopCount = 0

    Do
        loopCount = loopCount + 1
        
        apiUrlParam.Formula = """" & apiUrl & "&$offset=" & queryReturnLimit * (loopCount - 1) & """ meta [IsParameterQuery=true, Type=""Text"", IsParameterQueryRequired=true]"

        Application.StatusBar = ConvertReportTypeEnum(eReport) & IIf(getFuturesAndOptions, "_Combined", "_Futures_Only") & " : Querying API for records {" & loopCount & "}"

        If enableTimers Then queryTimer.Start

        On Error GoTo Catch_RefreshFailure
        QT.Refresh False
        On Error GoTo Finally

        If enableTimers Then queryTimer.Pause

        Application.StatusBar = vbNullString

        If socrataDataTable Is Nothing Then
            Set socrataDataTable = QT.ListObject
            socrataDataTable.name = queryName
        End If

        Dim returnedRecordsCount&, collectedData As Collection

        With socrataDataTable

            returnedRecordsCount = .ListRows.Count

            If returnedRecordsCount > 0 Then

                If loopCount = 1 Then
                    statusCode = SocrataStatus.NewDataQueried
                    Set collectedData = New Collection
                    ' Get a 1D array of column names.
                    collectedData.Add Application.Transpose(Application.Transpose(.HeaderRowRange.Value2)), "Headers"
                End If

                If enableTimers Then processTimer.Start
                collectedData.Add .DataBodyRange.Value2
                If enableTimers Then processTimer.Pause

            ElseIf loopCount = 1 Then
                ' Query successfully completed but no data was returned.
                statusCode = SocrataStatus.NoNewData
            End If

        End With

    Loop While returnedRecordsCount = queryReturnLimit And Not debugModeActive

    Set SocrataRetrievalPowerQuery = collectedData
Finally:
    'If Not socrataDataTable Is Nothing Then socrataDataTable.Delete

    With Application
        .EnableEvents = eventState: .StatusBar = vbNullString
    End With

    ThisWorkbook.Saved = savedState

    If Err.Number <> 0 Then
        statusCode = SocrataStatus.Failure
        Call PropagateError(Err, "SocrataRetrievalPowerQuery")
    End If

    Exit Function
Catch_RefreshFailure:
    AppendErrorDescription Err, "An error occurred while attempting to connect to the Socrata API for [ " & ConvertReportTypeEnum(eReport) & " ] getFuturesAndOptions=" & getFuturesAndOptions & "."
    GoTo Finally
Catch_Queries_Unavailable:
    AppendErrorDescription Err, "Workbook.Queries object unavailable."
    GoTo Finally
End Function
Private Function SocrataRetrievalQueryTable(eReport As ReportEnum, getFuturesAndOptions As Boolean, statusCode As SocrataStatus, apiUrl$, queryReturnLimit&, _
        Optional debugModeActive As Boolean = False, _
        Optional executionTimer As TimedTask) As Collection
        
        Dim columnTypes(1 To 200) As XlColumnDataType, socrataQueryTable As QueryTable, dateColumn&, codeColumn&, returnedRows&
        
        Dim tempDataCLCTN As Collection, loopCount&, enableTimers As Boolean, queryTimer As TimedTask, gatherDataTimer As TimedTask
        
        On Error GoTo Finally
        
        enableTimers = Not executionTimer Is Nothing
    
        If enableTimers Then
            With executionTimer
                Set queryTimer = .SubTask("Query Socrata with QueryTable.")
                Set gatherDataTimer = .SubTask("Gather Data.")
            End With
        End If
    
        dateColumn = 3: codeColumn = 6
        
        ' General purpose array that will work for all Report types. Unneeded values will be discarded.
        For loopCount = LBound(columnTypes) To UBound(columnTypes)
            Select Case loopCount
                Case 1, 4, 5, 8, 9, 10
                    'id, yyyy_report_week_ww, contract_market_name, cftc_region_code, cftc_commodity_code, commodity_name
                    columnTypes(loopCount) = xlSkipColumn
                Case dateColumn, codeColumn
                    columnTypes(loopCount) = xlTextFormat
                Case Else
                    columnTypes(loopCount) = xlGeneralFormat
            End Select
        Next loopCount
    
        With QueryT
            Set socrataQueryTable = .QueryTables.Add(Connection:="TEXT;" & apiUrl, Destination:=.Range("A1"))
        End With

        With socrataQueryTable

            loopCount = 0
            Do ' Loop until the API doesn't return anything.
                loopCount = loopCount + 1
                
                If loopCount > 1 Then .Connection = "TEXT;" & apiUrl & IIf(loopCount > 1, "&$offset=" & queryReturnLimit * (loopCount - 1), vbNullString)
                ' Adjusting the .Connection property in loopCount > 1 will wipe some properties.
                .TextFileCommaDelimiter = True
                .BackgroundQuery = False
                .SaveData = False
                .AdjustColumnWidth = False
                .PreserveFormatting = False
                .RefreshOnFileOpen = False
                .RefreshStyle = xlOverwriteCells
                .TextFileTextQualifier = xlTextQualifierDoubleQuote
                .TextFileCommaDelimiter = True
                .TextFileColumnDataTypes = columnTypes
                
                Application.StatusBar = ConvertReportTypeEnum(eReport) & IIf(getFuturesAndOptions, "_Combined", "_Futures_Only") & " : Querying API for records {" & loopCount & "}"
                
                If enableTimers Then queryTimer.Start
                On Error GoTo Catch_RefreshFailure
                .Refresh False
                On Error GoTo Finally
                If enableTimers Then queryTimer.Pause
                
                Application.StatusBar = vbNullString
                
                With .ResultRange
                
                    returnedRows = .Rows.Count - 1

                    If returnedRows > 0 Then
                        If loopCount = 1 Then
                            statusCode = SocrataStatus.NewDataQueried
                            If tempDataCLCTN Is Nothing Then Set tempDataCLCTN = New Collection
                            ' Get a 1D array of column names.
                            tempDataCLCTN.Add Application.Transpose(Application.Transpose(.Rows(1).Value2)), "Headers"
                        End If
                        
                        If enableTimers Then gatherDataTimer.Start
                        tempDataCLCTN.Add .Offset(1).Resize(returnedRows).Value2
                        If enableTimers Then gatherDataTimer.Pause
                    ElseIf loopCount = 1 Then
                        ' Query successfully completed but no data was returned.
                        statusCode = SocrataStatus.NoNewData
                    End If
                    
                End With
    
            Loop While returnedRows = queryReturnLimit And (Not debugModeActive Or (debugModeActive And loopCount < 2))
            
            QueryT.UsedRange.ClearContents
            .WorkbookConnection.Delete
            .Delete
            Erase columnTypes
            Set socrataQueryTable = Nothing
        End With
        Set SocrataRetrievalQueryTable = tempDataCLCTN
Finally:
    If Not socrataQueryTable Is Nothing Then
        With socrataQueryTable
            If Not .WorkbookConnection Is Nothing Then .WorkbookConnection.Delete
            .Delete
        End With
    End If
    
    Application.StatusBar = vbNullString
    
    If Err.Number <> 0 Then
        statusCode = SocrataStatus.Failure
        Call PropagateError(Err, "SocrataRetrievalQueryTable")
    End If
    
    Exit Function
Catch_RefreshFailure:
    AppendErrorDescription Err, "An error occurred while attempting to connect to the Socrata API for [ " & ConvertReportTypeEnum(eReport) & " ] getFuturesAndOptions=" & getFuturesAndOptions & "."
    GoTo Finally
End Function
Private Sub SocrataTest()
    Dim h() As Variant, tt As New TimedTask, statCode As SocrataStatus, eReport As ReportEnum
    
    On Error GoTo Display
    
    Const testdate As Date = #1/13/2024#: eReport = eLegacy
    
    With tt
        .Start Now & vbNewLine & "Query API [" & ConvertReportTypeEnum(eReport) & " - " & Format$(testdate, "yyyy-mm-dd]")
        #If DatabaseFile Then
            'TryGetCftcWithSocrataAPI h, eReport, True, statCode, greaterThanDate:=testdate, executionTimer:=tt, allowPowerQuery:=True
            TryGetCftcWithSocrataAPI h, eReport, True, statCode, greaterThanDate:=testdate, executionTimer:=tt, allowPowerQuery:=False, debugModeActive:=True
        #Else
            TryGetCftcWithSocrataAPI h, ConvertInitialToReportTypeEnum(ReturnReportType()), True, statCode, greaterThanDate:=testdate, executionTimer:=tt
        #End If
        .EndTask
        .DPrint
    End With
    Exit Sub
Display:
    DisplayErr Err, "SocrataTest"
End Sub
Public Function TryGetCftcWithSocrataAPI(ByRef outputA() As Variant, eReport As ReportEnum, getFuturesAndOptions As Boolean, statusCode As SocrataStatus, _
        Optional debugModeActive As Boolean = False, _
        Optional ByRef fieldInfoByEditedName As Collection, _
        Optional contractCode$ = vbNullString, _
        Optional ByRef greaterThanDate As Date = #1/1/1970#, _
        Optional executionTimer As TimedTask, _
        Optional allowPowerQuery As Boolean = False) As Boolean
    '===================================================================================================================
    'Summary: Retrieve data from the CFTC's Public Reporting Environment API.
    'Inputs:
    '        outputA - Array that will store retrieved data if successfull.
    '        greaterThanDate - Date which data was last updated to.
    '        eReport - One of L,D,T to represent what type of report to retrieve.
    '        getFuturesAndOptionsData - true if futures + options data should be retrieved; else, futures only data will be retrieved.
    '        contractCode - If supplied with a value than only data that with this contract code will be retrieved.
    '        fieldInfoByEditedName - Empty Collection that will store information for wanted fields.
    'Output: True if data was successfully retrieved.
    '===================================================================================================================

    Dim socrataDataCollection As Collection, apiUrl$, queryReturnLimit As Long, socrataData() As Variant, imperfectOperator$
    
    On Error GoTo Finally

    If LenB(contractCode) <> 0 Then contractCode = " AND cftc_contract_market_code='" & contractCode & "'"

    queryReturnLimit = IIf(debugModeActive, 400, 40000)
    imperfectOperator = IIf(debugModeActive, ">=", ">")
                    
    apiUrl = "https://publicreporting.cftc.gov/resource/" & GetSocrataApiEndpoint(eReport, CInt(getFuturesAndOptions)) & ".csv" & _
                "?$where=report_date_as_yyyy_mm_dd " & imperfectOperator & Format$(greaterThanDate, "'yyyy-mm-ddT00:00:00.000'") & _
                contractCode & "&$order=report_date_as_yyyy_mm_dd,id&$limit=" & queryReturnLimit

    Dim basicField As FieldInfo, columnInOutput As Long, columnInApiData As Long, savedState As Boolean, enableTimers As Boolean, _
    queryTimer As TimedTask, assignmentTimer As TimedTask, gatherDataTimer As TimedTask, eventState As Boolean
    
    savedState = ThisWorkbook.Saved
    
    With Application
        eventState = .EnableEvents: .EnableEvents = False
    End With
    
    enableTimers = Not executionTimer Is Nothing

    Set fieldInfoByEditedName = Nothing
    
    #Const UseHTTP = False
        
    #If Not Mac And UseHTTP Then
        
        If enableTimers Then Set queryTimer = executionTimer.SubTask("Query Socrata with GET request.")
        
        Dim apiResponse$, returnedRecordsA() As String, singleRecordA() As String, cftcRegionCodeColumn As Long, _
        iRow&, numberOfRecordsReturned&, loopCount As Long
        
        Const Comma$ = ",", Period$ = "."
        
        loopCount = 0
        
        Do
            loopCount = loopCount + 1
            
            Application.StatusBar = ConvertReportTypeEnum(eReport) & IIf(getFuturesAndOptions, "_Combined", "_Futures_Only") & " : Querying API for records {" & loopCount & "}"
            
            If enableTimers Then queryTimer.Start
            
            If TryGetRequest(apiUrl & IIf(loopCount > 1, "&$offset=" & queryReturnLimit * (loopCount - 1), vbNullString), apiResponse) Then
                
                If enableTimers Then queryTimer.Pause
            
                ' Splitting by vbLf will return an array with headers as the first element and a null string as the final element.
                ' Array will consist of Header;Data;Terminating Line Feed
                returnedRecordsA = Split(apiResponse, vbLf)
                apiResponse = vbNullString
                'Number of data rows = Ubound(returnedRecordsA) + (- 2 + 1)
                numberOfRecordsReturned = UBound(returnedRecordsA) - 1

                If numberOfRecordsReturned > 0 Then
                    
                    If enableTimers Then gatherDataTimer.Start
                    
                    For iRow = LBound(returnedRecordsA) To UBound(returnedRecordsA)
                        
                        If LenB(returnedRecordsA(iRow)) <> 0 Then
                                                        
                            ' Split on commas outside of quotes
                            singleRecordA = SplitOutsideOfQuotes(returnedRecordsA(iRow), Comma)
                            
                            If iRow = LBound(returnedRecordsA) Then
                                If loopCount = 1 Then
                                    statusCode = SocrataStatus.NewDataQueried
                                    ' Create collection of FieldInfo instances based on API headers.
                                    Set fieldInfoByEditedName = CreateFieldInfoMap(externalHeaders:=singleRecordA, _
                                                                    localDatabaseHeaders:=Application.Transpose(GetAvailableFieldsTable(eReport).DataBodyRange.columns(1).Value2), _
                                                                    externalHeadersFromSocrataAPI:=True)
                                End If

                                With fieldInfoByEditedName
                                    If loopCount = 1 Then cftcRegionCodeColumn = .item("cftc_region_code").ColumnIndex
                                    ReDim outputA(1 To numberOfRecordsReturned, 1 To .Count)
                                End With
                            Else
                                If enableTimers Then assignmentTimer.Start
                                columnInOutput = LBound(outputA, 2)
                                For Each basicField In fieldInfoByEditedName
                                    With basicField
                                        If Not .IsMissing Then
                                            ' This is Base 0
                                            columnInApiData = .ColumnIndex
                                            
                                            If singleRecordA(columnInApiData) = Period Then
                                                outputA(iRow, columnInOutput) = Empty
                                            ElseIf Not (columnInApiData = cftcRegionCodeColumn Or LenB(singleRecordA(columnInApiData)) = 0) Then
                                                Select Case .DataType
                                                    Case 7, 133
                                                        'adDate, adDbDate
                                                        outputA(iRow, columnInOutput) = CDate(Left$(singleRecordA(columnInApiData), 10))
                                                    Case 131
                                                        ' adNumeric
                                                        outputA(iRow, columnInOutput) = CDbl(singleRecordA(columnInApiData))
                                                    Case 3, 5
                                                        'adInteger, adDouble... In the original database files integer fields are double fields.
                                                        outputA(iRow, columnInOutput) = CLng(singleRecordA(columnInApiData))
                                                    Case 202, 200
                                                        'adVarWChar, adVarChar
                                                        outputA(iRow, columnInOutput) = Trim$(singleRecordA(columnInApiData))
                                                End Select
                                            End If
                                        End If
                                        columnInOutput = columnInOutput + 1
                                    End With
                                Next basicField
                                If enableTimers Then assignmentTimer.Pause
                            End If
                        End If
                    Next iRow
                    ' Save data to a collection if needed for compilation after loops.
                    If loopCount > 1 Or numberOfRecordsReturned = queryReturnLimit Then
                        If socrataDataCollection Is Nothing Then Set socrataDataCollection = New Collection
                        socrataDataCollection.Add outputA
                    End If
                    If enableTimers Then gatherDataTimer.Pause
                    
                ElseIf loopCount = 1 Then
                    ' Query successfully completed but no data was returned.
                     statusCode = SocrataStatus.NoNewData
                End If
                
            Else
                If enableTimers Then queryTimer.EndTask
                ' Failed to retrieve data via HTTP
                statusCode = SocrataStatus.Failure
                'err.Raise vbObjectError + 1000, Description:="Failed to Query with current url."
            End If
        Loop While numberOfRecordsReturned = queryReturnLimit And statusCode <> SocrataStatus.Failure And Not debugModeActive
                
        If statusCode = SocrataStatus.NewDataQueried Then
            If Not socrataDataCollection Is Nothing Then
                Select Case socrataDataCollection.Count
                    Case 1
                        'Exactly queryReturnLimit retrieved and is already stored in outputA
                    Case Is > 1
                        If enableTimers Then assignmentTimer.Start
                        outputA = CombineArraysInCollection(socrataDataCollection, Append_Type.Multiple_2d)
                        If enableTimers Then assignmentTimer.EndTask
                End Select
            End If
            
            If Not fieldInfoByEditedName Is Nothing Then
                ' Adjust indexes so that they are base 1.
                columnInOutput = LBound(outputA, 2)
                For Each basicField In fieldInfoByEditedName
                    basicField.ColumnIndex = columnInOutput
                    columnInOutput = columnInOutput + 1
                Next basicField
            Else
                Err.Raise vbObjectError + 799, Description:="fieldInfoByEditedName is nothing despite data being found."
            End If
        End If
        
    #Else
    
        If IsPowerQueryAvailable() And IsCreatorActiveUser() And allowPowerQuery Then
            Set socrataDataCollection = SocrataRetrievalPowerQuery(eReport, getFuturesAndOptions, statusCode, apiUrl, queryReturnLimit, debugModeActive, executionTimer)
        Else
            Set socrataDataCollection = SocrataRetrievalQueryTable(eReport, getFuturesAndOptions, statusCode, apiUrl, queryReturnLimit, debugModeActive, executionTimer)
        End If
        
        If Not socrataDataCollection Is Nothing And statusCode = SocrataStatus.NewDataQueried Then
            
            On Error GoTo Finally
            
            Dim socrataColumnNames() As Variant, combinerTimer As TimedTask
            
            With socrataDataCollection
                socrataColumnNames = .item("Headers"): .Remove "Headers"
            End With
            
            Erase socrataData
            
            Select Case socrataDataCollection.Count
                Case 1
                    socrataData = socrataDataCollection(1)
                Case Is > 1
                    If enableTimers Then Set combinerTimer = executionTimer.StartSubTask("Combine collected arrays.")
                    socrataData = CombineArraysInCollection(socrataDataCollection, Append_Type.Multiple_2d)
                    If enableTimers Then combinerTimer.EndTask
                Case Else
                    Err.Raise vbObjectError + 1002, Description:="'socrataDataCollection' has no items."
            End Select
            
            Set socrataDataCollection = Nothing
            
            If IsArrayAllocated(socrataData) Then
                
                Dim codeColumn&, dateColumn&, iCount&, wantedFieldsA() As Variant
                Const Period$ = "."
                
                ' Get an ordered list of wanted fields.
                wantedFieldsA = Application.Transpose(GetAvailableFieldsTable(eReport).DataBodyRange.columns(1).Value2)
                
                Set fieldInfoByEditedName = CreateFieldInfoMap(socrataColumnNames, wantedFieldsA, externalHeadersFromSocrataAPI:=True)
                
                Erase socrataColumnNames: Erase wantedFieldsA
                                
                With fieldInfoByEditedName
                    If .Count > 0 Then
                        ReDim outputA(LBound(socrataData, 1) To UBound(socrataData, 1), 1 To .Count)
                        codeColumn = .item("cftc_contract_market_code").ColumnIndex
                        dateColumn = .item("report_date_as_yyyy_mm_dd").ColumnIndex
                    Else
                        Err.Raise vbObjectError + 1003, Description:="CreateFieldInfoMap() didn't return any FieldInfo instances."
                    End If
                End With
                
                columnInOutput = LBound(socrataData, 2) - 1
                
                If enableTimers Then Set assignmentTimer = executionTimer.StartSubTask("Assign array elements.")
                
                On Error GoTo Catch_ElementAssignmentError
                
                For Each basicField In fieldInfoByEditedName
                    columnInOutput = columnInOutput + 1
                    With basicField
                        If Not .IsMissing Then
                            columnInApiData = .ColumnIndex
                            For iCount = LBound(socrataData, 1) To UBound(socrataData, 1)
                                If Not IsError(socrataData(iCount, columnInApiData)) Then
                                    Select Case columnInApiData
                                        Case codeColumn
                                            If Len(socrataData(iCount, columnInApiData)) <> 6 Then
                                                outputA(iCount, columnInOutput) = Format$(socrataData(iCount, columnInApiData), "000000")
                                            Else
                                                outputA(iCount, columnInOutput) = socrataData(iCount, columnInApiData)
                                            End If
                                        Case dateColumn
                                            Select Case VarType(socrataData(iCount, columnInApiData))
                                                Case vbDate
                                                    outputA(iCount, columnInOutput) = socrataData(iCount, columnInApiData)
                                                Case vbDouble, vbLong
                                                    outputA(iCount, columnInOutput) = CDate(socrataData(iCount, columnInApiData))
                                                Case vbString
                                                    outputA(iCount, columnInOutput) = CDate(Left$(socrataData(iCount, columnInApiData), 10))
                                            End Select
                                        Case Else
                                            If VarType(socrataData(iCount, columnInApiData)) = vbString Then
                                                If socrataData(iCount, columnInApiData) <> Period Then outputA(iCount, columnInOutput) = Trim$(socrataData(iCount, columnInApiData))
                                            ElseIf socrataData(iCount, columnInApiData) <> 0 Then
                                                outputA(iCount, columnInOutput) = socrataData(iCount, columnInApiData)
                                            End If
                                    End Select
                                End If
                            Next iCount
                        End If
                        ' The field reflects column within the api data. Adjust it to match column in outputA.
                        .ColumnIndex = columnInOutput
                    End With
                Next basicField
                
                If enableTimers Then assignmentTimer.EndTask
            Else
                Err.Raise vbObjectError + 1001, Description:="Variable 'socrataData' isn't initialized when it should be."
            End If
        End If
    #End If
    
    TryGetCftcWithSocrataAPI = (statusCode = SocrataStatus.NewDataQueried)
    
Finally:
    
    With Application
        .EnableEvents = eventState: .StatusBar = vbNullString
    End With
    
    ThisWorkbook.Saved = savedState
    
    If Err.Number <> 0 Then
        'Stop: Resume
        Erase outputA
        statusCode = SocrataStatus.Failure
        Call PropagateError(Err, "TryGetCftcWithSocrataAPI")
    End If
    
    Exit Function
Catch_ElementAssignmentError:
    AppendErrorDescription Err, "Error while attempting to fill output array."
    GoTo Finally
End Function

Public Function CFTC_Data_PowerQuery_Method(reportType As ReportEnum, retrieveCombinedData As Boolean) As Variant()
'===================================================================================================================
    'Summary: Retrieves the latest Weekly data with Power Query.
    'Inputs: reportType - One of L,D,T to represent what type of report to retrieve.
    '        retrieveCombinedData - true if futures + options data should be retrieved; else, futures only data will be retrieved.
    'Returns: An array of the most recent weekly CFTC data.
    'Notes: Use only on Windows.
'===================================================================================================================
    On Error GoTo Failure
        
    #If DatabaseFile Then
        
        Dim url$, Formula_AR$(), quotation$, Y As Long, table_name$, wb As Workbook
        
        quotation = Chr(34)
        
        url = "https://www.cftc.gov/dea/newcot/"
        
        Y = Application.Match(reportType, Array(eLegacy, eDisaggregated, eTFF), 0) - 1
        
        If Not retrieveCombinedData Then 'Futures Only
            url = url & Array("deafut.txt", "f_disagg.txt", "FinFutWk.txt")(Y)
        Else
            url = url & Array("deacom.txt", "c_disagg.txt", "FinComWk.txt")(Y)
        End If
        table_name = Split("Legacy,Disaggregated,TFF", ",")(Y)
        
        'Change Query URL
        Set wb = ThisWorkbook
        
        With wb.Queries(table_name)
            Formula_AR = Split(.Formula, quotation, 3)
            Formula_AR(1) = url
            .Formula = Join(Formula_AR, quotation)
        End With
    
        With Weekly.ListObjects(table_name)
            .QueryTable.Refresh False                               'Refresh Weekly Data Table
            CFTC_Data_PowerQuery_Method = .DataBodyRange.Value2     'Store contents of table in an array
        End With
    
    #Else
        With Weekly.ListObjects("Weekly").QueryTable
            .Refresh False
            CFTC_Data_PowerQuery_Method = .ResultRange.Value2
        End With
    #End If
    
    Exit Function
Failure:
    PropagateError Err, "CFTC_Data_PowerQuery_Method"
End Function

Public Function CFTC_Data_Text_Method(Last_Update As Date, reportType As ReportEnum, retrieveCombinedData As Boolean) As Variant()
'===================================================================================================================
    'Summary: Retrieves the latest Weekly using HTTP methods found on the Windows version of Excel.
    'Inputs: reportType - One of L,D,T to represent what type of report to retrieve.
    '        retrieveCombinedData - true if futures + options data should be retrieved; else, futures only data will be retrieved.
    '        Last_Update - Date that data was last retrieved for.
    'Returns: An array of the most recent weekly CFTC data.
    'Notes: Use only on Windows.
'===================================================================================================================
    Dim filePath$, url$, Y As Long
    
    On Error GoTo Failure
    url = "https://www.cftc.gov/dea/newcot/"
    
    Y = Application.Match(reportType, Array(eLegacy, eDisaggregated, eTFF), 0) - 1
    
    If Not retrieveCombinedData Then 'Futures Only
        url = url & Array("deafut.txt", "f_disagg.txt", "FinFutWk.txt")(Y)
    Else
        url = url & Array("deacom.txt", "c_disagg.txt", "FinComWk.txt")(Y)
    End If
    
    filePath = Environ$("TEMP") & "\" & Date & "_" & ConvertReportTypeEnum(reportType) & "_Weekly.txt"

    Call DownloadFile(url, filePath)
    
    CFTC_Data_Text_Method = Weekly_Text_File(filePath, reportType:=reportType, retrieveCombinedData:=retrieveCombinedData)
    
    Exit Function
Failure:
    PropagateError Err, "CFTC_Data_Text_Method"
End Function
Public Function CFTC_Data_QueryTable_Method(reportType As ReportEnum, retrieveCombinedData As Boolean) As Variant()
'===================================================================================================================
    'Summary: Retrieves the latest Weekly data with Power Query.
    'Inputs: reportType - One of L,D,T to represent what type of report to retrieve.
    '        retrieveCombinedData - true if futures + options data should be retrieved; else, futures only data will be retrieved.
    'Returns: An array of the most recent weekly CFTC data.
    'Notes: Use only on Windows.
'===================================================================================================================
    Dim Data_Query As QueryTable, data() As Variant, url$, _
     Y As Long, reEnableEventsOnExit As Boolean, _
    Found_Data_Query As Boolean, Error_While_Refreshing As Boolean, Workbook_Type$
    
    With Application
        reEnableEventsOnExit = .EnableEvents
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    Workbook_Type = IIf(retrieveCombinedData, "Combined", "Futures_Only")
    
    For Each Data_Query In QueryT.QueryTables
        If InStrB(1, Data_Query.name, ConvertReportTypeEnum(reportType) & "_CFTC_Data_Weekly_" & Workbook_Type) <> 0 Then
            Found_Data_Query = True
            Exit For
        End If
    Next Data_Query
    
    If Not Found_Data_Query Then 'If QueryTable isn't found then create it
Recreate_Query:
        url = "https://www.cftc.gov/dea/newcot/"
        
        Y = Application.Match(reportType, Array(eLegacy, eDisaggregated, eTFF), 0) - 1
        
        If Not retrieveCombinedData Then
            url = url & Array("deafut.txt", "f_disagg.txt", "FinFutWk.txt")(Y)
        Else
            url = url & Array("deacom.txt", "c_disagg.txt", "FinComWk.txt")(Y)
        End If
        
        With QueryT
            Set Data_Query = .QueryTables.Add(Connection:="TEXT;" & url, Destination:=.Range("A1"))
        End With
        
        With Data_Query
            
            .BackgroundQuery = False
            .SaveData = False
            .AdjustColumnWidth = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlOverwriteCells
            
            .TextFileColumnDataTypes = Filter_Market_Columns(convert_skip_col_to_general:=True, reportTypeEnum:=reportType, Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True)
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileCommaDelimiter = True
            
            .name = ConvertReportTypeEnum(reportType) & "_CFTC_Data_Weekly_" & Workbook_Type
            On Error GoTo Delete_Connection
Name_Connection:
            With .WorkbookConnection
                .RefreshWithRefreshAll = False
                .name = ConvertReportTypeEnum(reportType) & "_Weekly CFTC Data: " & Workbook_Type
            End With
            
        End With
        
        On Error GoTo 0
    
    End If
    
    On Error GoTo Failed_To_Refresh 'Recreate Query and try again exactly 1 more time
    
    With Data_Query
        .Refresh False
        With .ResultRange
            .Replace ".", Null, xlWhole
            CFTC_Data_QueryTable_Method = .value 'Store Data in an Array
            .ClearContents 'Clear the Range
        End With
        .Delete
    End With
    
    With Application
        .DisplayAlerts = True
        .EnableEvents = reEnableEventsOnExit
    End With
    
    Exit Function

Delete_Connection: 'Error handler is available when editing parameters for a new querytable and the connection name is already taken by a different query

    ThisWorkbook.Connections("Weekly CFTC Data: " & Workbook_Type).Delete
    On Error GoTo 0
    Resume Name_Connection
    
Failed_To_Refresh:
        
    If Not Data_Query Is Nothing Then
        With Data_Query
            .WorkbookConnection.Delete
            .Delete
        End With
    End If
    
    If Error_While_Refreshing = True Then
        PropagateError Err, "CFTC_Data_QueryTable_Method"
    Else
        Error_While_Refreshing = True
        Resume Recreate_Query
    End If
    
End Function
Public Function Historical_Parse(ByVal File_CLCTN As Collection, reportType As ReportEnum, retrieveCombinedData As Boolean, _
                                  Optional ByRef contractCode$ = vbNullString, _
                                  Optional After_This_Date As Date = 0, _
                                  Optional Kill_Previous_Workbook As Boolean = False) As Variant()
'===================================================================================================================
    'Summary: Retrieves data from Excel Workbooks.
    'Inputs: reportType - One of L,D,T to represent what type of report to retrieve.
    '        retrieveCombinedData - true if futures + options data should be retrieved; else, futures only data will be retrieved.
    '        File_CLCTN - Collection of file paths.
    '        contractCode - If given a value, then Excel workbooks will be filtered for a specific contract code.
    '        After_This_Date - Data after this date will be retrieved.
    '        Kill_Previous_Workbook - If a previous workbook exists then delete it.
    '        parsingMultipleWeeks - Not ALL data may have been downloaded. Maybe only specific years.
    '        Parse_All_Data -
'===================================================================================================================
    Dim Contract_WB As Workbook, Contract_WB_Path$
    
    Dim OS_BasedPathSeparator$
    
    On Error GoTo Historical_Parse_General_Error_Handle
    'filterForSpecificContract = LenB(contractCode) <> 0
    
    #If Mac Then
        OS_BasedPathSeparator = "/"
    #Else
        OS_BasedPathSeparator = "\"
    #End If
    
    Application.ScreenUpdating = False

    Contract_WB_Path = Left$(File_CLCTN(1), InStrRev(File_CLCTN(1), OS_BasedPathSeparator))

    Contract_WB_Path = Contract_WB_Path & ConvertReportTypeEnum(reportType) & "_COT_Yearly_Contracts_" & IIf(retrieveCombinedData, "Combined", "Futures_Only") & ".xlsb"

    If Not FileOrFolderExists(Contract_WB_Path) Then
        ' Compile text files into a single document.
        Set Contract_WB = Historical_TXT_Compilation(File_CLCTN, reportType:=reportType, Saved_Workbook_Path:=Contract_WB_Path, onMac:=False, parsingFuturesAndOptions:=retrieveCombinedData)
    ElseIf IsWorkbookOutdated(Contract_WB_Path) Or Kill_Previous_Workbook = True Then
        On Error Resume Next
        Kill Contract_WB_Path
        On Error GoTo 0
        Set Contract_WB = Historical_TXT_Compilation(File_CLCTN, reportType:=reportType, Saved_Workbook_Path:=Contract_WB_Path, onMac:=False, parsingFuturesAndOptions:=retrieveCombinedData)
    Else
        Set Contract_WB = Workbooks.Open(Contract_WB_Path)
        Contract_WB.Windows(1).Visible = False
    End If
    
    Historical_Parse = Historical_Excel_Aggregation(Contract_WB, getFuturesAndOptions:=retrieveCombinedData, contractCodeToFilterFor:=contractCode, Date_Input:=After_This_Date, ICE_Contracts:=False)
    
    Contract_WB.Close SaveChanges:=False

    Application.StatusBar = vbNullString
    
    Exit Function

Historical_Parse_General_Error_Handle:
    Call PropagateError(Err, "Historical_Parse")
End Function
Public Function Historical_TXT_Compilation(File_Collection As Collection, Saved_Workbook_Path$, onMac As Boolean, reportType As ReportEnum, parsingFuturesAndOptions As Boolean) As Workbook
    
    Dim File_TXT As Variant, fileNumber As Long, Data_STR$, File_Path$(), newWorkbook As Workbook
    
    Dim InfoF() As Variant, columnFormatTypesA() As Variant, D As Long, ICE_Filter As Boolean, ICE_Count As Long, OS_BasedPathSeparator$
    
    Dim File_Name$, CFTC_Count As Long, file_text$, outputFileNumber As Long, outputFileName$ 'g ', DD As Double
    
    Const Comma$ = ","
    
    On Error GoTo Query_Table_Method_For_TXT_Retrieval
        
    If onMac Then
        OS_BasedPathSeparator = "/"
    Else
        OS_BasedPathSeparator = "\"
    End If
    
    outputFileNumber = FreeFile
    outputFileName = Left$(File_Collection(1), InStrRev(File_Collection(1), OS_BasedPathSeparator)) & "Historic.txt"
    
    If FileOrFolderExists(outputFileName) Then Kill outputFileName
    
    Open outputFileName For Append As #outputFileNumber 'Write contents of string to text File
    
    fileNumber = FreeFile
    'Open each file in the collection and write their contents to a string.
    For Each File_TXT In File_Collection
    
        Application.StatusBar = "Parsing " & File_TXT
        DoEvents
        
        Open File_TXT For Input As fileNumber
            
            File_Name = Right$(File_TXT, Len(File_TXT) - InStrRev(File_TXT, OS_BasedPathSeparator))
            
            If File_Name Like "*ICE*" Then
                D = 0
                ICE_Count = ICE_Count + 1
                Do Until EOF(fileNumber)
                    D = D + 1
                    Line Input #fileNumber, Data_STR
                    
                    If (D > 1 And ICE_Count > 1) Or ICE_Count = 1 Then
                        'Only allow printing of headers if on first file
                        Print #outputFileNumber, Data_STR
                    End If
                Loop
            Else
                CFTC_Count = CFTC_Count + 1
                D = 0
                Do Until EOF(fileNumber)
                    D = D + 1
                    Line Input #fileNumber, Data_STR
                    
                    If (D > 1 And CFTC_Count > 1) Or CFTC_Count = 1 Then
                        'Only allow printing of headers if on first file
                        Print #outputFileNumber, Data_STR
                    End If
                Loop
            End If
            
        Close fileNumber
        
        'If LCase$(File_TXT) Like "*weekly*" Then Kill File_TXT
        
    Next File_TXT

    Close #outputFileNumber
    
    Application.StatusBar = "TXT file compilation was successful. Creating Workbook."
    DoEvents
    
    columnFormatTypesA = Filter_Market_Columns(convert_skip_col_to_general:=True, reportTypeEnum:=reportType, Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True, ICE:=False)

    ReDim InfoF(1 To UBound(columnFormatTypesA, 1))
    
    For D = 1 To UBound(columnFormatTypesA, 1) 'Fill in column numbers for use when supplying column filters to OpenTxt
        InfoF(D) = Array(D, columnFormatTypesA(D))
    Next D
    
    Erase columnFormatTypesA
    On Error GoTo Query_Table_Method_For_TXT_Retrieval
    
    #If Mac Then
        D = xlMacintosh
    #Else
        D = xlWindows
    #End If
    With Workbooks
    
        .OpenText fileName:=outputFileName, origin:=D, startRow:=1, DataType:=xlDelimited, _
                                    TextQualifier:=xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, Comma:=True, _
                                    FieldInfo:=InfoF, DecimalSeparator:=".", ThousandsSeparator:=",", TrailingMinusNumbers:=False, _
                                    Local:=False
        Set newWorkbook = Workbooks(.Count)

    End With
    
   With newWorkbook
        .Windows(1).Visible = False
        On Error Resume Next
        If Not onMac Then
            newWorkbook.SaveAs Saved_Workbook_Path, FileFormat:=xlExcel12
        End If
        On Error GoTo 0
    End With
    
    Set Historical_TXT_Compilation = newWorkbook
    Exit Function
Query_Table_Method_For_TXT_Retrieval:
    On Error GoTo -1
    On Error GoTo Parent_Handler

    InfoF = Query_Text_Files(File_Collection, combined_wb:=parsingFuturesAndOptions, reportType:=reportType)
    
    Application.StatusBar = "Data compilation was successful. Creating Workbook."
    DoEvents
    
    Set newWorkbook = Workbooks.Add
    
    With newWorkbook
    
        .Windows(1).Visible = False
        
        With .Worksheets(1)
            .DisplayPageBreaks = False
            .columns("C:C").NumberFormat = "@" 'Format as text
            .Range("A1").Resize(UBound(InfoF, 1), UBound(InfoF, 2)).Value2 = InfoF
        End With
        
    End With
    Set Historical_TXT_Compilation = newWorkbook
    Exit Function
    
Parent_Handler:
    Call PropagateError(Err, "Historical_TXT_Compilation", "An error occurred while compiling text files.")
End Function
Public Function Historical_Excel_Aggregation(Contract_WB As Workbook, _
                                        getFuturesAndOptions As Boolean, _
                                        Optional contractCodeToFilterFor$ = vbNullString, _
                                        Optional Date_Input As Date = 0, _
                                        Optional ICE_Contracts As Boolean = False, _
                                        Optional Weekly_CFTC_TXT As Boolean = False, Optional QueryTable_To_Filter As QueryTable) As Variant()
'===================================================================================================================
    'Summary: Filters and sorts data on a worksheet.
    'Inputs: Contract_WB - Workbook that contains workbook.
    '        contractCodeToFilterFor - If given a value then data will be filtered for this contract code.
    '        combined_workbook - true if futures + options data should be retrieved; else, futures only data will be retrieved.
    '        Date_Input - If not 0 then all data > than this will be filtered for.
    '        filterForSpecificContract - True if specified contract is wanted.
    '        Weekly_CFTC_TXT - True if file data is from the cftc website. Note the url available text file.
    '        QueryTable_To_Filter - Data may be within a query table.
    'Outputs: An array.
'===================================================================================================================
    Dim VAR_DTA() As Variant, Comparison_Operator$, iRow As Long
    
    Dim Combined_CLMN As Long, Disaggregated_Filter_STR$ 'Used if filtering ICE Contracts for Futures and Options
    
    Dim Filtering_QueryTable As Boolean, Source_RNG As Range, filterForSpecificContract As Boolean
    
    Const yymmdd_column As Long = 2
    Const Contract_Code_CLMN As Long = 4 'Column that holds Contract identifiers
    Const ICE_Contract_Code_CLMN As Long = 7
    Const Date_Field As Long = 3
    filterForSpecificContract = LenB(contractCodeToFilterFor) <> 0
    On Error GoTo Finally
    
    Filtering_QueryTable = (Not QueryTable_To_Filter Is Nothing)
    
    If Not Filtering_QueryTable Then
        Application.StatusBar = "Filtering Data."
        DoEvents
        Set Source_RNG = Contract_WB.Worksheets(1).UsedRange
    Else
        Set Source_RNG = QueryTable_To_Filter.ResultRange
    End If
    
    If Source_RNG.Cells.Count = 1 Then 'If worksheet is empty then display message
        GoTo Scripts_Failed_To_Collect_Data
    End If

    On Error GoTo Finally
    
    If ICE_Contracts Or Weekly_CFTC_TXT Then 'Weekly_CFTC_TXT should be unique to CFTC Weekly Text Files at the time of writing
        Comparison_Operator = ">="
    Else
        Comparison_Operator = ">"
    End If
    
    If ICE_Contracts Then
        Disaggregated_Filter_STR = IIf(getFuturesAndOptions, "*Combined*", "*FutOnly*")
        'Find column to be sorted based on the column header.
        On Error GoTo Catch_CombinedColumn_Not_Found
        Combined_CLMN = Application.Match("FutOnly_or_Combined", Source_RNG.Rows(1).Value2, 0)
        Comparison_Operator = Comparison_Operator & Format$(IIf(Date_Input = TimeSerial(0, 0, 0), DateSerial(2000, 1, 1), Date_Input), "YYMMDD")
    Else
        Comparison_Operator = Comparison_Operator & CLng(Date_Input)
    End If
    
    On Error GoTo Finally
    
Check_If_Code_Exists:

    With Source_RNG
    
        On Error Resume Next
        .Parent.ShowAllData
        On Error GoTo Finally
        'Sort date column in ascending order.
        .Sort key1:=.Cells(2, IIf(ICE_Contracts = True, yymmdd_column, Date_Field)), ORder1:=xlAscending, header:=IIf(Weekly_CFTC_TXT, xlNo, xlYes), MatchCase:=False
        ' Filter for wanted dates.
        .AutoFilter Field:=IIf(ICE_Contracts = True, yymmdd_column, Date_Field), Criteria1:=Comparison_Operator, Operator:=xlFilterValues
        
        If ICE_Contracts Then
            ' Sort by Combined Contracts or Futures Only.
            .Sort key1:=.Cells(2, Combined_CLMN), ORder1:=xlAscending, header:=xlYes, MatchCase:=False
            'Filter for "Combined" if condition met.
            .AutoFilter Field:=Combined_CLMN, Criteria1:=Disaggregated_Filter_STR, Operator:=xlFilterValues, VisibleDropDown:=False
        End If

        If filterForSpecificContract Then
            .AutoFilter Field:=Contract_Code_CLMN, Criteria1:=UCase(contractCodeToFilterFor), Operator:=xlFilterValues, VisibleDropDown:=False
            On Error GoTo Catch_ContractCode_Not_Found
        Else
            On Error GoTo Catch_NoVisibleData
        End If
        
        With .SpecialCells(xlCellTypeVisible)
            On Error GoTo Finally
            If .Cells.Count > Source_RNG.Rows(1).Cells.Count Then
            
                If Weekly_CFTC_TXT Then
                    VAR_DTA = .value
                Else
                    If .Areas.Count = 1 Then
                        ' Data excluding headers.
                        VAR_DTA = .Offset(1).Resize(.Rows.Count - 1).value
                    Else
                        VAR_DTA = .Areas(2).value
                    End If
                End If
                
                If ICE_Contracts Then
                
                    For iRow = LBound(VAR_DTA, 1) To UBound(VAR_DTA, 1)
                        
                        If IsEmpty(VAR_DTA(iRow, Contract_Code_CLMN)) Then
                            ' Convert Dates from YYMMDD
                            VAR_DTA(iRow, Date_Field) = DateSerial(Left(VAR_DTA(iRow, yymmdd_column), 2) + 2000, Mid(VAR_DTA(iRow, yymmdd_column), 3, 2), Right(VAR_DTA(iRow, yymmdd_column), 2))
                            ' Map contract codes to different column
                            VAR_DTA(iRow, Contract_Code_CLMN) = VAR_DTA(iRow, ICE_Contract_Code_CLMN)
                            VAR_DTA(iRow, ICE_Contract_Code_CLMN) = Empty
                        End If
                        
                    Next iRow
                    
                End If
            
                Historical_Excel_Aggregation = VAR_DTA
                
            ElseIf filterForSpecificContract Then
                GoTo Catch_ContractCode_Not_Found
            End If
            
        End With 'End .SpecialCells(xlCellTypeVisible)
        
    End With
    
    If Not Filtering_QueryTable Then
        Application.StatusBar = vbNullString
        DoEvents
    End If

Finally:
    If Err.Number <> 0 Then
        If Not Contract_WB Is ThisWorkbook Then
            With Contract_WB
                .Close False
                Kill .fullName
            End With
            Application.StatusBar = vbNullString
        End If
        PropagateError Err, "Historical_Excel_Aggregation"
    End If
    Exit Function
    
Catch_ContractCode_Not_Found: 'Used when user has input an invalid contract code

    If MsgBox("The Selected Contract Code [" & contractCodeToFilterFor & "] wasn't found" & vbNewLine & "Would you like to try again with a different Contract Code?", vbYesNo, "Please choose") _
                = vbYes Then
        contractCodeToFilterFor = UCase(Application.InputBox("Please supply the Contract Code of the desired contract"))
        GoTo Check_If_Code_Exists
    Else
        Application.StatusBar = vbNullString:
        If Not Contract_WB Is ThisWorkbook Then
            Contract_WB.Close False
        End If
        
        Call Re_Enable
        End
    End If
Catch_NoVisibleData:
    AppendErrorDescription Err, "Attempt to retrieve data from compiled worksheet failed. No visible data after filtering."
    GoTo Finally
Scripts_Failed_To_Collect_Data:
    AppendErrorDescription Err, "No data found on worksheet."
    GoTo Finally
Catch_CombinedColumn_Not_Found:
    AppendErrorDescription Err, "Could not locate Combined column in Disaggregated report."
    GoTo Finally
End Function
Public Function Weekly_Text_File(filePath As String, reportType As ReportEnum, retrieveCombinedData As Boolean) As Variant()
'===================================================================================================================
    'Summary: Filters and sorts data on a worksheet.
    'Inputs: Contract_WB - Workbook that contains workbook.
    '        contractCodeToFilterFor - If given a value then data will be filtered for this contract code.
    '        combined_workbook - true if futures + options data should be retrieved; else, futures only data will be retrieved.
    '        Date_Input - If not 0 then all data > than this will be filtered for.
    '        filterForSpecificContract - True if specified contract is wanted.
    '        Weekly_CFTC_TXT - True if file data is from the cftc website. Note the url available text file.
    '        QueryTable_To_Filter - Data may be within a query table.
    'Outputs: An array.
'===================================================================================================================
    Dim D As Long, FilterC() As Variant, InfoF() As Variant
    
    FilterC = Filter_Market_Columns(convert_skip_col_to_general:=True, Return_Filter_Columns:=True, reportTypeEnum:=reportType, Return_Filtered_Array:=False, Create_Filter:=True)
    
    ReDim InfoF(1 To UBound(FilterC, 1))
    
    For D = 1 To UBound(FilterC, 1)
        InfoF(D) = Array(D, FilterC(D))
    Next D
    
    Erase FilterC
    
    #If Mac Then
        D = xlMacintosh
    #Else
        D = xlWindows
    #End If

    On Error GoTo Error_While_Opening_Text_File

    With Workbooks
        .OpenText fileName:=filePath, origin:=D, startRow:=1, DataType:=xlDelimited, _
                            TextQualifier:=xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, Comma:=True, _
                            FieldInfo:=InfoF, DecimalSeparator:=".", ThousandsSeparator:=",", TrailingMinusNumbers:=False, _
                            Local:=False
        With .item(.Count)
            .Windows(1).Visible = False
             Weekly_Text_File = .Worksheets(1).UsedRange.value
            .Close False
        End With
    End With
    
    Kill filePath
    
    Exit Function

Error_While_Opening_Text_File:
    PropagateError Err, "Weekly_Text_File", "Error while attempting to open a Weekly based Text File."
    
End Function
Public Function Filter_Market_Columns(Return_Filter_Columns As Boolean, _
                                       Return_Filtered_Array As Boolean, _
                                       convert_skip_col_to_general As Boolean, _
                                       reportTypeEnum As ReportEnum, _
                                       Optional Create_Filter As Boolean = True, _
                                       Optional ByVal inputA As Variant, _
                                       Optional ICE As Boolean = False, _
                                       Optional ByVal Column_Status As Variant) As Variant
'======================================================================================================
'Generates an array referencing RAW data columns to determine if they should be kept or not
'If and array is given an return_filtered_array=True then the array will be filtered column wise based on the previous array
'======================================================================================================

    Dim ZZ As Long, output() As Variant, v As Long, Y As Long, columnOffset As Long, columnsRemaining As Long, _
    contractIdField As Long, alternateCftcCodeColumn As Long, _
    columnInOutput As Long, finalColumnIndex As Long, nameField As Long, filterLength As Long
    
    Dim CFTC_Wanted_Columns() As Variant, dateField As Long, skip_value As XlColumnDataType, twoDimensionalArray As Boolean
    
    On Error GoTo Propogate
    
    CFTC_Wanted_Columns = GetAvailableFieldsTable(reportTypeEnum).DataBodyRange.columns(2).Value2
    
    If ICE Then
        dateField = 2
        contractIdField = 7
    Else
        dateField = 3
        contractIdField = 4
        nameField = 1
    End If
        
    Select Case reportTypeEnum
        Case eLegacy
            alternateCftcCodeColumn = 127
        Case eDisaggregated
            alternateCftcCodeColumn = 187
        Case eTFF
            alternateCftcCodeColumn = 83
    End Select
    
    If convert_skip_col_to_general Then
        skip_value = xlGeneralFormat
    Else
        skip_value = xlSkipColumn
    End If
    
    If IsArray(inputA) Or IsMissing(inputA) Then
        filterLength = UBound(CFTC_Wanted_Columns, 1)
    Else
        filterLength = inputA.Count
    End If
    
    If Create_Filter = True And IsMissing(Column_Status) Then 'IF column Status is empty or if it is empty
        ReDim Column_Status(1 To filterLength)

        For v = LBound(Column_Status) To UBound(Column_Status)
            ' Allows entry into block regardless of if ICE or CFTC is needed for dates or contract code
            On Error GoTo Catch_OutsideBounds
            
            If (CFTC_Wanted_Columns(v, 1) = True Or v = dateField Or v = contractIdField) Then
            
                Select Case v
                
                    Case dateField 'column 2 or 3 depending on if ICE or not
                        Column_Status(v) = IIf(ICE, xlGeneralFormat, xlYMDFormat) 'xlMDYFormat
                    Case nameField, contractIdField
                        Column_Status(v) = xlTextFormat
                    Case 2, 3, 4, 7 'These numbers may overlap with dates column or contract field
                                    'The previous cases will prevent it from executing unnecessarily depending on if ICE or not
                        Column_Status(v) = skip_value
                    Case Else
                        Column_Status(v) = xlGeneralFormat
                End Select
            Else
                If v = alternateCftcCodeColumn And convert_skip_col_to_general Then
                    Column_Status(v) = xlTextFormat
                Else
                    Column_Status(v) = skip_value
                End If
            End If
Assign_Next_FilterColumn:
        Next v
    End If
    
    On Error GoTo Propogate
    
    If Return_Filter_Columns = True Then
        Filter_Market_Columns = Column_Status
    ElseIf Return_Filtered_Array = True Then
         'Don't worry about text files.they are filtered in the same sub that they are opened in
         'FYI dateField would be 2 if doing TXT files..2 is used for ICE contracts because of exchange inconsistency
        On Error Resume Next
        
        If IsArray(inputA) Then
            Y = 0
            Do 'Determine the total number of dimensions
                Y = Y + 1
                v = LBound(inputA, Y)
            Loop Until Err.Number <> 0
            On Error GoTo 0
            If Y - 1 = 2 Then twoDimensionalArray = True
        ElseIf TypeName(inputA) = "Collection" Then
            twoDimensionalArray = False
        End If
        
        If twoDimensionalArray Then
            ReDim output(1 To UBound(inputA, 1), 1 To UBound(Filter(Column_Status, xlSkipColumn, False)) + 1)
            finalColumnIndex = UBound(output, 2)
        Else
            ReDim output(1 To UBound(Filter(Column_Status, xlSkipColumn, False)) + 1)
            finalColumnIndex = UBound(output, 1)
        End If
        
        Y = 0
        
        For v = LBound(Column_Status) To UBound(Column_Status)
            If Column_Status(v) <> xlSkipColumn Then
                Select Case v
                    Case nameField
                        columnInOutput = 2
                    Case dateField
                        columnInOutput = 1
                    Case contractIdField
                        columnInOutput = finalColumnIndex
                    Case Else
                        'Find the next value that excludes the above cases
                        Do
                            Y = Y + 1
                        Loop Until 2 < Y And Y < finalColumnIndex
                        columnInOutput = Y
                End Select
                
                If twoDimensionalArray Then
                    For ZZ = LBound(output, 1) To UBound(output, 1)
                        output(ZZ, columnInOutput) = inputA(ZZ, v - LBound(Column_Status) + LBound(inputA, 2))
                    Next ZZ
                Else
                    If IsObject(inputA(v)) Then
                        Set output(columnInOutput) = inputA(v - LBound(Column_Status) + LBound(inputA))
                    Else
                        output(columnInOutput) = inputA(v - LBound(Column_Status) + LBound(inputA))
                    End If
                End If
            End If
        Next v
        
        Filter_Market_Columns = output
    End If
    
    Exit Function
    
Catch_OutsideBounds:
    If Not IsArray(inputA) And Err.Number = 9 Then
        Column_Status(v) = xlGeneralFormat
        Resume Assign_Next_FilterColumn
    Else
        Err.Description = "Outside Bounds"
        GoTo Propogate
    End If
Propogate:
    PropagateError Err, "Filter_Market_Columns"
End Function
Public Function Query_Text_Files(ByVal TXT_File_Paths As Collection, reportType As ReportEnum, combined_wb As Boolean) As Variant
'===================================================================================================================
    'Summary: Queries text files in TXT_File_Paths and adds their contents(array) to a collection
    'Inputs: reportType - One of L,D,T to represent what type of report to retrieve.
    '        combined_wb  - true if futures + options data should be retrieved; else, futures only data will be retrieved.
    'Returns: An array of the most recent weekly CFTC data.
    'Notes: Use only on Windows.
'===================================================================================================================
    Dim QT As QueryTable, file As Variant, Found_QT As Boolean, Field_Info() As Variant, Output_Arrays As New Collection, _
    Field_Info_ICE() As Variant
     
    Dim headerCount As Long
    
     On Error GoTo Propagate
     
    For Each QT In QueryT.QueryTables 'Search for the following query if it exists
        If InStrB(1, QT.name, "TXT Import") <> 0 Then
            Found_QT = True
            Exit For
        End If
    Next QT
    
    Field_Info = Filter_Market_Columns(convert_skip_col_to_general:=True, reportTypeEnum:=reportType, Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True) '^^ CFTC Column filter
    
    If reportType = eDisaggregated Then 'ICE Data column filter
        Field_Info_ICE = Filter_Market_Columns(convert_skip_col_to_general:=True, reportTypeEnum:=reportType, Return_Filter_Columns:=True, Return_Filtered_Array:=False, Create_Filter:=True, ICE:=True)
    End If
    
    For Each file In TXT_File_Paths
        
        Application.StatusBar = "Querying: " & file
        DoEvents
        
        If Not Found_QT Then
            Set QT = QueryT.QueryTables.Add(Connection:="TEXT;" & file, Destination:=QueryT.Cells(1, 1))
            With QT
                .name = "TXT Import"
                .BackgroundQuery = False
                .SaveData = False
            End With
            Found_QT = True 'So that this statement isn't executed again
        End If
        
        With QT
            .Connection = "TEXT;" & file
            .TextFileCommaDelimiter = True
            .TextFileConsecutiveDelimiter = False
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            
            If file Like "*.csv" And reportType = eDisaggregated Then 'ICE Workbooks
                .TextFileColumnDataTypes = Field_Info_ICE
            Else
                .TextFileColumnDataTypes = Field_Info
            End If
            
            .RefreshStyle = xlOverwriteCells
            .AdjustColumnWidth = False
            .Destination = QueryT.Cells(1, 1)
            
            .Refresh False
            
            headerCount = headerCount + 1
            
            With .ResultRange
                If headerCount = 1 Then
                    Output_Arrays.Add .Value2
                Else
                    Output_Arrays.Add .Offset(1).Resize(.Rows.Count - 1).Value2
                End If
                .ClearContents
            End With
        End With
    
    Next file
    
    If Output_Arrays.Count > 1 Then
        Query_Text_Files = CombineArraysInCollection(Output_Arrays, Append_Type.Multiple_2d)
    Else
        Query_Text_Files = Output_Arrays(1)
    End If
    
    QT.Delete
    
    Exit Function
    
Propagate:
    If Not QT Is Nothing Then
        QT.Delete
    End If
    PropagateError Err, "Query_Text_Files"
End Function
Public Function TryGetPriceData(ByRef inputData As Variant, ByVal inputDataPriceColumn As Long, contractDataOBJ As ContractInfo, _
    overwriteAllPrices As Boolean, datesAreInColumnOne As Boolean, Optional yahooCookie As String) As Boolean
'===================================================================================================================
    'Summary: Retrieves price data.
    'Inputs: inputData -
    '        inputDataPriceColumn - Column within inputData to store prices in.
    '        contractDataOBJ - Contract instance that contains symbol info and where to get prices from.
    '        overwriteAllPrices - Clears price column in inputData.
    '        datesAreInColumnOne -  If true then dates are assumed to be in column 1 else in column 3.
'===================================================================================================================

    Dim Start_Date As Date, End_Date As Date, url$, Yahoo_Finance_Parse As Boolean, Stooq_Parse As Boolean
    
    Dim unixEndTime&, unixStartTime&, PriceSymbol$, dateColumn As Long, Response_STR$
    
    Const unmodified_COT_DateColumn As Long = 3, UseAlternateLink As Boolean = True
    
    'Yahoo bases there URLs on the date converted to UNIX time
    Const UnixStartDate As Date = #1/1/1970#
    
    With contractDataOBJ
        If Not .HasSymbol Then Exit Function
        PriceSymbol = .PriceSymbol
        Yahoo_Finance_Parse = .UseYahooPrices
        Stooq_Parse = Not Yahoo_Finance_Parse
    End With
    
    dateColumn = IIf(datesAreInColumnOne, 1, unmodified_COT_DateColumn)
    
    On Error GoTo Exit_Price_Parse
    
    Start_Date = inputData(1, dateColumn)
    End_Date = inputData(UBound(inputData, 1), dateColumn)
    
    If Yahoo_Finance_Parse Then
        
        End_Date = DateAdd("d", 1, End_Date) '1 more day than is in range to encapsulate that day
        unixStartTime = DateDiff("s", UnixStartDate, Start_Date) 'Convert to UNIX time
        unixEndTime = DateDiff("s", UnixStartDate, End_Date) 'An extra day is added to encompass the End Day
        
        If Not UseAlternateLink Then
            url = "https://query1.finance.yahoo.com/v7/finance/download/"
        Else
            url = "https://query2.finance.yahoo.com/v8/finance/chart/"
        End If
        
        url = url & PriceSymbol & "?period1=" & unixStartTime & "&period2=" & unixEndTime & "&interval=1d&events=history&includeAdjustedClose=true"
                
    End If
    
    #If Mac Then
    
        Dim QueryTable_Found As Boolean, QT As QueryTable
        
        Const QT_Connection_Type$ = "TEXT;", Query_Name$ = "Yahoo Finance Query"
        
        On Error GoTo Exit_Price_Parse
        'On Error GoTo 0
        'Determine if QueryTable Exists
        For Each QT In QueryT.QueryTables
            If InStrB(QT.name, Query_Name) <> 0 Then  'Instr method used in case Excel appends a number to the name
                QueryTable_Found = True
                Exit For
            End If
        Next QT
        
        If Not QueryTable_Found Then Set QT = QueryT.QueryTables.Add(QT_Connection_Type & url, QueryT.Cells(1, 1))
        
        With QT
        
            If Not QueryTable_Found Then
                .BackgroundQuery = False
                .name = Query_Name
                ' If an error occurs then delete the already existing connection and then try again.
                'On Error GoTo Workbook_Connection_Name_Already_Exists
                '.WorkbookConnection.Name = Replace$(Query_Name, "Query", "Prices")
                'On Error GoTo Exit_Price_Parse
            Else
                .Connection = QT_Connection_Type & url
            End If
            
            .RefreshOnFileOpen = False
            .RefreshStyle = xlOverwriteCells
            .SaveData = False
            
            On Error GoTo Remove_QT_And_Connection 'Delete both the Querytable and the connection and exit the sub
             .Refresh False
            On Error GoTo Exit_Price_Parse
            
            With .ResultRange
                ' .value returns an array of comma separated values in a single column.
                If Yahoo_Finance_Parse Or Stooq_Parse Then
                    If UseAlternateLink Then
                        Response_STR = .Value2
                    End If
                End If
                
                .ClearContents
            End With
        End With
        
        Set QT = Nothing
        Query_Name = vbNullString
        QT_Connection_Type = vbNullString
        
    #Else
        On Error GoTo Exit_Price_Parse
        TryGetRequest url, Response_STR
    #End If
    
    Dim priceByDate As Object, adjustedClose$(), timeStamps$(), oldestDate As Date, iCount As Long
        
    If UseAlternateLink And InStrB(Response_STR, "timestamp") > 0 Then

        adjustedClose = Split(Split(Split(Split(Response_STR, """adjclose"":")(2), "]")(0), "[")(1), ",")
        timeStamps = Split(Split(Split(Split(Response_STR, """timestamp"":")(1), "]")(0), "[")(1), ",")
        ' Setting priceByDate to something other than nothing will allow the function to know that data was returned.
        
        Set priceByDate = GetDictionaryObject()

        On Error Resume Next
        
        With priceByDate
            For iCount = LBound(timeStamps) To UBound(timeStamps)
                .item(CLng(DateAdd("s", CLng(timeStamps(iCount)), UnixStartDate))) = CDbl(adjustedClose(iCount))
            Next iCount
            oldestDate = .Keys(0)
        End With
        
        With Err
            If .Number <> 0 Then .Clear
        End With
        
        On Error GoTo Exit_Price_Parse
        
        Erase timeStamps: Erase adjustedClose
    End If
            
    url = vbNullString
    On Error GoTo Exit_Price_Parse
    
    If Not priceByDate Is Nothing And (Yahoo_Finance_Parse Or Stooq_Parse) Then
        
        If UseAlternateLink Then
            If priceByDate Is Nothing Or InStrB(Response_STR, 404) = 1 Or LenB(Response_STR) = 0 Or InStr(LCase$(Response_STR), "error") < 100 Then
                'Something likely wrong with the URl
                Exit Function
            End If
            
            Dim foundAlternateDate As Boolean, integerDate As Long
            
            With priceByDate
                For iCount = LBound(inputData, 1) To UBound(inputData, 1)
                    integerDate = inputData(iCount, dateColumn)
                    If .Exists(integerDate) Then
                        ' Exact match found.
                        inputData(iCount, inputDataPriceColumn) = .item(integerDate)
                    Else
                        ' Assign using the last close within a 1 week period.
                        Do
                            integerDate = DateAdd("d", -1, integerDate)
                            If .Exists(integerDate) Then
                                inputData(iCount, inputDataPriceColumn) = .item(integerDate)
                                foundAlternateDate = True
                            End If
                        Loop While foundAlternateDate = False And DateDiff("d", integerDate, inputData(iCount, dateColumn)) <= 7 And integerDate > oldestDate
                        
                        If Not foundAlternateDate Then
                            If overwriteAllPrices And Not IsEmpty(inputData(iCount, inputDataPriceColumn)) Then
                                inputData(iCount, inputDataPriceColumn) = Empty
                            End If
                        Else
                            foundAlternateDate = False
                        End If
                    End If
                Next iCount
            End With
        Else
'            Dim D_OHLC_AV$(), priceData$(), Initial_Split_CHR$, Secondary_Split_STR$, closePriceColumn as Long
'            If Yahoo_Finance_Parse Then
'                'Finding Splitting_Charachter
'                Initial_Split_CHR = Mid$(Response_STR, InStr(1, Response_STR, "Volume") + Len("volume"), 1)
'            ElseIf Stooq_Parse Then
'                Initial_Split_CHR = vbNewLine
'            End If
'
'            priceData = Split(Response_STR, Initial_Split_CHR)
'            Secondary_Split_STR = ","
'            closePriceColumn = 4 'Base 0 location of close prices within the queried array
'
'            If LenB(Response_STR) <> 0 Then Response_STR = vbNullString
'            If LenB(Initial_Split_CHR) <> 0 Then Initial_Split_CHR = vbNullString
'
'            Dim dateKey$
'
'            #If Mac Then
'                Set adjustedClose = New Dictionary
'            #Else
'                Set adjustedClose = CreateObject("Scripting.Dictionary")
'            #End If
'
'            With adjustedClose
'                For iCount = LBound(priceData) To UBound(priceData)
'                    D_OHLC_AV = Split(priceData(iCount), Secondary_Split_STR)
'                    If IsNumeric(D_OHLC_AV(closePriceColumn)) Then
'                        .Item(CStr(D_OHLC_AV(0))) = CDbl(D_OHLC_AV(closePriceColumn))
'                    End If
'                Next iCount
'
'                For iCount = LBound(inputData) To UBound(inputData)
'                    dateKey = inputData(iCount, dateColumn)
'                    If .Exists(dateKey) Then
'                        inputData(iCount, inputDataPriceColumn) = .Item(dateKey)
'                    ElseIf overwriteAllPrices Then
'                        inputData(iCount, inputDataPriceColumn) = Empty
'                    End If
'                Next iCount
'            End With
        End If
    End If
    
    TryGetPriceData = Not priceByDate Is Nothing
    
Exit_Price_Parse:
        
    Exit Function
    
#If Mac Then

Remove_QT_And_Connection:
    QT.Delete
    Exit Function
Workbook_Connection_Name_Already_Exists:
    ThisWorkbook.Connections(Replace$(Query_Name, "Query", "Prices")).Delete
    
    QT.WorkbookConnection.name = Replace$(Query_Name, "Query", "Prices")
    Resume Next
    
#End If

'Error_While_Splitting:
'    If Err.Number = 13 Then 'type mismatch error from using cdate on a non-date string
'        Resume Increment_Y
'    Else
'        Exit Function
'    End If
'Propagate:
'    PropagateError Err, "TryGetPriceData"
End Function

Public Sub Paste_To_Range(Optional Table_DataB_RNG As Range, Optional Data_Input As Variant, _
        Optional Sheet_Data As Variant, Optional Historical_Paste As Boolean = False, _
        Optional Target_Sheet As Worksheet, _
        Optional Overwrite_Data As Boolean = False)
'===================================================================================================================
    'Summary: Places data at the bottom of a specified table.
    'Inputs: Table_DataB_RNG -
    '        Data_Input - Data to place in table when Historical_Paste is False.
    '        Sheet_Data - Data that is already present within a table or data to place if Historical_Paste is True.
    '        Historical_Paste - True if a table needs to be created and not normal weekly data additions.
    '        Target_Sheet - Worksheet that data will be placed on.
    '        Overwrite_Data - True if you want to clear any already present rows. ONly applicable if Historical_Paste is True
'===================================================================================================================
    Dim Model_Table As ListObject, Invalid_STR$(), i As Long, _
    Invalid_Found() As Variant, newRowNumber As Long, rowNumber As Long, ColumnNumber As Long
    
    If Not Historical_Paste Then 'If Weekly/Block data addition
        
        If Not Overwrite_Data Then
            'Search in reverse order for dates that are too old to add to sheet.
            'Compare the Max date in data to upload and alrady on the sheet to determine how much if any of the data should be placed on the sheet.
            
            i = LBound(Data_Input, 1)

            Do While Data_Input(i, 1) <= Sheet_Data(UBound(Sheet_Data, 1), 1)
                i = i + 1
                If i > UBound(Data_Input, 1) Then Exit Do
            Loop

            If i > UBound(Data_Input, 1) Then
                Exit Sub
            ElseIf i <> LBound(Data_Input, 1) Then
            
                ReDim Invalid_Found(1 To UBound(Data_Input, 1) - i, 1 To UBound(Data_Input, 2))
                'Fill array with wanted data.
                For rowNumber = i To UBound(Data_Input, 1)
                
                    newRowNumber = newRowNumber + 1
                    
                    For ColumnNumber = 1 To UBound(Data_Input, 2)
                        Invalid_Found(newRowNumber, ColumnNumber) = Data_Input(rowNumber, ColumnNumber)
                    Next ColumnNumber
                    
                Next rowNumber
                
                Data_Input = Invalid_Found
            End If
        Else
            Table_DataB_RNG.ClearContents
            'Table_DataB_RNG.ListObject.AutoFilter.ShowAllData
        End If
        
        On Error GoTo No_Table
        
        Dim rowCountBeforeResize&, newDataRow&
        
        With Table_DataB_RNG
            .Worksheet.DisplayPageBreaks = False
            'Overwritten range depends on Overwrite Data Boolean, If true then overwrite all data on the worksheet
            With .ListObject
                
                rowCountBeforeResize = .ListRows.Count
                newDataRow = UBound(Sheet_Data, 1) + 1
                If Not Overwrite_Data Then
                    If rowCountBeforeResize <> UBound(Data_Input, 1) + UBound(Sheet_Data, 1) Then
                        .Resize .Range.Resize(UBound(Data_Input, 1) + UBound(Sheet_Data, 1) + 1, .ListColumns.Count)
                    End If
                ElseIf rowCountBeforeResize <> UBound(Data_Input, 1) Then
                    .Resize .Range.Resize(UBound(Data_Input, 1) + 1, .ListColumns.Count)
                End If
            End With
            
            .Cells(IIf(Overwrite_Data = False, newDataRow, 1), 1).Resize(UBound(Data_Input, 1), UBound(Data_Input, 2)).Value2 = Data_Input  'bottom row +1,1st column
        
        End With 'pastes the bottom row of the array if bottom date is greater than previous
    ElseIf Historical_Paste = True Then 'pastes to active sheet and retrieves headers from sheet

        If Overwrite_Data Then
            MsgBox "Within the Paste_To_Range sub OVerwrite_Data cannot be true if Historical_Paste is true."
            Exit Sub
        End If
        
        On Error GoTo PROC_ERR_Paste
        Set Model_Table = GetAvailableContractInfo(1).TableSource
        
        With Model_Table
            .DataBodyRange.Copy 'copy and paste formatting
            Target_Sheet.Range(.HeaderRowRange.Address).Value2 = .HeaderRowRange.Value2
        End With
        
        With Target_Sheet
        
            .Range("A2").Resize(UBound(Sheet_Data, 1), UBound(Sheet_Data, 2)).Value2 = Sheet_Data
            
            With .ListObjects.Add(xlSrcRange, .UsedRange, , xlYes)
                .DataBodyRange.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            End With
            
            .Hyperlinks.Add Anchor:=.Cells(1, 1), Address:=vbNullString, SubAddress:="'" & HUB.name & "'!A1", TextToDisplay:=.Cells(1, 1).Value2
            
            On Error GoTo Re_Name '{Finding Valid Worksheet Name}
            .name = Split(Sheet_Data(UBound(Sheet_Data, 1), 2), " -")(0)
        
        End With
        
        Application.StatusBar = "Data has been added to sheet. Calculating Formulas"
            
    End If
    
    On Error GoTo 0
    
    Exit Sub
        
Re_Name:
   MsgBox " If you were attempting to add a new contract then the Worksheet name could not be changed automatically."
    Resume Next
PROC_ERR_Paste:
    MsgBox "Error: (" & Err.Number & ") " & Err.Description, vbCritical
    Resume Next
No_Table:
    MsgBox "If you are seeing this then either a table could not be found in cell A1 or your version " & _
    "of Excel does not support the listbody object. Further data will not be updated. Contact me via email."
    Call Re_Enable: End
End Sub

Public Function CreateFieldInfoMap(externalHeaders As Variant, localDatabaseHeaders As Variant, externalHeadersFromSocrataAPI As Boolean) As Collection
'==========================================================================================================
' Creates a Collection of FieldInfo insances for fields that are found within both externalHeaders and localDatabaseHeaders.
' Variables:
'   externalHeaders: 1D array of column names associated with each field from apiData
'   databaseFieldsByEditedName: Columns from a localy saved database.
'==========================================================================================================
    Dim iCount As Long, externalHeaderIndexByEditedName As New Collection, item As Variant, databaseFieldsByEditedName As New Collection, FI As FieldInfo

    On Error GoTo Abandon_Processes
    ' Column names from the api source are often spelled incorrectly or aren't standardized in their naming.
    With externalHeaderIndexByEditedName
        For iCount = LBound(externalHeaders) To UBound(externalHeaders)
            If externalHeadersFromSocrataAPI Then
                If InStrB(externalHeaders(iCount), "spead") <> 0 Then externalHeaders(iCount) = Replace$(externalHeaders(iCount), "spead", "spread")
                If InStrB(externalHeaders(iCount), "postions") <> 0 Then externalHeaders(iCount) = Replace$(externalHeaders(iCount), "postions", "positions")
                If InStrB(externalHeaders(iCount), "open_interest") <> 0 Then externalHeaders(iCount) = Replace$(externalHeaders(iCount), "open_interest", "oi")
                If InStrB(externalHeaders(iCount), "__") <> 0 Then externalHeaders(iCount) = Replace$(externalHeaders(iCount), "__", "_")
                .Add iCount, externalHeaders(iCount)
            Else
                .Add iCount, StandardizedDatabaseFieldNames(CStr(externalHeaders(iCount)))
            End If
        Next iCount
    End With
    
    Dim FieldInfoMap As New Collection, endings$(), EditedName$, mainLoopCount As Long
    
    With databaseFieldsByEditedName
        For iCount = LBound(localDatabaseHeaders) To UBound(localDatabaseHeaders)
            EditedName = StandardizedDatabaseFieldNames(CStr(localDatabaseHeaders(iCount)))
            .Add Array(EditedName, localDatabaseHeaders(iCount)), EditedName
        Next iCount
    End With
        
    Dim endingsIterator As Long, endingStrippedName$, digitIncrement As Long, _
    foundMainEditedName As Boolean, secondaryIndex As Long, newKey$
    
    ' This array is ordered in the manner that they appear within the api columns.
    endings = Split("_all,_old,_other", ",")
    
    ' Loop through databaseFieldsByEditedName and determine if the edited name exists within externalHeaderIndexByEditedName.
    ' Regardless of if it does, create a FieldInfo instance and add to FieldInfoMap.
    With FieldInfoMap
        For Each item In databaseFieldsByEditedName
            EditedName = item(0)
            mainLoopCount = mainLoopCount + 1
            foundMainEditedName = False
            
            If HasKey(FieldInfoMap, EditedName) Then
                ' FieldInfo instance has already been added. Ensure its order within the collection.
                foundMainEditedName = True
                Set FI = .item(EditedName)
                .Remove EditedName
                .Add FI, FI.EditedName, After:=databaseFieldsByEditedName(mainLoopCount - 1)(0)
            ElseIf HasKey(externalHeaderIndexByEditedName, EditedName) Then
                ' Exact match between column name sources.
                Set FI = CreateFieldInfoInstance(EditedName, ColumnIndex:=externalHeaderIndexByEditedName(EditedName), mappedName:=CStr(item(1)), IsMissing:=False, fromSocrata:=externalHeadersFromSocrataAPI)

                If .Count = 0 Then
                    .Add FI, EditedName
                Else
                    .Add FI, EditedName, After:=databaseFieldsByEditedName(mainLoopCount - 1)(0)
                End If
                
                externalHeaderIndexByEditedName.Remove EditedName
                foundMainEditedName = True
            Else
                For endingsIterator = LBound(endings) To UBound(endings)
                    ' Checking if the name ends with the pattern.
                    If EditedName Like "*" + endings(endingsIterator) Then
                        endingStrippedName = Replace$(EditedName, endings(endingsIterator), vbNullString)
                        digitIncrement = 0
                        
                        For secondaryIndex = endingsIterator To UBound(endings)
                            Dim apiFieldName$, placementKnown As Boolean
                            
                            newKey = vbNullString
                            placementKnown = False
                            
                            If secondaryIndex = endingsIterator And HasKey(externalHeaderIndexByEditedName, endingStrippedName) Then
                                newKey = EditedName
                                apiFieldName = endingStrippedName
                                placementKnown = True
                                foundMainEditedName = True
                            ElseIf secondaryIndex > endingsIterator Then
                                
                                digitIncrement = digitIncrement + 1
                                apiFieldName = endingStrippedName & "_" & digitIncrement
                                
                                If HasKey(externalHeaderIndexByEditedName, apiFieldName) Then
                                    newKey = endingStrippedName + endings(secondaryIndex)
                                End If
                                
                            End If
                            
                            If LenB(newKey) <> 0 Then
                                Set FI = CreateFieldInfoInstance(newKey, externalHeaderIndexByEditedName(apiFieldName), CStr(databaseFieldsByEditedName(newKey)(1)), False, fromSocrata:=externalHeadersFromSocrataAPI)

                                If placementKnown Then
                                    .Add FI, newKey, After:=databaseFieldsByEditedName(mainLoopCount - 1)(0)
                                Else
                                    .Add FI, newKey
                                End If
                                                            
                                ' Removal is just for viewing how many and which api columns weren't found.
                                externalHeaderIndexByEditedName.Remove apiFieldName
                            End If
                        Next secondaryIndex
                    End If
                Next endingsIterator
            End If
            ' This conditional adds a FieldInfo instance with the IsMissing property set to true.
            If Not foundMainEditedName Then
                Set FI = CreateFieldInfoInstance(EditedName, -1, CStr(item(1)), True, fromSocrata:=externalHeadersFromSocrataAPI)
                'Place after previous field by name.
                .Add FI, EditedName, After:=databaseFieldsByEditedName(mainLoopCount - 1)(0)
            End If
        Next item
    End With
    
    Set CreateFieldInfoMap = FieldInfoMap
    Exit Function
Abandon_Processes:
    PropagateError Err, "CreateFieldInfoMap"
End Function



