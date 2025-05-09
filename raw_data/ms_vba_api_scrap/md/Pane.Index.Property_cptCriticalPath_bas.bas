Attribute VB_Name = "cptCriticalPath_bas"
'<cpt_version>v3.1.0</cpt_version>
Option Explicit
Private CritField As String 'Stores comma seperated values for each task showing which paths they are a part of
Private GroupField As String 'Stores a single value - used to group/sort tasks in final CP view
'Custom type used to store driving path vars
Type DrivingPaths
    PrimaryFloat As Double 'Stores True Float vlaue for Primary Driving Path
    FindPrimary As Boolean 'Tracks evluation progress by noting Primary Found
    SecondaryFloat As Double 'Stores True Float vlaue for Secondary Driving Path
    FindSecondary As Boolean 'Tracks evluation progress by noting Secondary Found
    TertiaryFloat As Double 'Stores True Float vlaue for Secondary Driving Path
    FindTertiary As Boolean 'Tracks evluation progress by noting Tertiar Found
    FourthFloat As Double 'Stores True Float Value for fourth Driving Path
    FindFourth As Boolean 'Tracks evaluation progress by noting fourth found
    FifthFloat As Double 'Stores True Float Value for fourth Driving Path
    FindFifth As Boolean 'Tracks evaluation progress by noting fourth found
End Type
Private tDrivingPaths As DrivingPaths 'var to store DrivingPaths type
Private SecondaryDrivers() As String 'Array of Secondary Drivers to be analyzed
Private SecondaryDriverCount As Integer 'Count of secondary Drivers
Private TertiaryDrivers() As String 'Array of tertiary drivers to be analyzed
Private TertiaryDriverCount As Integer 'Count of tertiary drivers
Private FourthDrivers() As String 'Array of fourth drivers to be analyzed
Private FourthDriverCount As Integer 'Count of fourth drivers
Private FifthDrivers() As String 'Array of fifth drivers to be analyzed
Private FifthDriverCount As Integer 'Count of fifth drivers
Private AnalyzedTasks As Collection 'Collection of task relationships analyzied (From UID - To UID); unique to each path analysis
'Custom type used to store Driving Task data
Type DrivingTask
    UID As String
    tFloat As Double
End Type
Private DrivingTasks() As DrivingTask 'var to store DrivingTask type
Private drivingTasksCount As Integer 'count of DrivingTasks
Public singlePath As Boolean 'cpt controlled var for limited results to a single path
Public export_to_PPT As Boolean 'cpt controlled var for controlling user notification of completed analysis
Private CustTextFields() As String 'v2.9.0 Array of custTextFields
Private CustNumFields() As String 'v2.9.0 Array of custNumFields
Private curProj As Project 'Stores active user project - not compatible with Master/Sub Architecture v2.9.0 - set as module var for cust field mapping
Private masterProj As Boolean 'v3.0.0 stores master project status of active project based on subproject count
Private subP As SubProject 'v3.0.0 used to iterate through subprojects collection
Private subPID As Integer 'v3.0.0 used to temporarily store subproject ID
Private tempproj As Project 'v3.0.0 used to temporarily reference subprojects
Private firstTask As Boolean 'v3.0.0 used to track seed task for each path

Sub DrivingPaths()
'Primary analysis module that controls analysis
'workflow through Primary, Secondary and Tertiary
'driving paths.

    Dim t As Task 'Stores initial user selected task
    Dim tdp As TaskDependency
    Dim tdps As TaskDependencies
    Dim i As Integer 'Used to iterate through Primary/Secondary/Tertiary driver arrays
    Dim analysisTaskUID As String 'Stores user selected task for recall and selection after setting final view
    
    'Store users active project
    Set curProj = ActiveProject 'v2.9.0 get active project before displaying field selection form
    
    'v3.0.0 - check for subprojects
    If curProj.Subprojects.Count > 1 Then
        masterProj = True
    Else
        masterProj = False
    End If
    
    'used to avoid code break during intial error checks
    On Error Resume Next
    
    'Validate users selected view type
    If curProj.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
        MsgBox "Please select a View with a Task Table."
        curProj = Nothing
        Exit Sub
    End If
    
    'Validate users selected window pane - select the task table if not active
    If curProj.Application.ActiveWindow.ActivePane.Index <> 1 Then
        curProj.Application.ActiveWindow.TopPane.Activate
    End If
    
    'Exit if multiple tasks are selected
    If curProj.Application.ActiveSelection.Tasks.Count > 1 Then
        MsgBox "Select a single activity only."
        curProj = Nothing
        Exit Sub
    End If
    
    'store task of activeselection
    Set t = curProj.Application.ActiveCell.Task
    
    'Check for null task rows
    If t Is Nothing Then
        MsgBox "Select a task"
        curProj = Nothing
        Exit Sub
    End If
    
    'Avoid analyzing completed tasks
    If t.PercentComplete = 100 Then
        MsgBox "Select an incomplete task"
        curProj = Nothing
        Exit Sub
    End If
    
    'Avoid analysis on summary rows
    If t.Summary = True Then
        MsgBox "Select a non-summary task"
        curProj = Nothing
        Exit Sub
    End If
    
    'v2.9.0 Diplay Field Selection dialog
    Dim critPathFieldMapForm As cptCritPathFields_frm
    Set critPathFieldMapForm = New cptCritPathFields_frm
    
    With critPathFieldMapForm
    
        ReadCustomFields curProj
    
        For i = 1 To UBound(CustNumFields)
            .GroupField_Combobox.AddItem CustNumFields(i)
        Next i
        For i = 1 To UBound(CustTextFields)
            .GroupField_Combobox.AddItem CustTextFields(i)
            .PathField_Combobox.AddItem CustTextFields(i)
        Next i
        
        .Show
        
        If .Tag = "cancel" Then
            Set critPathFieldMapForm = Nothing
            Set curProj = Nothing
            Exit Sub
        End If
        
        'Hardcoded field requirements
        'v2.9.0 - get user field map
        CritField = .PathField_Combobox.Text
        GroupField = .GroupField_Combobox.Text
    
    End With
    
    'Suspend calculations and screen updating
    curProj.Application.Calculation = pjManual
    curProj.Application.ScreenUpdating = False
    
    On Error GoTo CleanUp
    
    '**********************************************
    'On Error GoTo 0 '*****used for debug only*****
    '**********************************************
    
    'v3.0.0 Assign Custom Field names and create lookup table for each subproject
    If masterProj = True Then
        For Each subP In curProj.Subprojects
            FileOpenEx subP.Path, True
            Set tempproj = ActiveProject
            SetGroupCPFieldLookupTable GroupField, CritField, tempproj
        Next subP
        curProj.Activate
    End If
    
    'v3.0.0 run no matter what the masterProj condition is
    'still need to update fields in Master Project file
    'in case tasks exist at top level
    SetGroupCPFieldLookupTable GroupField, CritField, curProj
    
    'Erase previous Crit and Group field values
    CleanCritFlag curProj
    
    'Erase any previously created/modified view elements
    CleanViews curProj
    
    'Initialize Analyzed Tasks Collection
    Set AnalyzedTasks = New Collection
    
    'Add selected task to Analyzed Tasks collection and store UID for later reference
    '**NOTE** in master project scenario, will present as master project unique for selected task
    AnalyzedTasks.Add t.UniqueID, t.UniqueID & "-" & t.UniqueID
    analysisTaskUID = t.UniqueID

    'Set default Float values
    tDrivingPaths.PrimaryFloat = 0
    tDrivingPaths.SecondaryFloat = 0
    tDrivingPaths.TertiaryFloat = 0
    tDrivingPaths.FourthFloat = 0
    tDrivingPaths.FifthFloat = 0
    
    'Now finding Primary Path
    tDrivingPaths.FindPrimary = True
    tDrivingPaths.FindSecondary = False
    tDrivingPaths.FindTertiary = False
    tDrivingPaths.FindFourth = False
    tDrivingPaths.FindFifth = False
    
    'Set default driver counts
    SecondaryDriverCount = 0
    TertiaryDriverCount = 0
    drivingTasksCount = 0
    FourthDriverCount = 0
    FifthDriverCount = 0
    
    '********************************
    '***Find Primary Driving Paths***
    '********************************
    
    'Store dependencies of user selected task
    Set tdps = t.TaskDependencies
    
    'Note that selected task is visible on paths 1,2,3
    t.SetField FieldNameToFieldConstant(CritField), "1,2,3"
    t.SetField FieldNameToFieldConstant(GroupField), "1"
    
    'Evlauate list of dependencies on selected analysis task
    For Each tdp In tdps
    
        'v3.0.0
        firstTask = True
    
        'evaluate task dependencies, add to analyzed tasks collection as needed, and review for criticality
        evaluateTaskDependencies tdp, t, curProj, AnalyzedTasks
        
    Next tdp 'Next user selected analysis task dependency
    
    '<---cpt:exit here for single driving path--->
    If singlePath Then GoTo ShowAndTell
    
    'Clear variables for re-use in evaluating secondary driver
    Set tdps = Nothing
    Set tdp = Nothing
    Set t = Nothing
    Set AnalyzedTasks = New Collection
    
    'Iterate through drivingtasks array to find next path driver
    FindNextDriver
    
    '**********************************
    '***Find Secondary Driving Paths***
    '**********************************
    
    'Find Secondary if one exists (driver count is greater than 0)
    If SecondaryDriverCount > 0 Then
    
        'Note that we are now evaluating the secondary driving path
        tDrivingPaths.FindSecondary = True
        
        'iterate through list of secondary drivers
        For i = 1 To SecondaryDriverCount
        
            'avoid evaluating outside of the array bounds
            If i > SecondaryDriverCount Then Exit For
            
            'store the current driving task
            Set t = curProj.Tasks.UniqueID(SecondaryDrivers(i))
            
            'add the driving task to the analyzed tasks collection
            AnalyzedTasks.Add t.UniqueID & "-" & t.UniqueID
            
            'If the task has not already been analyzed during previous path analysis,
            'set the Crit and Group Field values
            If t.GetField(FieldNameToFieldConstant(CritField)) = vbNullString Then
                With t
                    .SetField FieldNameToFieldConstant(CritField), "2"
                    .SetField FieldNameToFieldConstant(GroupField), "2"
                End With
            Else
            
                'If the task has already been analyzed during the previous path analysis,
                'append path value to the Crit and Group Fields
                If InStr(t.GetField(FieldNameToFieldConstant(CritField)), "2") = 0 Then
                    t.SetField FieldNameToFieldConstant(CritField), t.GetField(FieldNameToFieldConstant(CritField)) & ",2"
                End If
                
            End If
            
            'Store secondary driving task dependencies
            Set tdps = t.TaskDependencies
            
            'Evlauate list of dependencies on secondary driving task
            For Each tdp In tdps
            
                'evaluate task dependencies, add to analyzed tasks collection as needed, and review for criticality
                evaluateTaskDependencies tdp, t, curProj, AnalyzedTasks
                
            Next tdp 'Next secondary driver dependency
            
        Next i 'next Secondary Path Driver
        
    End If
    
    'Clear variables for re-use in evaluating secondary driver
    Set tdps = Nothing
    Set tdp = Nothing
    Set t = Nothing
    Set AnalyzedTasks = New Collection
    
    'Iterate through drivingtasks array to find next path driver
    FindNextDriver
    
    '*********************************
    '***Find Tertiary Driving Paths***
    '*********************************
    
    'Find tertiary if one exists (driver count is greater than 0)
    If TertiaryDriverCount > 0 Then
    
        'Note that we are now evaluating the tertiary driving path
        tDrivingPaths.FindTertiary = True
        
        'iterate through list of tertiary drivers
        For i = 1 To TertiaryDriverCount
        
            'avoid evaluating outside of the array bounds
            If i > TertiaryDriverCount Then Exit For
            
            'store the current driving task
            Set t = curProj.Tasks.UniqueID(TertiaryDrivers(i))
            
            'add the driving task to the analyzed tasks collection
            AnalyzedTasks.Add t.UniqueID & "-" & t.UniqueID
            
            'If the task has not already been analyzed during previous path analysis,
            'set the Crit and Group Field values
            If t.GetField(FieldNameToFieldConstant(CritField)) = vbNullString Then
                With t
                    .SetField FieldNameToFieldConstant(CritField), "3"
                    .SetField FieldNameToFieldConstant(GroupField), "3"
                End With
            Else
            
                'If the task has already been analyzed during the previous path analysis,
                'append path value to the Crit and Group Fields
                If InStr(t.GetField(FieldNameToFieldConstant(CritField)), "3") = 0 Then
                    t.SetField FieldNameToFieldConstant(CritField), t.GetField(FieldNameToFieldConstant(CritField)) & ",3"
                End If
                
            End If
            
            'Store tertiary driving task dependencies
            Set tdps = t.TaskDependencies
            
            'Evlauate list of dependencies on tertiary driving task
            For Each tdp In tdps
            
                'evaluate task dependencies, add to analyzed tasks collection as needed, and review for criticality
                evaluateTaskDependencies tdp, t, curProj, AnalyzedTasks
                
            Next tdp 'Next tertiary driver dependency
            
        Next i 'next Tertiary Path Driver
        
    End If
    
    'Clear variables for re-use in evaluating secondary driver
    Set tdps = Nothing
    Set tdp = Nothing
    Set t = Nothing
    Set AnalyzedTasks = New Collection
    
    'Iterate through drivingtasks array to find next path driver
    FindNextDriver
    
    '*********************************
    '***Find Fourth Driving Paths***
    '*********************************
    
    'Find fourth path if one exists (driver count is greater than 0)
    If FourthDriverCount > 0 Then
    
        'Note that we are now evaluating the fourth driving path
        tDrivingPaths.FindFourth = True
        
        'iterate through list of drivers
        For i = 1 To FourthDriverCount
        
            'avoid evaluating outside of the array bounds
            If i > FourthDriverCount Then Exit For
            
            'store the current driving task
            Set t = curProj.Tasks.UniqueID(FourthDrivers(i))
            
            'add the driving task to the analyzed tasks collection
            AnalyzedTasks.Add t.UniqueID & "-" & t.UniqueID
            
            'If the task has not already been analyzed during previous path analysis,
            'set the Crit and Group Field values
            If t.GetField(FieldNameToFieldConstant(CritField)) = vbNullString Then
                With t
                    .SetField FieldNameToFieldConstant(CritField), "4"
                    .SetField FieldNameToFieldConstant(GroupField), "4"
                End With
            Else
            
                'If the task has already been analyzed during the previous path analysis,
                'append path value to the Crit and Group Fields
                If InStr(t.GetField(FieldNameToFieldConstant(CritField)), "4") = 0 Then
                    t.SetField FieldNameToFieldConstant(CritField), t.GetField(FieldNameToFieldConstant(CritField)) & ",4"
                End If
                
            End If
            
            'Store fourth driving task dependencies
            Set tdps = t.TaskDependencies
            
            'Evlauate list of dependencies on fourth driving task
            For Each tdp In tdps
            
                'evaluate task dependencies, add to analyzed tasks collection as needed, and review for criticality
                evaluateTaskDependencies tdp, t, curProj, AnalyzedTasks
                
            Next tdp 'Next fourth driver dependency
            
        Next i 'next fourth Path Driver
        
    End If
    
    '*********************************
    '***Find Fourth Driving Paths***
    '*********************************
    
    'Clear variables for re-use in evaluating next driver
    Set tdps = Nothing
    Set tdp = Nothing
    Set t = Nothing
    Set AnalyzedTasks = New Collection
    
    'Iterate through drivingtasks array to find next path driver
    FindNextDriver
    
    'Find fifth driver if one exists (driver count is greater than 0)
    If FifthDriverCount > 0 Then
    
        'Note that we are now evaluating the fifth driving path
        tDrivingPaths.FindFifth = True
        
        'iterate through list of fifth drivers
        For i = 1 To FifthDriverCount
        
            'avoid evaluating outside of the array bounds
            If i > FifthDriverCount Then Exit For
            
            'store the current driving task
            Set t = curProj.Tasks.UniqueID(FifthDrivers(i))
            
            'add the driving task to the analyzed tasks collection
            AnalyzedTasks.Add t.UniqueID & "-" & t.UniqueID
            
            'If the task has not already been analyzed during previous path analysis,
            'set the Crit and Group Field values
            If t.GetField(FieldNameToFieldConstant(CritField)) = vbNullString Then
                With t
                    .SetField FieldNameToFieldConstant(CritField), "5"
                    .SetField FieldNameToFieldConstant(GroupField), "5"
                End With
            Else
            
                'If the task has already been analyzed during the previous path analysis,
                'append path value to the Crit and Group Fields
                If InStr(t.GetField(FieldNameToFieldConstant(CritField)), "5") = 0 Then
                    t.SetField FieldNameToFieldConstant(CritField), t.GetField(FieldNameToFieldConstant(CritField)) & ",5"
                End If
                
            End If
            
            'Store fifth driving task dependencies
            Set tdps = t.TaskDependencies
            
            'Evlauate list of dependencies on fifth driving task
            For Each tdp In tdps
            
                'evaluate task dependencies, add to analyzed tasks collection as needed, and review for criticality
                evaluateTaskDependencies tdp, t, curProj, AnalyzedTasks
                
            Next tdp 'Next fifth driver dependency
            
        Next i 'next fifth Path Driver
        
    End If
    
    
ShowAndTell:
    
    'Create and Apply the "ClearPlan Driving Path" Table, View, Group, and Filter
    SetupCPView GroupField, curProj, analysisTaskUID
    
CleanUp:

    'If error encountered, alert the user, otherwise notify of completion
    If err Then
        MsgBox "Error Encountered"
    Else
        If Not (export_to_PPT) Then MsgBox "Complete", vbOKOnly, "ClearPlan Critical Path Analyzer"
    End If

    'Clear variables
    Set tdps = Nothing
    Set tdp = Nothing
    Set t = Nothing
    Erase SecondaryDrivers, TertiaryDrivers, DrivingTasks
    Set AnalyzedTasks = Nothing
    SecondaryDriverCount = 0
    TertiaryDriverCount = 0
    drivingTasksCount = 0
    Set AnalyzedTasks = Nothing
    
    'Enable calculations and screenupdating
    curProj.Application.Calculation = pjAutomatic
    curProj.Application.ScreenUpdating = True
    
    'release project variable
    Set curProj = Nothing

End Sub

Private Sub evaluateTaskDependencies(ByVal tdp As TaskDependency, ByVal t As Task, ByVal curProj As Project, ByRef curAnalyzedTasks As Collection)
'Evaluate each task dependency, ignoring complete preds, then store as an analyzed relationship and evaluate criticality

    'v3.0.0 new variables
    Dim real_ToUID As Long
    Dim real_FromUID As Long
    Dim subIndex As Integer

    'v3.0.0 need to convert the
    If firstTask = True And masterProj = True Then
        firstTask = False
        If tdp.To.ExternalTask = True Then
            subIndex = get_subProj_index(curProj, tdp.To.Project)
            real_ToUID = get_tdp_MasterUID(tdp.To.UniqueID, subIndex)
        Else
            subIndex = get_subProj_index(curProj, curProj.Subprojects(tdp.To.Project).Path)
            real_ToUID = get_tdp_MasterUID(tdp.To.UniqueID, subIndex)
        End If
    Else
        real_ToUID = tdp.To.UniqueID
    End If
    
    'Only evaluate incomplete predecessors
    If real_ToUID = t.UniqueID And tdp.From.PercentComplete <> 100 Then
        'v3.0.0 account for master project condition
        If masterProj Then
        
            If tdp.To.ExternalTask = True Then
                subIndex = get_subProj_index(curProj, tdp.To.Project)
                real_ToUID = get_tdp_MasterUID(tdp.To.UniqueID, subIndex)
            Else
                subIndex = get_subProj_index(curProj, curProj.Subprojects(tdp.To.Project).Path)
                real_ToUID = get_tdp_MasterUID(tdp.To.UniqueID, subIndex)
            End If
            
            If tdp.From.ExternalTask = True Then
                subIndex = get_subProj_index(curProj, tdp.From.Project)
                real_FromUID = get_tdp_MasterUID(tdp.From.UniqueID, subIndex)
            Else
                subIndex = get_subProj_index(curProj, curProj.Subprojects(tdp.From.Project).Path)
                real_FromUID = get_tdp_MasterUID(tdp.From.UniqueID, subIndex)
            End If
            
        Else
        
            real_ToUID = tdp.To.UniqueID
            real_FromUID = tdp.From.UniqueID

        End If
        
        'Check dependency for existance in analyzed tasks collection
        If ExistsInCollection(curAnalyzedTasks, real_FromUID & "-" & real_ToUID) = False Then 'v3.0.0 updated with real UID for master projects
            'If dependency has not been analyzed, add to analyzed tasks collection
            curAnalyzedTasks.Add real_FromUID, real_FromUID & "-" & real_ToUID 'v3.0.0 updated with real uid for master projects
            'Calculate True Float value and evaluate against list of driving tasks
            CheckCritTask curProj, tdp
        End If
    End If
    
End Sub

Private Sub SetGroupCPFieldLookupTable(ByVal GroupField As String, ByVal CritField As String, ByVal currentProject As Project)
'Set Crit and Group field names, assign lookup table to Group Field
    
    'v3.0.0 remove crit field attributes
    currentProject.Application.CustomFieldPropertiesEx FieldID:=FieldNameToFieldConstant(CritField), Attribute:=pjFieldAttributeNone, SummaryCalc:=pjCalcNone, GraphicalIndicators:=False, AutomaticallyRolldownToAssn:=False
    'currentProject.Application.CustomFieldRename FieldID:=FieldNameToFieldConstant(CritField), NewName:="CP Driving Paths"
    
    'Setup Lookup Table Properties
    currentProject.Application.CustomFieldPropertiesEx FieldID:=FieldNameToFieldConstant(GroupField), Attribute:=pjFieldAttributeNone
    currentProject.Application.CustomOutlineCodeEditEx FieldID:=FieldNameToFieldConstant(GroupField), OnlyLookUpTableCodes:=True, OnlyLeaves:=False, LookupDefault:=False, SortOrder:=0
    currentProject.Application.CustomFieldPropertiesEx FieldID:=FieldNameToFieldConstant(GroupField), Attribute:=pjFieldAttributeValueList, SummaryCalc:=pjCalcNone, GraphicalIndicators:=False, AutomaticallyRolldownToAssn:=False
    'currentProject.Application.CustomFieldRename FieldID:=FieldNameToFieldConstant(GroupField), NewName:="CP Driving Path Group ID"
    
    'Assign Lookup Table Values
    currentProject.Application.CustomFieldValueListAdd FieldNameToFieldConstant(GroupField), "1", "Primary"
    currentProject.Application.CustomFieldValueListAdd FieldNameToFieldConstant(GroupField), "2", "Secondary"
    currentProject.Application.CustomFieldValueListAdd FieldNameToFieldConstant(GroupField), "3", "Tertiary"
    currentProject.Application.CustomFieldValueListAdd FieldNameToFieldConstant(GroupField), "4", "Quaternary"
    currentProject.Application.CustomFieldValueListAdd FieldNameToFieldConstant(GroupField), "5", "Quinary"
    currentProject.Application.CustomFieldValueListAdd FieldNameToFieldConstant(GroupField), "0", "Noncritical"


End Sub
Private Sub SetupCPView(ByVal GroupField As String, ByVal curProj As Project, ByVal tUID As String)
'Setup CP View with Table & Grouping by Path Value

    Dim t As Task 'used to store user selected anlaysis task
    
    'Create CP Driving Path Table
    curProj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, Create:=True, ShowAddNewColumn:=True, OverwriteExisting:=True, FieldName:="ID", Width:=5, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, LockFirstColumn:=True, ColumnPosition:=0
    
    'Add fields to CP Driving Path Table
    curProj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:="Unique ID", Width:=10, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=1, LockFirstColumn:=True
    curProj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:=GroupField, Title:="Driving Path", Width:=5, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=1
    curProj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:="Name", Width:=45, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=2
    curProj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:="Duration", Width:=10, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=3
    curProj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:="Start", Width:=15, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=4
    curProj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:="Finish", Width:=15, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=5
    curProj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:="Total Slack", Width:=10, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=6

    'Create CP Driving Path Filter
    curProj.Application.FilterEdit Name:="*ClearPlan Driving Path Filter", TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:=GroupField, test:="is greater than", Value:="0", ShowInMenu:=False, ShowSummaryTasks:=False
    
    'On Error Resume Next
    
    'Create CP Driving Path Group
    curProj.TaskGroups.Add Name:="*ClearPlan Driving Path Group", FieldName:=GroupField
    
    'Create CP Driving Path view if necessary
    curProj.Application.ViewEditSingle Name:="*ClearPlan Driving Path View", Create:=True, ShowInMenu:=True, Table:="*ClearPlan Driving Path Table", Filter:="*ClearPlan Driving Path Filter", Group:="*ClearPlan Driving Path Group"
    
    'Apply the CP Driving Path view
    curProj.Application.ViewApply Name:="*ClearPlan Driving Path View"
    
    'Sort the View by Finish, then by Duration to produce Waterfall Gantt
    curProj.Application.Sort key1:="Finish", Ascending1:=True, Key2:="Duration", Ascending2:=False, Outline:=False
    
    'Select all tasks and zoom the Gantt to display all tasks in view
    curProj.Application.SelectAll
    curProj.Application.ZoomTimescale Selection:=True
    
    curProj.Application.SelectRow 1
    
    'Iterate through each task in view and color the Gantt bars based on CP Group Code
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then 'Fix issue 44 for v2.8
            Select Case t.GetField(FieldNameToFieldConstant(GroupField))
            
                'v3.0.0 added consideration for master projects, which require targeted subproject edits via "ProjectName" variable
                Case "1"
                    If masterProj Then
                        t.Application.GanttBarFormatEx TaskID:=t.ID, GanttStyle:=1, StartColor:=192, MiddleColor:=192, EndColor:=192, projectName:=curProj.Subprojects(t.Project).Path
                    Else
                        t.Application.GanttBarFormatEx TaskID:=t.ID, GanttStyle:=1, StartColor:=192, MiddleColor:=192, EndColor:=192
                    End If
        
                Case "2"
                    If masterProj Then
                        t.Application.GanttBarFormatEx TaskID:=t.ID, GanttStyle:=1, StartColor:=3243501, MiddleColor:=3243501, EndColor:=3243501, projectName:=curProj.Subprojects(t.Project).Path
                    Else
                        t.Application.GanttBarFormatEx TaskID:=t.ID, GanttStyle:=1, StartColor:=3243501, MiddleColor:=3243501, EndColor:=3243501
                    End If
                    
                Case "3"
                    If masterProj Then
                        t.Application.GanttBarFormatEx TaskID:=t.ID, GanttStyle:=1, StartColor:=65535, MiddleColor:=65535, EndColor:=65535, projectName:=curProj.Subprojects(t.Project).Path
                    Else
                        t.Application.GanttBarFormatEx TaskID:=t.ID, GanttStyle:=1, StartColor:=65535, MiddleColor:=65535, EndColor:=65535
                    End If
                    
                Case "4"
                    If masterProj Then
                        t.Application.GanttBarFormatEx TaskID:=t.ID, GanttStyle:=1, StartColor:=11788485, MiddleColor:=11788485, EndColor:=11788485, projectName:=curProj.Subprojects(t.Project).Path
                    Else
                        t.Application.GanttBarFormatEx TaskID:=t.ID, GanttStyle:=1, StartColor:=11788485, MiddleColor:=11788485, EndColor:=11788485
                    End If
                    
                Case "5"
                    If masterProj Then
                        t.Application.GanttBarFormatEx TaskID:=t.ID, GanttStyle:=1, StartColor:=15189684, MiddleColor:=15189684, EndColor:=15189684, projectName:=curProj.Subprojects(t.Project).Path
                    Else
                        t.Application.GanttBarFormatEx TaskID:=t.ID, GanttStyle:=1, StartColor:=15189684, MiddleColor:=15189684, EndColor:=15189684
                    End If
                
                Case Else
            
            End Select
        End If
    
    Next t
    
    'select the users original analysis task
    curProj.Application.FindEx "UniqueID", "equals", tUID

End Sub
Private Sub CleanCritFlag(ByVal curProj As Project)
'Remove previous analysis values from the Crit and Group fields

    Dim t As Task 'store task var
    
    'iterate through every task in the project
    For Each t In curProj.Tasks
    
        If Not t Is Nothing Then 'Fix issue #44 for v2.8
        
            'Reset values
            t.SetField FieldNameToFieldConstant(CritField), vbNullString
            'v3.0.0
            If t.Summary = False Then t.SetField FieldNameToFieldConstant(GroupField), "0"
            
        End If
    Next t

End Sub

Private Sub CleanViews(ByVal curProj As Project)
'Iterate through all Views, Tables, Filters, and Groups
'Delete previously created CP View Elements to avoid user modification errors

    Dim cpView As View
    Dim allViews As Views
    Dim cpTable As Table
    Dim allTables As Tables
    Dim cpFilter As Filter
    Dim allFilters As Filters
    Dim cpGroup As Group
    Dim allGroups As Groups
    
    'Set vars
    Set allViews = curProj.Views
    Set allTables = curProj.TaskTables
    Set allFilters = curProj.TaskFilters
    Set allGroups = curProj.TaskGroups
    
    'If the CPCritPathView is active, choose a different view
    curProj.Application.ViewApply Name:="Gantt Chart"

    'Clean up Views
    For Each cpView In allViews
        If cpView.Name = "*ClearPlan Driving Path View" Then
            cpView.Delete
            Exit For
        End If
    Next cpView
    
    'Clean up Tables
    For Each cpTable In allTables
        If cpTable.Name = "*ClearPlan Driving Path Table" Then
            cpTable.Delete
            Exit For
        End If
    Next cpTable
    
    'Clean up Filters
    For Each cpFilter In allFilters
        If cpFilter.Name = "*ClearPlan Driving Path Filter" Then
            cpFilter.Delete
            Exit For
        End If
    Next cpFilter
    
    'Clean up Groups
    For Each cpGroup In allGroups
        If cpGroup.Name = "*ClearPlan Driving Path Group" Then
            cpGroup.Delete
            Exit For
        End If
    Next cpGroup

End Sub
Private Function alreadyFound(ByVal t As Task) As Boolean
'Check for existing values in the Crit Field - if found, task has been evaluated previously

    If t.GetField(FieldNameToFieldConstant(CritField)) <> vbNullString Then
        alreadyFound = True
    Else
        alreadyFound = False
    End If
    
End Function

Private Sub FindNextDriver()
'Iterate through Driving Tasks array to find driving tasks based on True Float value

    Dim i As Integer 'Counter used to iterate through DrivingTasks array
    Dim driverCount As Integer 'count of driving tasks found
    Dim driverFloat As Double 'float value of driving tasks

    'If no drivers were found, exit the subroutine
    If drivingTasksCount = 0 Then
        Exit Sub
    End If

    'Find Secondary Driving Task if the Find Secondary has not yet been set to True
    If tDrivingPaths.FindSecondary = False Then
        
        'Store default float and count values
        driverFloat = 0
        driverCount = 0
    
        'Iterate through Driving Tasks array and find the least float value
        For i = 1 To UBound(DrivingTasks)
        
            'store first float value, otherwise evaluate current float value against previously stored value
            If driverFloat = 0 And DrivingTasks(i).tFloat <> 0 Then
                driverFloat = DrivingTasks(i).tFloat
            Else
                With DrivingTasks(i)
                    If .tFloat < driverFloat And .tFloat <> 0 Then
                        driverFloat = .tFloat
                    End If
                End With
            End If
        Next i 'Next Driving Task
        
        'Find all drivers with similar float and store as parallel driving tasks
        If driverFloat <> 0 Then
            For i = 1 To UBound(DrivingTasks)
                With DrivingTasks(i)
                    If .tFloat = driverFloat Then
                        driverCount = driverCount + 1
                        ReDim Preserve SecondaryDrivers(1 To driverCount)
                        SecondaryDrivers(driverCount) = .UID
                    End If
                End With
            Next i 'Next Driving Task
        End If
        
        'Set Secondary Float value equal to the evaluated driving task float
        tDrivingPaths.SecondaryFloat = driverFloat
        
        'set secondary driver count
        SecondaryDriverCount = driverCount
        
    ElseIf tDrivingPaths.FindTertiary = False Then
    
        'Store default float and count values
        driverFloat = 0
        driverCount = 0
    
        'Iterate through Driving Tasks array and find the least float value
        For i = 1 To UBound(DrivingTasks)
        
            'store first float value, otherwise evaluate current float value against previously stored value
            If DrivingTasks(i).tFloat > tDrivingPaths.SecondaryFloat And driverFloat = 0 Then
                driverFloat = DrivingTasks(i).tFloat
            Else
                If DrivingTasks(i).tFloat > tDrivingPaths.SecondaryFloat And DrivingTasks(i).tFloat < driverFloat Then
                    driverFloat = DrivingTasks(i).tFloat
                End If
            End If
        Next i 'Next driving task
        
        'Find all drivers with similar float and store as parallel driving tasks
        If driverFloat <> 0 Then
            For i = 1 To UBound(DrivingTasks)
                With DrivingTasks(i)
                    If .tFloat = driverFloat Then
                        driverCount = driverCount + 1
                        ReDim Preserve TertiaryDrivers(1 To driverCount)
                        TertiaryDrivers(driverCount) = .UID
                    End If
                End With
            Next i 'Next Driving Task
        End If
        
        'Set Secondary Float value equal to the evaluated driving task float
        tDrivingPaths.TertiaryFloat = driverFloat
        
        'set secondary driver count
        TertiaryDriverCount = driverCount
        
    ElseIf tDrivingPaths.FindFourth = False Then
    
        'Store default float and count values
        driverFloat = 0
        driverCount = 0
    
        'Iterate through Driving Tasks array and find the least float value
        For i = 1 To UBound(DrivingTasks)
        
            'store first float value, otherwise evaluate current float value against previously stored value
            If DrivingTasks(i).tFloat > tDrivingPaths.FourthFloat And driverFloat = 0 Then
                driverFloat = DrivingTasks(i).tFloat
            Else
                If DrivingTasks(i).tFloat > tDrivingPaths.FourthFloat And DrivingTasks(i).tFloat < driverFloat Then
                    driverFloat = DrivingTasks(i).tFloat
                End If
            End If
        Next i 'Next driving task
        
        'Find all drivers with similar float and store as parallel driving tasks
        If driverFloat <> 0 Then
            For i = 1 To UBound(DrivingTasks)
                With DrivingTasks(i)
                    If .tFloat = driverFloat Then
                        driverCount = driverCount + 1
                        ReDim Preserve FourthDrivers(1 To driverCount)
                        FourthDrivers(driverCount) = .UID
                    End If
                End With
            Next i 'Next Driving Task
        End If
        
        'Set Secondary Float value equal to the evaluated driving task float
        tDrivingPaths.FourthFloat = driverFloat
        
        'set secondary driver count
        FourthDriverCount = driverCount
        
    ElseIf tDrivingPaths.FindFifth = False Then
    
        'Store default float and count values
        driverFloat = 0
        driverCount = 0
    
        'Iterate through Driving Tasks array and find the least float value
        For i = 1 To UBound(DrivingTasks)
        
            'store first float value, otherwise evaluate current float value against previously stored value
            If DrivingTasks(i).tFloat > tDrivingPaths.FifthFloat And driverFloat = 0 Then
                driverFloat = DrivingTasks(i).tFloat
            Else
                If DrivingTasks(i).tFloat > tDrivingPaths.FifthFloat And DrivingTasks(i).tFloat < driverFloat Then
                    driverFloat = DrivingTasks(i).tFloat
                End If
            End If
        Next i 'Next driving task
        
        'Find all drivers with similar float and store as parallel driving tasks
        If driverFloat <> 0 Then
            For i = 1 To UBound(DrivingTasks)
                With DrivingTasks(i)
                    If .tFloat = driverFloat Then
                        driverCount = driverCount + 1
                        ReDim Preserve FifthDrivers(1 To driverCount)
                        FifthDrivers(driverCount) = .UID
                    End If
                End With
            Next i 'Next Driving Task
        End If
        
        'Set Secondary Float value equal to the evaluated driving task float
        tDrivingPaths.FifthFloat = driverFloat
        
        'set secondary driver count
        FifthDriverCount = driverCount
    
    End If

End Sub

Private Function FindInArray(UID As String) As Variant
'Search DrivingTasks array for a task UID

    Dim i As Long 'counter to iterate through Driving Tasks
    
    For i = LBound(DrivingTasks) To UBound(DrivingTasks)
        If DrivingTasks(i).UID = UID Then
            FindInArray = i
            Exit Function
        End If
    Next i

    FindInArray = Null

End Function

Private Sub CheckCritTask(ByVal curProj As Project, ByVal tdp As TaskDependency)
'Compare current task dependency against full list of Driving Tasks and
'add-to/create/replace list of Path Drivers if critical

    Dim tdps As TaskDependencies 'store task dependencies
    Dim tdpI As TaskDependency 'store task dependency
    Dim tempFloat As Double 'tempFloat value used to compare float amongst all preds
    Dim i As Variant 'used to store unique ID of driving task if found in Driving Tasks array
    Dim predT As Task 'var to store pred task of evaluated dependency relationship
    Dim succT As Task 'var to store succ task of evaluated dependency relationship
    Dim predCritCoding As String 'var to store/modify existing Crit field values
    Dim subpIndex As Integer 'v3.0.0
    Dim realPredUID As Long 'v3.0.0
    Dim realSuccUID As Long 'v3.0.0
    
    'Assign the dependency predecessor task to predT var
    'v3.0.0 consider mast project condition
    If masterProj Then
        If tdp.From.ExternalTask = True Then
            subpIndex = get_subProj_index(curProj, tdp.From.Project)
            If subpIndex = 0 Then 'subproject is not present
                Exit Sub
            Else
                realPredUID = get_external_MasterUID(tdp.From, subpIndex)
                Set predT = curProj.Tasks.UniqueID(realPredUID)
            End If
                
        Else
            subpIndex = get_subProj_index(curProj, curProj.Subprojects(tdp.From.Project).Path)
            realPredUID = get_tdp_MasterUID(tdp.From.UniqueID, subpIndex)
            Set predT = curProj.Tasks.UniqueID(realPredUID)
        End If
    Else
        realPredUID = tdp.From.UniqueID
        Set predT = curProj.Tasks.UniqueID(tdp.From.UniqueID)
    End If
    
    
    'store predecessor task Crit path coding
    predCritCoding = predT.GetField(FieldNameToFieldConstant(CritField))
    
    'Assign the dependency successor task to the succT var
    'v3.0.0 consider master project condition - succ T will never be an external task
    If masterProj Then
        subpIndex = get_subProj_index(curProj, curProj.Subprojects(tdp.To.Project).Path)
        realSuccUID = get_tdp_MasterUID(tdp.To.UniqueID, subpIndex)
        Set succT = curProj.Tasks.UniqueID(realSuccUID)
    Else
        realSuccUID = tdp.To.UniqueID
        Set succT = curProj.Tasks.UniqueID(tdp.To.UniqueID)
    End If
    
    'get the TrueFloat of Dependency relationship
    tempFloat = TrueFloat(predT, succT, tdp.Type, tdp.Lag, tdp.LagType)

    'If not evaluating the last path, and the TrueFloat value is not 0
    'Evaluate total network float and store in Driving Tasks array
    If tDrivingPaths.FindFifth = False And tempFloat <> 0 Then
    
        'If other Driving Tasks have been found, Evaluate further
        If drivingTasksCount > 0 Then
            
            'Look for predecessor task in Driving Tasks Array
            i = FindInArray(CStr(realPredUID)) 'v3.0.0
    
            'If the task exists in the Driving Tasks array, evaluate further
            If Not IsNull(i) Then
            
                'if currently evaluating primary path, evaluate further
                If tDrivingPaths.FindSecondary = False Then
                
                    'if the dependency True Flaot is less than the previously stored float value
                    '(i.e. there are redundant links in the network), then store the lower float value
                    If tempFloat < DrivingTasks(i).tFloat Then
                        DrivingTasks(i).tFloat = tempFloat
                    End If
                Else 'if evaluating secondary path
                
                    'if the dependency float value + the previous path float is less then the
                    'previously stored float vlaue, then store the lower float value
                    If tDrivingPaths.FindTertiary = False And tempFloat + tDrivingPaths.SecondaryFloat < DrivingTasks(i).tFloat Then
                        DrivingTasks(i).tFloat = tempFloat + tDrivingPaths.SecondaryFloat
                    ElseIf tDrivingPaths.FindFourth = False And tempFloat + tDrivingPaths.TertiaryFloat < DrivingTasks(i).tFloat Then
                        DrivingTasks(i).tFloat = tempFloat + tDrivingPaths.TertiaryFloat
                    ElseIf tDrivingPaths.FindFifth = False And tempFloat + tDrivingPaths.FourthFloat < DrivingTasks(i).tFloat Then
                        DrivingTasks(i).tFloat = tempFloat + tDrivingPaths.FourthFloat
                    End If
                End If
            Else 'If the task does not exist in the Driving Tasks array
            
                'Add new driver to the driving task count and store in the array
                drivingTasksCount = drivingTasksCount + 1
                ReDim Preserve DrivingTasks(1 To drivingTasksCount)
                DrivingTasks(drivingTasksCount).UID = realPredUID 'v3.0.0
                
                'If evaluating the Primary Path, then store the float
                If tDrivingPaths.FindSecondary = False Then
                    DrivingTasks(drivingTasksCount).tFloat = tempFloat
                Else 'If evaluating secondary path, add float to the driving path network float value
                    DrivingTasks(drivingTasksCount).tFloat = tempFloat + tDrivingPaths.SecondaryFloat
                End If
            End If
        Else 'No other driving tasks found, this is the first driving task
            
            'Add the new driver to the driving tasks count and store in array
            drivingTasksCount = drivingTasksCount + 1
            ReDim DrivingTasks(1 To drivingTasksCount) 'removed Preserve - should not be neccessary when finding first driving task
            DrivingTasks(drivingTasksCount).UID = realPredUID 'v3.0.0
            
            'If evaluating the Primary Path, then store the float
            If tDrivingPaths.FindSecondary = False Then
                DrivingTasks(drivingTasksCount).tFloat = tempFloat
            Else 'If evaluating secondary path, add float to the driving path network float value
                DrivingTasks(drivingTasksCount).tFloat = tempFloat + tDrivingPaths.SecondaryFloat
            End If
        End If
    End If
    
    'Evaluate new driver if True Float is 0
    If tempFloat = 0 Then
        
        'If other drivers exist, and evaluating Primary or Secondary path, evaluate further
        If drivingTasksCount > 0 And tDrivingPaths.FindTertiary = False Then
        
            'Look for predecessor task in Driving Tasks Array
            i = FindInArray(CStr(realPredUID)) 'v3.0.0
    
            'If the task exists in the driving tasks array, update the float value
            If Not IsNull(i) Then
                DrivingTasks(i).tFloat = tempFloat
            Else 'If this is a new driver
            
                'Store the driving task in the Driving Tasks array
                drivingTasksCount = drivingTasksCount + 1
                ReDim Preserve DrivingTasks(1 To drivingTasksCount)
                With DrivingTasks(drivingTasksCount)
                    .UID = realPredUID 'v3.0.0
                    .tFloat = tempFloat
                End With
            End If
            
        Else 'If no other driving tasks exists and not evaluating the last path
            If tDrivingPaths.FindFifth = False Then

                'Store the new driving task
                drivingTasksCount = drivingTasksCount + 1
                ReDim DrivingTasks(1 To drivingTasksCount) 'removed Preserve - should not be neccessary when finding first driving task
                With DrivingTasks(drivingTasksCount)
                    .UID = realPredUID 'v3.0.0
                    .tFloat = tempFloat
                End With
            End If
        End If
    
        'If evaluating Primary Path, code the Crit and Group field values
        If tDrivingPaths.FindPrimary = True And tDrivingPaths.FindSecondary = False Then
            With predT
                .SetField FieldNameToFieldConstant(CritField), "1"
                .SetField FieldNameToFieldConstant(GroupField), "1"
            End With
        ElseIf tDrivingPaths.FindSecondary = True And tDrivingPaths.FindTertiary = False Then
            'If evaluating the secondary path, code the Crit and Group field values
            
            'If no existing code, then no need to concatenate
            If predCritCoding = vbNullString Then
                With predT
                    .SetField FieldNameToFieldConstant(CritField), "2"
                    .SetField FieldNameToFieldConstant(GroupField), "2"
                End With
            Else 'if existing code, then concatenate string
                If InStr(predCritCoding, "2") = 0 Then
                    predT.SetField FieldNameToFieldConstant(CritField), predCritCoding & ",2"
                End If
            End If
            
        ElseIf tDrivingPaths.FindTertiary = True And tDrivingPaths.FindFourth = False Then
            
            'If no existing code, then no need to concatenate
            If predCritCoding = vbNullString Then
                With predT
                    .SetField FieldNameToFieldConstant(CritField), "3"
                    .SetField FieldNameToFieldConstant(GroupField), "3"
                End With
            Else 'if existing code, then concatenate string
                If InStr(predCritCoding, "3") = 0 Then
                    predT.SetField FieldNameToFieldConstant(CritField), predCritCoding & ",3"
                End If
            End If
            
        ElseIf tDrivingPaths.FindFourth = True And tDrivingPaths.FindFifth = False Then
            
            'If no existing code, then no need to concatenate
            If predCritCoding = vbNullString Then
                With predT
                    .SetField FieldNameToFieldConstant(CritField), "4"
                    .SetField FieldNameToFieldConstant(GroupField), "4"
                End With
            Else 'if existing code, then concatenate string
                If InStr(predCritCoding, "4") = 0 Then
                    predT.SetField FieldNameToFieldConstant(CritField), predCritCoding & ",4"
                End If
            End If
            
        Else
        
            'If no existing code, then no need to concatenate
            If predCritCoding = vbNullString Then
                With predT
                    .SetField FieldNameToFieldConstant(CritField), "5"
                    .SetField FieldNameToFieldConstant(GroupField), "5"
                End With
            Else 'if existing code, then concatenate string
                If InStr(predCritCoding, "5") = 0 Then
                    predT.SetField FieldNameToFieldConstant(CritField), predCritCoding & ",5"
                End If
            End If
            
        End If
    
        'store dependecies of the currently evaluted dependency
        Set tdps = predT.TaskDependencies
        
        'Iterate through the dependencies of the dependency
        For Each tdpI In tdps
        
            'evaluate task dependencies, add to analyzed tasks collection as needed, and review for criticality
            evaluateTaskDependencies tdpI, predT, curProj, AnalyzedTasks

        Next tdpI 'Next dependency of the currently evaluated dependency
    End If
        
End Sub

Private Function TrueFloat(ByVal tPred As Task, ByVal tSucc As Task, ByVal dType As Integer, ByVal dLag As Double, dlagtype As Integer) As Double
'Find True Float Value
'True Float is the dependency level 'free float' value,
'taking into consideration all duration types (including eDays),
'task calendars, leads/lags, etc

    Dim pDate As Date 'Store predecessor date (start or fin depending on link type)
    Dim sDate As Date 'Store successor date (start or fin depending on link type)
    Dim sCalObj As Calendar 'Store successor task calendar or project calendar if task cal = N/A
    Dim pCalObj As Calendar 'Store predecessor task calendar or project calendar if task cal = N/A
    Dim tempFloat As Double 'store True Float for function return
    Dim subpIndex As Integer 'v3.0.0
    
    'If pred task has a task calendar, store
    If tPred.Calendar <> "None" Then
        Set pCalObj = tPred.CalendarObject
    Else 'If no task calendar, store project cal
        'v3.0.0 consider master project condition
        If masterProj = True Then
            If tPred.Project = curProj.Tasks.UniqueID(0).Project Then 'task is in master project
                Set pCalObj = curProj.Calendar
            Else
                Set pCalObj = curProj.Subprojects(tPred.Project).SourceProject.Calendar
            End If
        Else
            Set pCalObj = ActiveProject.Calendar
        End If
    End If
    
    'If succ task has a task calendar, store
    If tSucc.Calendar <> "None" Then
        Set sCalObj = tSucc.CalendarObject
    Else 'If no task calendar, store project cal
        'v3.0.0 consider master project condition
        If masterProj = True Then
            If tSucc.Project = curProj.Tasks.UniqueID(0).Project Then 'task is in master project
                Set sCalObj = curProj.Calendar
            Else
                Set sCalObj = curProj.Subprojects(tSucc.Project).SourceProject.Calendar
            End If
        Else
            Set sCalObj = ActiveProject.Calendar
        End If
    End If
    
    'if dependency lag is greater than or equal to 0
    If dLag >= 0 Then
    
        'evaluate the depenency type
        Select Case dType
            
            Case 0 'Finish to Finish
                
                'Set predecessor date equal to the pred Finish date plus the lag value (considering any succ task cal)
                pDate = Application.DateAdd(tPred.Finish, Application.DurationFormat(dLag, dlagtype), sCalObj)
                
                'successor date equals the succ finish
                'includes leveling delay test v2.9.0
                
                If tSucc.LevelingDelay > 0 Then
                    
                    sDate = tSucc.EarlyFinish
                    
                Else
                
                    sDate = tSucc.Finish
                
                End If
            
            Case 1 'Finish to Start
            
                'Set predecessor date equal to the pred Finish date plus the lag value (considering any succ task cal)
                pDate = Application.DateAdd(tPred.Finish, Application.DurationFormat(dLag, dlagtype), sCalObj)
                
                'successor date equals the succ start
                'includes leveling delay test v2.9.0
                
                If tSucc.LevelingDelay > 0 Then
                
                    sDate = tSucc.EarlyStart
                
                Else
                
                    sDate = tSucc.Start
                
                End If
            
            Case 2 'Start to Finish
            
                'Set predecessor date equal to the pred Start date plus the lag value (considering any succ task cal)
                pDate = Application.DateAdd(tPred.Start, Application.DurationFormat(dLag, dlagtype), sCalObj)
                
                'successor date equals the succ finish
                'includes leveling delay test v2.9.0
                
                If tSucc.LevelingDelay > 0 Then
                
                    sDate = tSucc.EarlyFinish
                
                Else
                
                    sDate = tSucc.Finish
                
                End If
            
            Case 3 'Start to Start
            
                'Set predecessor date equal to the pred start date plus the lag value (considering any succ task cal)
                pDate = Application.DateAdd(tPred.Start, Application.DurationFormat(dLag, dlagtype), sCalObj)
                
                'successor date equals the succ start
                'includes leveling delay test v2.9.0
                
                If tSucc.LevelingDelay > 0 Then
                
                    sDate = tSucc.EarlyStart
                
                Else
                
                    sDate = tSucc.Start
                
                End If
                
            Case Else
        End Select
    
    'if lag is less than 0 (lead)
    ElseIf dLag < 0 Then
    
        'evaluate the dependency type
        Select Case dType
            
            Case 0 'Finish to Finish
            
                'pred date equals the pred finish
                pDate = tPred.Finish
                
                'succ date equals the succ finish plus the absolute value of the lag (considering any succ task cal)
                'adding the lead to the succ date is the same as subtracting it from the pred date (DateAdd cannot handle negative values)
                
                'includes leveling delay test v2.9.0
                
                If tSucc.LevelingDelay > 0 Then
                
                    sDate = Application.DateAdd(tSucc.EarlyFinish, Application.DurationFormat(Abs(dLag), dlagtype), sCalObj)
                
                Else
                
                    sDate = Application.DateAdd(tSucc.Finish, Application.DurationFormat(Abs(dLag), dlagtype), sCalObj)
            
                End If
            
            Case 1 'Finish to Start
            
                'pred date equals the pred finish
                pDate = tPred.Finish
                
                'succ date equals the succ start plus the absolute value of the lag (considering any succ task cal)
                'adding the lead to the succ date is the same as subtracting it from the pred date (DateAdd cannot handle negative values)
                
                'includes leveling delay test v2.9.0
                
                If tSucc.LevelingDelay > 0 Then
                
                    sDate = Application.DateAdd(tSucc.EarlyStart, Application.DurationFormat(Abs(dLag), dlagtype), sCalObj)
                
                Else
                
                    sDate = Application.DateAdd(tSucc.Start, Application.DurationFormat(Abs(dLag), dlagtype), sCalObj)
            
                End If
            
            Case 2 'Start to Finish
            
                'pred date equals the pred start
                pDate = tPred.Start
                
                'succ date equals the succ finish plus the absolute value of the lag (considering any succ task cal)
                'adding the lead to the succ date is the same as subtracting it from the pred date (DateAdd cannot handle negative values)
                
                'includes leveling delay test v2.9.0
                
                If tSucc.LevelingDelay > 0 Then
                
                    sDate = Application.DateAdd(tSucc.Finish, Application.DurationFormat(Abs(dLag), dlagtype), sCalObj)
                
                Else
                
                    sDate = Application.DateAdd(tSucc.Finish, Application.DurationFormat(Abs(dLag), dlagtype), sCalObj)
            
                End If
            
            Case 3 'Start to Start
            
                'pred date equals the pred start
                pDate = tPred.Start
                
                'succ date equals the succ start plus the absolute value of the lag (considering any succ task cal)
                'adding the lead to the succ date is the same as subtracting it from the pred date (DateAdd cannot handle negative values)
                
                'includes leveling delay test v2.9.0
                
                If tSucc.LevelingDelay > 0 Then
                
                    sDate = Application.DateAdd(tSucc.EarlyStart, Application.DurationFormat(Abs(dLag), dlagtype), sCalObj)
                
                Else
                
                    sDate = Application.DateAdd(tSucc.Start, Application.DurationFormat(Abs(dLag), dlagtype), sCalObj)
                
                End If
                
            Case Else
        End Select
    End If
    
    'v2.8.2 check for edays
    If Left(GetLettersOnly(tPred.DurationText), 1) <> "e" Then
    
        'no edays; subtract the pred date from the succ date, using the pred calendar, to get the True Float value
        tempFloat = Application.DateDifference(pDate, sDate, pCalObj)
        
    Else
    
        'using edays; calculate date diff in minutes
        tempFloat = DateDiff("n", pDate, sDate)
    
    End If
    
    'Return the True Float value
    TrueFloat = tempFloat

End Function

Public Function ExistsInCollection(ByVal col As Collection, ByVal key As Variant) As Boolean
'Check for task dependency relationship in the analyzed tasks collection

    Dim f As Boolean 'stores boolean value 'True' if relationship exists in the collection
    
    'If error encountered, value does not exist in the collection
    On Error GoTo err
    
    f = IsObject(col.Item(key)) 'Store found item; if not found, will produce error
    ExistsInCollection = True 'Set True
    Exit Function
err: 'If error encountered, item does not exist - return "False" boolean vlaue
    ExistsInCollection = False
End Function

Function GetLettersOnly(str As String) As String
'v2.8.2 - strip out non-alpha characters from input string
'used to evaluate task duration text for elapsed day prefix "e"

    Dim i As Long, letters As String, letter As String

    letters = vbNullString

    For i = 1 To Len(str)
        letter = VBA.Mid$(str, i, 1)

        If Asc(LCase(letter)) >= 97 And Asc(LCase(letter)) <= 122 Then
            letters = letters + letter
        End If
    Next
    GetLettersOnly = letters
End Function

Private Sub ReadCustomFields(ByVal curProj As Project)
'v2.9.0 - added to allow user selection of custom fields

    Dim i As Integer

    'Read local Custom Text Fields
    For i = 1 To 30

        If Len(curProj.Application.CustomFieldGetName(FieldNameToFieldConstant("Text" & i))) > 0 Then
            ReDim Preserve CustTextFields(1 To i)
            CustTextFields(i) = curProj.Application.CustomFieldGetName(FieldNameToFieldConstant("Text" & i))
        Else
            ReDim Preserve CustTextFields(1 To i)
            CustTextFields(i) = "Text" & i
        End If

    Next i

    'Read local Custom Number Fields
    For i = 1 To 20

        If Len(curProj.Application.CustomFieldGetName(FieldNameToFieldConstant("Number" & i))) > 0 Then
            ReDim Preserve CustNumFields(1 To i)
            CustNumFields(i) = curProj.Application.CustomFieldGetName(FieldNameToFieldConstant("Number" & i))
        Else
            ReDim Preserve CustNumFields(1 To i)
            CustNumFields(i) = "Number" & i
        End If

    Next i


End Sub

Function get_subProj_index(ByVal masterProj As Project, ByVal subprojectFilename As String) As Integer
'v3.0.0 gets the subproject index
'used for calculating the Master Project UID

    Dim subP As SubProject
    
    For Each subP In masterProj.Subprojects
    
        If subP.Path = subprojectFilename Then
        
            get_subProj_index = masterProj.Subprojects(subP.Index).InsertedProjectSummary.UniqueID
            Exit Function
        
        End If
    
    Next subP
    
    get_subProj_index = 0

End Function

Function get_tdp_MasterUID(ByVal subP_UID As Long, ByVal subP_Index As Integer) As Long
'v3.0.0 convert subproject format UID to master project uid format
    
    If subP_Index = 0 Then
        get_tdp_MasterUID = subP_UID
    Else
        get_tdp_MasterUID = subP_UID + (4194304 * (subP_Index + 1))
    End If
    Exit Function
    
End Function

Function get_external_MasterUID(ByVal subP_Task As Task, ByVal subP_Index As Integer) As Long
'v3.0.0 get corresponding subproject UID for external reference task

    get_external_MasterUID = subP_Task.GetField(185073906) Mod 4194304 + (4194304 * (subP_Index + 1))
    Exit Function

End Function
