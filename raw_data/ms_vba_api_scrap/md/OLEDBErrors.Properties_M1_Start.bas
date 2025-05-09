Attribute VB_Name = "M1_Start"

'|***********************************************[ MAXVAK TOOL / Module: M1_Start ]***********************************************|
'|                                                                                                                                |
'|                                               [[[    Author: Nils Kuppen    ]]]                                                |
'|                                               [[[       For: JUMBO.com      ]]]                                                |
'|                                               [[[ EFC Den Bosch \ SiSu \ VO ]]]                                                |
'|                                                                                                                                |
'|                  This module contains all macro's that are indirectly triggered by buttons on the "START" tab                  |
'|                                                                                                                                |
'|                      Relative file path: \\Code\VBA\M1_Start.bas                                                               |
'|                      Updated by fn: ZZ99.update_vba                                                                            |
'|                      Triggered by: config.xml\config\update\vba = true                                                         |
'|                                                                                                                                |
'|***********************************************************[ (c) 2024 ]*********************************************************|

    '// Sub refrsh_tbl_kpi( )
    '// Update KPI Table
    '// Writes steps to log
    '// Export to SharePoint

Sub rfrsh_tbl_kpi()

    Dim r As Range: Set r = Range("kpi_log")
    Dim rDt As Range: Set rDt = Range("kpi_dt")
    Dim s As String

    Range("errchck").ClearContents

            dt = Format(Mid(rDt, 10), "ww", 2, 2)               '// Week number of last KPI export
            wk = Format(Date, "ww", 2, 2)                       '// Current week number
            pdf = Format(FileDateTime(pdfRapport), "ww", 2, 2)  '// Date of last management report

        If wk <> pdf Then   '// Check for new management report
                MsgBox "Management Rapportage niet beschikbaar. Probeer later opnieuw.", _
                        vbOKOnly + vbCritical, "Export KPI"
                GoTo Eind
        End If

        If dt = wk Then     '// Check if KPI was exported in current week
            Result = MsgBox("KPI is deze week al geexporteerd. Opnieuw?", vbYesNo + vbQuestion, "Export KPI")
            If Result = vbNo Then GoTo Eind
        End If

            dt = Time
            i% = 1
        
    On Error GoTo errHndlr

            r.ClearContents

        'If Mid(rDt, 10) < Date Then tblRfrsh Blad1, 0, r, i, "Maxvak tabel ververst"
            
                    tblRfrsh Blad2, 0, r, i, "KPI ververst"
                    tblRfrsh Blad1, 0, r, i, "Check compleet"
                    
                    If Range("UPDATE_1") = False Then

            rDt = "laatste: " & Date '// Set current date for KPI export

        If Range("errchck") = "" Then
            If ThisWorkbook.ReadOnly = False Then   '// Do not save or export in Read-Only mode
                    s = "Bestand opgeslagen"
                    ThisWorkbook.Save
                    msg = "KPI is ververst en bestand is opgeslagen"
                    fLog r, i, s
                
                    s = "Export SharePoint succesvol"
                    Application.ScreenUpdating = False
                    M2_ImportExport.export_KPI_sp
                    'M2_ImportExport.export_sharepoint
                    Application.ScreenUpdating = True
                    DoEvents
                    fLog r, i, s
            Else:
                    s = "Bestand alleen-lezen"
                    msg = "KPI is ververst en bestand is NIET opgeslagen"
                    fLog r, i, s
            End If
        End If
        
            'rDt = "laatste: " & Date '// Set current date for KPI export

        MsgBox msg & vbCrLf & vbCrLf & "Tijd: " & Format(Time - dt, "hh:mm:ss"), vbOKOnly + vbInformation, "Klaar!"

Eind:

        Blad1.Select

Exit Sub

errHndlr:

        s = Err.Description
        Err.Clear
        Resume Next

End Sub


    '// tblRfrsh( Sheet, Table type, Range, Index, String )
    '// Refresh query or pivot tables and log

Sub tblRfrsh(ByVal ws As Worksheet, ByVal x%, r As Range, i%, ByVal t$)

        On Error GoTo errHndlr

    Application.ScreenUpdating = False
    Application.DisplayAlerts = True

                s$ = t

        Select Case x
                        Case 0: ws.ListObjects(1).Refresh
                        Case 1: ws.PivotTables(1).RefreshTable
        End Select

                fLog r, i, s

    Application.DisplayAlerts = False
    Application.ScreenUpdating = True
    DoEvents

    '//extra error handeling specifically for PQ errors:
    If Application.OLEDBErrors.Count + Application.ODBCErrors.Count > 0 Then

        Range("errchck") = 1
        
        j = 1: strErr = vbCrLf & vbCrLf
        For Each pqErr In Application.OLEDBErrors
            strErr = strErr & j & " : " & pqErr.ErrorString & vbCrLf
            j = j + 1
        Next pqErr
        
        j = 1
        For Each orErr In Application.ODBCErrors
            strErr = strErr & j & " : " & pqErr.ErrorString & vbCrLf
            j = j + 1
        Next orErr
    
        MsgBox "Power Query errors: " & strErr
    
    End If


Exit Sub

errHndlr:

        s = Err.Description
        Err.Clear
        Resume Next

End Sub

    '// Sub fLog( Range, Index, String )
    '// Write updates/errors to log range

Sub fLog(ByRef r As Range, ByRef i%, ByRef s$)

        r.Cells(i, 1) = s
        i = i + 1

End Sub
