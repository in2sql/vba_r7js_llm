# How to: Create an HTML File with a Table of Contents based on Cell Data

## Business Description
This code example shows how to take data from a worksheet and create a table of contents in an HTML file. The worksheet should have data in columns A, B, and C that correspond to the first, second, and third levels of the table of contents hierarchy.

## Behavior
This code example shows how to take data from a worksheet and create a table of contents in an HTML file. The worksheet should have data in columns A, B, and C that correspond to the first, second, and third levels of the table of contents hierarchy. The HTML file is stored in the same working folder as the active workbook.

## Example Usage
```vba
Sub CreateHTML()
   'Define your variables.
   Dim iRow As Long
   Dim iStage As Integer
   Dim iCounter As Integer
   Dim iPage As Integer
   
   'Create an .htm file in the same directory as your active workbook.
   Dim sFile As String
   sFile = ActiveWorkbook.Path & "\test.htm"
   Close
   
   'Open up the temp HTML file and format the header.
   Open sFile For Output As #1
   Print #1, "<html>"
   Print #1, "<head>"
   Print #1, "<style type=""text/css"">"
   Print #1, "  body { font-size:12px;font-family:tahoma } "
   Print #1, "</style>"
   Print #1, "</head>"
   Print #1, "<body>"
   
   'Start on the 2nd row to avoid the header.
   iRow = 2
   
   'Translate the first column of the table into the first level of the hierarchy.
   Do While WorksheetFunction.CountA(Rows(iRow)) > 0
      If Not IsEmpty(Cells(iRow, 1)) Then
         For iCounter = 1 To iStage
            Print #1, "</ul>"
            iStage = iStage - 1
         Next iCounter
         Print #1, "<ul>"
         Print #1, "<li><a href=""" & iPage & ".html"">" & Cells(iRow, 1).Value & "</a>"
         iPage = iPage + 1
         If iStage < 1 Then
            iStage = iStage + 1
         End If
      End If
      
    'Translate the second column of the table into the second level of the hierarchy.
      If Not IsEmpty(Cells(iRow, 2)) Then
         For iCounter = 2 To iStage
            Print #1, "</ul>"
            iStage = iStage - 1
         Next iCounter
         Print #1, "<ul>"
         Print #1, "<li><a href=""" & iPage & ".html"">" & Cells(iRow, 2).Value & "</a>"
         iPage = iPage + 1
         If iStage < 2 Then
            iStage = iStage + 1
         End If
      End If
      
      'Translate the third column of the table into the third level of the hierarchy.
      If Not IsEmpty(Cells(iRow, 3)) Then
         If iStage < 3 Then
            Print #1, "<ul>"
         End If
         Print #1, "<li><a href=""" & iPage & ".html"">" & Cells(iRow, 3).Value & "</a>"
         iPage = iPage + 1
         If iStage < 3 Then
            iStage = iStage + 1
         End If
      End If
      iRow = iRow + 1
   Loop
   
   'Add ending HTML tags
   For iCounter = 2 To iStage
      Print #1, "    </ul>"
      iStage = iStage - 1
   Next iCounter
   Print #1, "</body>"
   Print #1, "</html>"
   Close
   Shell "hh " & vbLf & sFile, vbMaximizedFocus
End Sub
```