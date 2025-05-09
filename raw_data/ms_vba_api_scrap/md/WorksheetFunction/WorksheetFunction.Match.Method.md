# WorksheetFunction Match Method

## Business Description
Returns the relative position of an item in an array that matches a specified value in a specified order. Use MATCH instead of one of the LOOKUP functions when you need the position of an item in a range instead of the item itself.

## Behavior
Returns the relative position of an item in an array that matches a specified value in a specified order. Use MATCH instead of one of the LOOKUP functions when you need the position of an item in a range instead of the item itself.

## Example Usage
```vba
Sub HighlightMatches()
    Application.ScreenUpdating = False
    
    'Declare variables
    Dim var As Variant, iSheet As Integer, iRow As Long, iRowL As Long, bln As Boolean
       
       'Set up the count as the number of filled rows in the first column of Sheet1.
       iRowL = Cells(Rows.Count, 1).End(xlUp).Row
       
       'Cycle through all the cells in that column:
       For iRow = 1 To iRowL
          'For every cell that is not empty, search through the first column in each worksheet in the
          'workbook for a value that matches that cell value.

          If Not IsEmpty(Cells(iRow, 1)) Then
             For iSheet = ActiveSheet.Index + 1 To Worksheets.Count
                bln = False
                var = Application.Match(Cells(iRow, 1).Value, Worksheets(iSheet).Columns(1), 0)
                
                'If you find a matching value, indicate success by setting bln to true and exit the loop;
                'otherwise, continue searching until you reach the end of the workbook.
                If Not IsError(var) Then
                   bln = True
                   Exit For
                End If
             Next iSheet
          End If
          
          'If you do not find a matching value, do not bold the value in the original list;
          'if you do find a value, bold it.
          If bln = False Then
             Cells(iRow, 1).Font.Bold = False
             Else
             Cells(iRow, 1).Font.Bold = True
          End If
       Next iRow
    Application.ScreenUpdating = True
End SubSub HighlightMatches()

Application.ScreenUpdating = False

'Declare variables
Dim var As Variant, iSheet As Integer, iRow As Long, iRowL As Long, bln As Boolean
   
   'Set up the count as the number of filled rows in the first column of Sheet1.
   iRowL = Cells(Rows.Count, 1).End(xlUp).Row
   
   'Cycle through all the cells in that column:
   For iRow = 1 To iRowL
      'For every cell that is not empty, search through all the columns in all the worksheets in the
      'workbook for a value that matches that cell value.
      If Not IsEmpty(Cells(iRow, 1)) Then
         For iSheet = ActiveSheet.Index + 1 To Worksheets.Count
            bln = False
            var = Application.Match(Cells(iRow, 1).Value, Worksheets(iSheet).Columns(1), 0)
            
            'If you find a matching value, indicate success by setting bln to true and exit the loop;
            'otherwise, continue searching until you reach the end of the workbook.
            If Not IsError(var) Then
               bln = True
               Exit For
            End If
         Next iSheet
      End If
      
      'If you do not find a matching value, do not bold the value in the original list;
      'if you do find a value, bold it.
      If bln = False Then
         Cells(iRow, 1).Font.Bold = False
         Else
         Cells(iRow, 1).Font.Bold = True
      End If
   Next iRow
Application.ScreenUpdating = True
End Sub
```