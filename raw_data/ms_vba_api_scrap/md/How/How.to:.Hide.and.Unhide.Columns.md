# How to: Hide and Unhide Columns

## Business Description
This example finds all the cells in the first four columns that have a constant 

## Behavior
This example finds all the cells in the first four columns that have a constant "X" in them and hides the column that contains the X.

## Example Usage
```vba
Sub Hide_Columns()

    'Excel objects.
    Dim m_wbBook As Workbook
    Dim m_wsSheet As Worksheet
    Dim m_rnCheck As Range
    Dim m_rnFind As Range
    Dim m_stAddress As String

    'Initialize the Excel objects.
    Set m_wbBook = ThisWorkbook
    Set m_wsSheet = m_wbBook.Worksheets("Sheet1")
    
    'Search the four columns for any constants.
    Set m_rnCheck = m_wsSheet.Range("A1:D1").SpecialCells(xlCellTypeConstants)
    
    'Retrieve all columns that contain an X. If there is at least one, begin the DO/WHILE loop.
    With m_rnCheck
        Set m_rnFind = .Find(What:="X")
        If Not m_rnFind Is Nothing Then
            m_stAddress = m_rnFind.Address
             
            'Hide the column, and then find the next X.
            Do
                m_rnFind.EntireColumn.Hidden = True
                Set m_rnFind = .FindNext(m_rnFind)
            Loop While Not m_rnFind Is Nothing And m_rnFind.Address <> m_stAddress
        End If
    End With

End Sub
```