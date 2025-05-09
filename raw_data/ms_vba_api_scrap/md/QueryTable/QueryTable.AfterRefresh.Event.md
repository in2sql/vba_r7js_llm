# QueryTable AfterRefresh Event

## Business Description
Occurs after a query is completed or canceled.

## Behavior
Occurs after a query is completed or canceled.

## Example Usage
```vba
Private Sub QueryTable_AfterRefresh(Success As Boolean) 
 If Success Then 
 ' Query completed successfully 
 Else 
 ' Query failed or was cancelled 
 End If 
End Sub
```