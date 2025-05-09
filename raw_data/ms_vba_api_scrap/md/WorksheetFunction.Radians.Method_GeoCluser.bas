Attribute VB_Name = "Module1"
Sub Cluster()

'initialize the original string starting point
n = 3

'start a loop that goes through all the geographic data points in the original array
Do

'define the original latitude and longitude
Lat1 = Cells(n, 10)
Long1 = Cells(n, 11)

'initialize the starting point of the custered array
    m = 2
    
'start a loop to check all the clustered points
    Do
'define the cluser coordinates
    Lat2 = Cells(m, 12)
    Long2 = Cells(m, 13)
'compute the distance to a point in the cluster
    Distance = Application.WorksheetFunction.Acos(Cos(Application.WorksheetFunction.Radians(90 - Lat1)) * Cos(Application.WorksheetFunction.Radians(90 - Lat2)) + Sin(Application.WorksheetFunction.Radians(90 - Lat1)) * Sin(Application.WorksheetFunction.Radians(90 - Lat2)) * Cos(Application.WorksheetFunction.Radians(Long1 - Long2))) * 6371
'if the any point from the original array is close to a point in the cluster, exit the loop, otherwise check the next cluser point
'repeat the process until all the points in the cluster array have been checked
        If Distance < 0.1 Then
        Cells(m, 14) = Cells(m, 14) + Cells(n, 8)
        Cells(m, 15) = Cells(m, 15) & "|" & Cells(n, 7)
        Exit Do
        Else
        End If
        m = m + 1
    Loop While Cells(m, 12) <> ""
    
'if all the cells have been checked (hence the next spot is empty), make a new record
If Cells(m, 12) = "" Then
'record latitude
Cells(m, 12) = Cells(n, 10)
'record longitude
Cells(m, 13) = Cells(n, 11)
'add a new record to the array of all points in the cluster
Cells(m, 14) = Cells(n, 8)
'add the addtional frequency associated with the new point to the cluster frequency
Cells(m, 15) = Cells(n, 7)
Else
End If

n = n + 1

'perform the check for all original points in the list (21365)
Loop While n < 21365
        
        
End Sub
