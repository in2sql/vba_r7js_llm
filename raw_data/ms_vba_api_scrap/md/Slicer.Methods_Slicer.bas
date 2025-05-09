Attribute VB_Name = "Slicer"
Public Sub fexTimelineSlicerTimeSetting(slicerName As String, startDate As Date, endDate As Date)

    ActiveWorkbook.SlicerCaches(slicerName).TimelineState. _
    SetFilterDateRange startDate, endDate

End Sub


