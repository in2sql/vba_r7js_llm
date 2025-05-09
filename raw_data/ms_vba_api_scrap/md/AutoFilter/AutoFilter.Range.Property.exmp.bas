Workbooks(1).Activate 
Set hr = ActiveSheet.Hyperlinks(1).Range 
ActiveWindow.ScrollRow = hr.Row 
ActiveWindow.ScrollColumn = hr.Column