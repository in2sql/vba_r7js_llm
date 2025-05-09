Worksheets("Sheet1").Activate 
Set bigRange = Application.Union(Range("Range1"), Range("Range2")) 
bigRange.Formula = "=RAND()"