Attribute VB_Name = "SmartArtFont"
Option Explicit

Sub SmartArtFont()
    
    Dim sld As Slide
    Dim shp As Shape
    Dim subshp As Shape
    Dim i As Long
    
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
        
    If shp.Type = msoSmartArt Then
      For i = 1 To shp.SmartArt.AllNodes.Count
          shp.SmartArt.AllNodes(i).TextFrame2.TextRange.Font.Name = "UULA Sans"
          shp.SmartArt.AllNodes(i).TextFrame2.TextRange.Font.NameComplexScript = "UULA Sans"
      Next i
   End If
   
   If shp.Type = msoChart Then
      shp.Chart.ChartArea.Format.TextFrame2.TextRange.Font.Name = "UULA Sans"
      shp.Chart.ChartArea.Format.TextFrame2.TextRange.Font.NameComplexScript = "UULA Sans"
    End If
   
    Next shp
    Next sld
End Sub