Attribute VB_Name = "ExcelExportDataToXML"
' ---------------------------------------------------------------------------------------------------------------------------------
' -- ExcelExportDataToXML.bas                                                                                                    --
' --                                                                                                                             --
' ---------------------------------------------------------------------------------------------------------------------------------
' --                                                      Revision History                                                       --
' ---------------------------------------------------------------------------------------------------------------------------------
' -- VV.vv.DOYyy.bb (dd MMM yy) - Initial Creation/Development Update/Maintenance Update                                         --
' --                                                                                                                             --
' --  1.00.23523.xx (23 Aug 23) - Initial Creation {J. Laccone}                                                                  --
' --                                 Added header, added reference data                                                          --
' --                                                                                                                             --
' ---------------------------------------------------------------------------------------------------------------------------------
' --                                                          Reference                                                          --
' ---------------------------------------------------------------------------------------------------------------------------------
' --                                                                                                                             --
' -- Map An Excel Worksheet To XML In VBA                                                                                        --
' -- ------------------------------------                                                                                        --
' -- https://stackoverflow.com/questions/45046249/automatically-map-an-excel-xmlmap-to-a-worksheet-in-vba-without-knowing-the-sche
' --                                                                                                                             --
' ---------------------------------------------------------------------------------------------------------------------------------
Sub ExportDataToXML()

   ' -----------------------------------------
   ' -- Variable Declaration/Initialization --
   ' -----------------------------------------
   Dim outputFile As String
   Dim objMapToExport As XmlMap
   Dim strXPath As String
   Dim lstCol1 As ListObject

   ' Specify the destination file for the exported data
   outputFile = "xml-export-example-document.xml"


   ' ---------------------
   ' -- XML Map Loading --
   ' ---------------------

   ' Add the XML map to the maps collection
   ' ref: https://learn.microsoft.com/en-us/office/vba/api/excel.xmlmaps.add
   Set objMapToExport = ActiveWorkbook.XmlMaps.Add("xml-export-example-document-map.xml", "xml-export-example-document")

   ' Name the map
   objMapToExport.Name = "Xeed_Map"


   ' ---------------------------------------
   ' -- Excel Table (ListObject) Creation --
   ' ---------------------------------------

   ' Create an Excel table to contain the source data to be exported
   ' ref: https://learn.microsoft.com/en-us/office/vba/api/excel.listobjects.add
   Set lstCol1 = Sheet1.ListObjects.Add(xlSrcRange, Range("A1:C4"), , xlYes)

   ' Name the Excel table
   lstCol1.Name = "dataTable1"


   ' -----------------------------------
   ' -- Create Column To Tag Bindings --
   ' -----------------------------------
   ' ref: https://learn.microsoft.com/en-us/office/vba/api/excel.listcolumn.xpath

   ' -- Column: column1 --

   ' XPath expression for XML tag representing column
   strXPath = "/xml-export-example-document/record/column1"

   ' Bind data to tag via XPath
   lstCol1.ListColumns(1).XPath.SetValue objMapToExport, strXPath


   ' -- Column: column2 --

   ' XPath expression for XML tag representing column
   strXPath = "/xml-export-example-document/record/column2"

   ' Bind data to tag via XPath
   lstCol1.ListColumns(2).XPath.SetValue objMapToExport, strXPath


   ' -- Column: column2 --

   ' XPath expression for XML tag representing column
   strXPath = "/xml-export-example-document/record/column3"

   ' Bind data to tag via XPath
   lstCol1.ListColumns(3).XPath.SetValue objMapToExport, strXPath


   ' ----------------------------------
   ' -- Export the Excel data to XML --
   ' ----------------------------------

   ' Verify that the data is mapped and can be correctly exported
   If objMapToExport.IsExportable Then

      ' Save the mapped data to the specified XML file
      ' ref: https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.saveasxmldata
      ActiveWorkbook.SaveAsXMLData outputFile, objMapToExport

   Else

      ' Alert the user that the data is not correctly mapped and can't be exported
      MsgBox "Can't use " & objMapToExport.Name & " to export the contents of the worksheet to XML data"

   End If

End Sub
