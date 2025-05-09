Sub DisplayXMLMap() 
 Dim objCustomer As XmlMap 
 
 Set objCustomer = ActiveWorkbook.XmlMaps.Add( _ 
 "Customers.xsd", "Root") 
 
 objCustomer.Name = "Customers" 
 
 Application.DisplayXMLSourcePane 
 objCustomer 
End Sub