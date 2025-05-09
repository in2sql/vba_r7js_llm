# QueryTable EditWebPage Property

## Business Description
Returns or sets the web page Uniform Resource Locator (URL) for a web query. Read/write Variant.

## Behavior
Returns or sets the web page Uniform Resource Locator (URL) for a web query. Read/writeVariant.

## Example Usage
```vba
Sub ReturnURL() 
 
 ' Set the EditWebPage property to a source. 
 Range("A1").QueryTable.EditWebPage= "C:\MyHomepage.htm" 
 
 ' Display the source to the user. 
 MsgBox Range("A1").QueryTable.EditWebPageEnd Sub
```