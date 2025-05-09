# Workbook PasswordEncryptionProvider Property

## Business Description
Returns a String specifying the name of the algorithm encryption provider that Microsoft Excel uses when encrypting passwords for the specified workbook. Read-only.

## Behavior
Returns aStringspecifying the name of the algorithm encryption provider that Microsoft Excel uses when encrypting passwords for the specified workbook. Read-only.

## Example Usage
```vba
Sub SetPasswordOptions() 
 
 With ActiveWorkbook 
 If .PasswordEncryptionProvider<> "Microsoft RSA SChannel Cryptographic Provider" Then 
 .SetPasswordEncryptionOptions _ 
 PasswordEncryptionProvider:="Microsoft RSA SChannel Cryptographic Provider", _ 
 PasswordEncryptionAlgorithm:="RC4", _ 
 PasswordEncryptionKeyLength:=56, _ 
 PasswordEncryptionFileProperties:=True 
 End If 
 End With 
 
End Sub
```