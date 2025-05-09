# Workbook PasswordEncryptionKeyLength Property

## Business Description
Returns a Long indicating the key length of the algorithm Microsoft Excel uses when encrypting passwords for the specified workbook. Read-only.

## Behavior
Returns aLongindicating the key length of the algorithm Microsoft Excel uses when encrypting passwords for the specified workbook. Read-only.

## Example Usage
```vba
Sub SetPasswordOptions() 
 
 With ActiveWorkbook 
 If .PasswordEncryptionKeyLength< 56 Then 
 .SetPasswordEncryptionOptions _ 
 PasswordEncryptionProvider:="Microsoft RSA SChannel Cryptographic Provider", _ 
 PasswordEncryptionAlgorithm:="RC4", _ 
 PasswordEncryptionKeyLength:=56, _ 
 PasswordEncryptionFileProperties:=True 
 End If 
 End With 
 
End Sub
```