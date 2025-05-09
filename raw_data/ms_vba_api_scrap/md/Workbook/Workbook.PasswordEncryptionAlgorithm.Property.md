# Workbook PasswordEncryptionAlgorithm Property

## Business Description
Returns a String indicating the algorithm Microsoft Excel uses to encrypt passwords for the specified workbook. Read-only.

## Behavior
Returns aStringindicating the algorithm Microsoft Excel uses to encrypt passwords for the specified workbook. Read-only.

## Example Usage
```vba
Sub SetPasswordOptions() 
 
 ActiveWorkbook.SetPasswordEncryptionOptions _ 
 PasswordEncryptionProvider:="Microsoft RSA SChannel Cryptographic Provider", _PasswordEncryptionAlgorithm:="RC4", _ 
 PasswordEncryptionKeyLength:=56, _ 
 PasswordEncryptionFileProperties:=True 
 
End Sub
```