'algoritmo de criptografia simétrica AES (Advanced Encryption Standard):
'Nota: É importante que você armazene a senha usada para criptografar a string de forma segura, pois sem ela não será possível descriptografar a string.

'Criptografar uma string:

Function EncryptString(ByVal strText As String, ByVal strPass As String) As String
    Dim bytText() As Byte
    Dim bytPass() As Byte
    Dim bytSalt() As Byte = {1, 2, 3, 4, 5, 6, 7, 8}
    Dim intLength As Integer
    Dim intRemaining As Integer
    Dim objCrypto As Object
    Dim objKey As Object
    Dim objIV As Object

    'Convert the plaintext string to a byte array.
    bytText = StrConv(strText, vbFromUnicode)

    'Create the key and IV based on the password.
    intLength = Len(strPass)
    intRemaining = intLength Mod 8
    If intRemaining > 0 Then
        strPass = strPass & String(8 - intRemaining, Chr(0))
    End If
    bytPass = StrConv(strPass, vbFromUnicode)

    'Create the encryption objects.
    Set objCrypto = CreateObject("System.Security.Cryptography.RijndaelManaged")
    Set objKey = objCrypto.CreateEncryptor(bytPass, bytSalt)
    Set objIV = objCrypto.IV

    'Encrypt the plaintext.
    EncryptString = objKey.TransformFinalBlock(bytText, 0, UBound(bytText) + 1)

    'Concatenate the IV and the ciphertext and convert to a string.
    EncryptString = Convert.ToBase64String(objIV) & Convert.ToBase64String(EncryptString)
End Function

'Descriptografar uma string
Function DecryptString(ByVal strText As String, ByVal strPass As String) As String
    Dim bytText() As Byte
    Dim bytPass() As Byte
    Dim bytSalt() As Byte = {1, 2, 3, 4, 5, 6, 7, 8}
    Dim intLength As Integer
    Dim intRemaining As Integer
    Dim objCrypto As Object
    Dim objKey As Object
    Dim objIV As Object
    Dim intIVLength As Integer

    'Extract the IV from the ciphertext.
    intIVLength = 24
    ReDim bytText(intIVLength - 1)
    bytText = Convert.FromBase64String(Left(strText, intIVLength))
    strText = Right(strText, Len(strText) - intIVLength)

    'Convert the ciphertext string to a byte array.
    bytText = Convert.FromBase64String(strText)

    'Create the key and IV based on the password.
    intLength = Len(strPass)
    intRemaining = intLength Mod 8
    If intRemaining > 0 Then
        strPass = strPass & String(8 - intRemaining, Chr(0))
    End If
    bytPass = StrConv(strPass, vbFromUnicode)

    'Create the decryption objects.
    Set objCrypto = CreateObject("System.Security.Cryptography.RijndaelManaged")
    Set objKey = objCrypto.CreateDecryptor(bytPass, bytSalt)
    Set objIV = objCrypto.IV

    'Decrypt the ciphertext.
    bytText = objKey.TransformFinalBlock(bytText, 0, UBound(bytText) + 1)

    'Convert the plaintext byte array to a string.
    DecryptString = StrConv(bytText, vbUnicode)
End Function



