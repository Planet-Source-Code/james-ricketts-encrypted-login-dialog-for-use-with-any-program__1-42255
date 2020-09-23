Attribute VB_Name = "mod_encryption"
Public filename As String
Public newpasswordlist As Boolean
Public authorisationneeded As Boolean
Public authorisationpass As String
Public switchtomainform As Boolean
Public encryptedvar, decryptedvar As String
Public Data(20) As String
Public switchtoserverform As Boolean
Public Index As Integer
Public loggedin As String
Public users(100) As String
Public username As String

Dim store(20) As String
Dim number(20) As Integer
Dim buffer(20) As String * 1
Dim encrypted(20), decrypted(20) As Integer
Dim encbuffer(20), decbuffer(20) As String * 1

Public Sub SetFormTopmost(TheForm As Form)

SetWindowPos TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
SWP_NOZORDER + SWP_NOMOVE + SWP_NOSIZE

End Sub

Public Function Encrypt(Data As String, length As Integer)

encryptedvar = ""

For i = 1 To length
    store(i) = Left(Data, i)
    buffer(i) = (Right(store(i), 1))
    number(i) = Asc(buffer(i))
    encrypted(i) = ((number(i) * 2) - 45)
    encbuffer(i) = Chr(encrypted(i))
    encryptedvar = (encryptedvar & encbuffer(i))
Next

Data = encryptedvar

End Function


Public Function Decrypt(Data As String, length As Integer)

decryptedvar = ""

For i = 1 To length
    store(i) = Left(Data, i)
    buffer(i) = (Right(store(i), 1))
    number(i) = Asc(buffer(i))
    decrypted(i) = ((number(i) + 45) / 2)
    If number(i) = "0" Then
    decrypted(i) = "0"
    End If
    decbuffer(i) = Chr(decrypted(i))
    decryptedvar = (decryptedvar & decbuffer(i))
Next

Data = decryptedvar

End Function
