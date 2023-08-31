Attribute VB_Name = "AES_lib"
Option Explicit
Const MACKey = "............................................"
Const passkey = "............"

Dim AES As Object
Dim utf8 As Object
Dim mem As Object
Dim Mac As Object

Public Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Source As Any, ByVal Bytes As Long)
                          
Private Function B64Encode(Bytes)
    Dim BlockSize, b64Block, offset, Length, result
    Dim b64Enc As Object
    Set b64Enc = CreateObject("System.Security.Cryptography.ToBase64Transform")
    BlockSize = b64Enc.InputBlockSize
    For offset = 0 To LenB(Bytes) - 1 Step BlockSize
        Length = LenB(Bytes) - offset
        If BlockSize < Length Then
            Length = BlockSize
        End If
        b64Block = b64Enc.TransformFinalBlock((Bytes), offset, Length)
        result = result & utf8.GetString((b64Block))
    Next
    B64Encode = result
End Function

Private Function B64Decode(b64Str)
    Dim b64Dec As Object
    Dim Bytes
    Set b64Dec = CreateObject("System.Security.Cryptography.FromBase64Transform")
    Bytes = utf8.GetBytes_4(b64Str)
    B64Decode = b64Dec.TransformFinalBlock((Bytes), 0, LenB(Bytes))
End Function

Private Function ConcatBytes(a, b)
    mem.SetLength (0)
    mem.Write (a), 0, LenB(a)
    mem.Write (b), 0, LenB(b)
    ConcatBytes = mem.ToArray()
End Function

Private Function EqualBytes(a, b)
    Dim diff, i
    EqualBytes = False
    If LenB(a) <> LenB(b) Then Exit Function
    diff = 0
    For i = 1 To LenB(a)
        diff = diff Or (AscB(MidB(a, i, 1)) Xor AscB(MidB(b, i, 1)))
    Next
    EqualBytes = Not diff
End Function

Private Function ComputeMAC(msgBytes, keyBytes)
    Mac.Key = keyBytes
    ComputeMAC = Mac.ComputeHash_2((msgBytes))
End Function

Function Encrypt(plaintext, passkey)
    Dim aesKeyBytes, macKeyBytes, aesEnc, plainBytes, cipherBytes, macBytes, aesKey
    Set AES = CreateObject("System.Security.Cryptography.RijndaelManaged")
    Set utf8 = CreateObject("System.Text.UTF8Encoding")
    Set mem = CreateObject("System.IO.MemoryStream")
    Set Mac = CreateObject("System.Security.Cryptography.HMACSHA256")

    aesKey = B64Encode(utf8.GetBytes_4(passkey))
    Call AES.GenerateIV
    aesKeyBytes = B64Decode(aesKey)
    macKeyBytes = B64Decode(MACKey)
    Set aesEnc = AES.CreateEncryptor_2((aesKeyBytes), AES.IV)
   
    plainBytes = utf8.GetBytes_4(plaintext)
    cipherBytes = aesEnc.TransformFinalBlock((plainBytes), 0, LenB(plainBytes))
    macBytes = ComputeMAC(ConcatBytes(AES.IV, cipherBytes), macKeyBytes)
    Encrypt = B64Encode(macBytes) & ":" & B64Encode(AES.IV) & ":" & _
              B64Encode(cipherBytes)
End Function

Function Decrypt(macIVCiphertext, passkey)
    Dim aesKeyBytes, macKeyBytes, tokens, macBytes, ivBytes, cipherBytes, macActual, plainBytes, aesDec, aesKey
    Set AES = CreateObject("System.Security.Cryptography.RijndaelManaged")
    Set utf8 = CreateObject("System.Text.UTF8Encoding")
    Set mem = CreateObject("System.IO.MemoryStream")
    Set Mac = CreateObject("System.Security.Cryptography.HMACSHA256")
    
    aesKey = B64Encode(utf8.GetBytes_4(passkey))
    aesKeyBytes = B64Decode(aesKey)
    macKeyBytes = B64Decode(MACKey)
    tokens = Split(macIVCiphertext, ":")
    macBytes = B64Decode(tokens(0))
    ivBytes = B64Decode(tokens(1))
    cipherBytes = B64Decode(tokens(2))
    macActual = ComputeMAC(ConcatBytes(ivBytes, cipherBytes), macKeyBytes)
    If Not EqualBytes(macBytes, macActual) Then
        Err.Raise vbObjectError + 1000, "Decrypt()", "Bad MAC"
    End If
    Set aesDec = AES.CreateDecryptor_2((aesKeyBytes), (ivBytes))
    plainBytes = aesDec.TransformFinalBlock((cipherBytes), 0, LenB(cipherBytes))
    Decrypt = utf8.GetString((plainBytes))
End Function

Sub EncryptFile(filePath, outputfile, pw)
    On Error Resume Next
    Dim payload As String
    If pw = "" Then
        Exit Sub
    End If
    payload = Encrypt(pw, passkey)
    CopyBinaryFile filePath, outputfile, 0, payload
End Sub

Function DecryptFile(filePath, outputfile, pw)
    On Error Resume Next
    Dim payload As String, txt As String
    DecryptFile = False
    If pw = "" Then
        Exit Function
    End If
    txt = DecryptHeader(filePath)
    If pw = txt Then
        CopyBinaryFile filePath, outputfile, 94, ""
        DecryptFile = True
    End If
End Function

Function DecryptHeader(srcfile)
    On Error Resume Next
    Dim fileNumber As Integer
    Dim fileSize As Long, pos As Long
    Dim Bytes() As Byte
    Dim txt As String
    DecryptHeader = ""
    fileNumber = FreeFile
    Open srcfile For Binary Access Read As fileNumber
    fileSize = FileLen(srcfile)
    ReDim Bytes(1 To fileSize) As Byte
    Get fileNumber, , Bytes
    Close fileNumber
    txt = ""
    For pos = 1 To 94
        DoEvents
        txt = txt & Chr(Bytes(pos))
    Next
    DecryptHeader = Decrypt(txt, passkey)
End Function

Function aes_passkey()
    aes_passkey = passkey
End Function

Private Sub CopyBinaryFile(srcfile, destfile, Optional ByVal offset As Long = 0, Optional ByVal payload As String = "")
    On Error Resume Next
    Dim fileNumber As Integer
    Dim fileSize As Long, newsize As Long
    Dim Bytes() As Byte
    Dim lessbytes() As Byte
    Dim txt As String
    
    ' Open the binary file for binary access
    fileNumber = FreeFile
    Open srcfile For Binary Access Read As fileNumber

    ' Determine the size of the file
    fileSize = FileLen(srcfile)
    newsize = fileSize - offset

    ' Read the entire file into a byte array
    ReDim Bytes(1 To fileSize) As Byte
    ReDim lessbytes(1 To newsize) As Byte
    Get fileNumber, , Bytes
    Close fileNumber

    ' Process the byte array as needed
    If offset = 0 Then
        lessbytes = Bytes
    Else
        CopyMemory lessbytes(1), Bytes(offset + 1), newsize
    End If
    
    If Dir(destfile) <> "" Then
        Kill destfile
    End If
    
    ' Open the binary file for binary access
    fileNumber = FreeFile
    Open destfile For Binary Access Write As fileNumber
    DoEvents
    If payload <> "" Then
        Put fileNumber, , payload
    End If
    Put fileNumber, , lessbytes
    Close fileNumber
    Erase Bytes
    Erase lessbytes
End Sub


