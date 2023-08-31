Attribute VB_Name = "Main"
Dim passkey As String

Sub MainMenu()
    frmMain.Show
End Sub

Sub Decrypt_Text()
    On Error Resume Next
    Dim txt As String
    passkey = frmMain.tPasskey.Text
    If passkey = "" Then
        MsgBox "Please enter the passkey", vbInformation, "Encryptor"
        frmMain.tPasskey.SetFocus
        Exit Sub
    End If
    txt = "" & frmMain.tMsg.Text
    If txt = "" Then
        MsgBox "Please enter the text", vbInformation, "Encryptor"
        frmMain.tMsg.SetFocus
        Exit Sub
    End If
    txt = Decrypt(txt, passkey)
    If frmMain.tMsg.Text = txt Then
        MsgBox "Unable to decrypted text", vbExclamation, "Encryptor"
    Else
        frmMain.tMsg.Text = txt
    End If
End Sub

Sub Encrypt_Text()
    On Error Resume Next
    Dim txt As String
    txt = "" & frmMain.tMsg.Text
    passkey = frmMain.tPasskey.Text
    If passkey = "" Then
        MsgBox "Please enter the passkey", vbInformation, "Encryptor"
        frmMain.tPasskey.SetFocus
        Exit Sub
    End If
    If txt = "" Then
        MsgBox "Please enter the text", vbInformation, "Encryptor"
        frmMain.tMsg.SetFocus
        Exit Sub
    End If
    txt = Encrypt(txt, passkey)
    frmMain.tMsg.Text = txt
End Sub

Sub Decrypt_File()
    On Error Resume Next
    Dim txt As String, fn As String, newfile As String
    passkey = frmMain.tPasskey.Text
    If passkey = "" Then
        MsgBox "Please enter the passkey", vbInformation, "Encryptor"
        frmMain.tPasskey.SetFocus
        Exit Sub
    End If
    fn = Prompt4File()
    If fn = "" Then
        Exit Sub
    End If
    newfile = fn & ".dec"
    newfile = InputBox("Enter filename of the decrypted file:", "Encryptor", newfile)
    If newfile = "" Then
        Exit Sub
    End If
    If DecryptFile(fn, newfile, passkey) Then
        fn = Dir(newfile)
        MsgBox "File decrypted as " & fn, vbInformation, "Encryptor"
    Else
        MsgBox "Unable to decrypted file " & fn, vbExclamation, "Encryptor"
    End If
End Sub

Sub Encrypt_File()
    On Error Resume Next
    Dim txt As String, fn As String, newfile As String
    passkey = frmMain.tPasskey.Text
    If passkey = "" Then
        MsgBox "Please enter the passkey", vbInformation, "Encryptor"
        frmMain.tPasskey.SetFocus
        Exit Sub
    End If
    fn = Prompt4File()
    If fn = "" Then
        Exit Sub
    End If
    newfile = fn & ".enc"
    newfile = InputBox("Enter filename of the encrypted file:", "Encryptor", newfile)
    If newfile = "" Then
        Exit Sub
    End If
    Call EncryptFile(fn, newfile, passkey)
    fn = Dir(newfile)
    MsgBox "File encrypted as " & fn, vbInformation, "Encryptor"
End Sub

Sub Decrypt_Folder()
    On Error Resume Next
    Dim txt As String, fn As String, newfile As String, fld As String, n As Integer, arr_fn
    passkey = frmMain.tPasskey.Text
    If passkey = "" Then
        MsgBox "Please enter the passkey", vbInformation, "Encryptor"
        frmMain.tPasskey.SetFocus
        Exit Sub
    End If
    fn = Prompt4File()
    If fn = "" Then
        Exit Sub
    End If
    arr_fn = Split(fn, "\")
    n = UBound(arr_fn)
    fld = arr_fn(n)
    arr_fn = Split(fld, ".")
    newfile = Desktop() & arr_fn(0) & ".zip"
    FileDelete newfile
    fld = Desktop() & arr_fn(0)
    Call DecryptFile(fn, newfile, passkey)
    If DecryptFile(fn, newfile, passkey) Then
        Call unzip2folder(newfile, fld)
        FileDelete newfile
        Application.ScreenRefresh
        MsgBox "Folder Decrypted as " & fld, vbInformation, "Encryptor"
    Else
        MsgBox "Unable to decrypted file " & fn, vbExclamation, "Encryptor"
    End If
End Sub

Sub Encrypt_Folder()
    On Error Resume Next
    Dim txt As String, fn As String, newfile As String, fld As String, n As Integer, arr_fn
    passkey = frmMain.tPasskey.Text
    If passkey = "" Then
        MsgBox "Please enter the passkey", vbInformation, "Encryptor"
        frmMain.tPasskey.SetFocus
        Exit Sub
    End If
    fld = Prompt4Folder
    If fld = "" Then
        Exit Sub
    End If
    arr_fn = Split(fld, "\")
    n = UBound(arr_fn)
    fn = Desktop() & arr_fn(n) & ".zip"
    newfile = Desktop() & arr_fn(n) & ".enc"
    newfile = InputBox("Enter filename of the encrypted file:", "Encryptor", newfile)
    If newfile = "" Then
        Exit Sub
    End If
    FileDelete fn
    Call CreateZipFile(fld, fn)
    Call EncryptFile(fn, newfile, passkey)
    FileDelete fn
    Application.ScreenRefresh
    fn = Dir(newfile)
    MsgBox "Folder encrypted as " & fn, vbInformation, "Encryptor"
End Sub

Sub Encrypted_File_Info()
    On Error Resume Next
    Dim txt As String, fn As String, aes_code As String
    fn = Prompt4File()
    If fn = "" Then
        Exit Sub
    End If
    aes_code = InputBox("Enter secret code to reverse back the file encryption key :", "Encryptor", "")
    If aes_code <> aes_passkey() Then
        Exit Sub
    End If
    txt = DecryptHeader(fn)
    If txt = "" Then
        MsgBox "Not a valid encrypted file " & fn, vbExclamation, "Encryptor"
        Exit Sub
    End If
    MsgBox "This file is encrypted with key : " & txt, vbInformation, "Encryptor"
End Sub

Private Sub CleanupFolder(myfolder)
    On Error Resume Next
    Dim fn, folder_name
    Call CreateFolder(myfolder)
    folder_name = myfolder
    If Right(folder_name, 1) <> "\" Then
        folder_name = folder_name & "\"
    End If
    If Dir(folder_name, vbDirectory) = "" Then
        Exit Sub
    End If
    fn = Dir(folder_name & "*.*")
    Do While fn <> ""
        DoEvents
        Kill folder_name & fn
        fn = Dir()
    Loop
End Sub

Sub ClearMsg()
    On Error Resume Next
    frmMain.tMsg.Text = ""
End Sub

Private Sub CreateFolder(myfolder)
    On Error Resume Next
    If Dir(myfolder, vbDirectory) = "" Then
        MkDir myfolder
    End If
End Sub

Private Function CreateZipFile(folderToZipPath, zippedFileFullName)
    On Error GoTo zip_error
    Dim ShellApp As Object, myfolder As Variant, zipname As Variant
    CreateZipFile = False
    myfolder = folderToZipPath
    zipname = zippedFileFullName
    Open zipname For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
    Set ShellApp = CreateObject("Shell.Application")
    ShellApp.Namespace(zipname).CopyHere ShellApp.Namespace(myfolder).items
    On Error Resume Next
    Do Until ShellApp.Namespace(zipname).items.Count = ShellApp.Namespace(myfolder).items.Count
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    CreateZipFile = True
    Exit Function
zip_error:
    CreateZipFile = False
End Function

Private Function Desktop()
    On Error Resume Next
    Desktop = Replace(Environ$("appdata"), "AppData\Roaming", "Desktop\")
End Function

Private Sub FileDelete(fn)
    On Error Resume Next
    Kill fn
End Sub

Private Sub unzip2folder(myzip, myfolder, Optional ByVal EmptyFolder As Boolean = True)
    On Error Resume Next
    Dim oApp As Object, folder_name As Variant, zipname As Variant
    zipname = myzip
    folder_name = myfolder
    If EmptyFolder Then ' make sure folder is empty
        Call CleanupFolder(folder_name)
    End If
    Set oApp = CreateObject("Shell.Application")
    oApp.Namespace(folder_name).CopyHere oApp.Namespace(zipname).items
    Set oApp = Nothing
End Sub

Private Function Prompt4File()
    Dim fd As Office.FileDialog
    Dim strFile As String
    strFile = ""
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.Clear
        .Title = "Choose a file"
        .AllowMultiSelect = False
        .InitialFileName = Desktop()
        If .Show = True Then
            strFile = .SelectedItems(1)
        End If
    End With
    Prompt4File = strFile
End Function

Private Function Prompt4Folder()
    Dim fd As Office.FileDialog
    Dim strFile As String
    strFile = ""
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Filters.Clear
        .Title = "Choose a folder"
        .AllowMultiSelect = False
        .InitialFileName = Desktop()
        If .Show = True Then
            strFile = .SelectedItems(1)
        End If
    End With
    Prompt4Folder = strFile
End Function

