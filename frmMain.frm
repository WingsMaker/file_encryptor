VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Developed by Lim Kim Huat"
   ClientHeight    =   7180
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   6020
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdClear_Click()
    Call ClearMsg
End Sub

Private Sub cmdDecryptFile_Click()
    Call Decrypt_File
End Sub

Private Sub cmdDecryptFolder_Click()
    Call Decrypt_Folder
End Sub

Private Sub cmdDecryptText_Click()
    Call Decrypt_Text
End Sub

Private Sub cmdEncryptFile_Click()
    Call Encrypt_File
End Sub

Private Sub cmdEncryptFolder_Click()
    Call Encrypt_Folder
End Sub

Private Sub cmdEncryptText_Click()
    Call Encrypt_Text
End Sub

Private Sub cmdFileInfo_Click()
    Call Encrypted_File_Info
End Sub

Private Sub lblTitle_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call ClearMsg
End Sub

Private Sub cmdClose_Click()
    Sheets("Main").Visible = True
    ActiveWindow.WindowState = xlNormal
    ActiveWorkbook.Save
    Me.Hide
End Sub
