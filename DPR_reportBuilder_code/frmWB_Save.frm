VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWB_Save 
   Caption         =   "DPR Report Builder"
   ClientHeight    =   3108
   ClientLeft      =   180
   ClientTop       =   756
   ClientWidth     =   7284
   OleObjectBlob   =   "frmWB_Save.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWB_Save"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim bUnload As Boolean
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdBox_Click()
    sDir = "\\tsclient\C\Users\" & sUser & "\Box\"
    lblSave.visible = True
    lblSave.Caption = "Saving file to your local Box folder. Please wait......"
    Call Save_WB
    lblSave.visible = False
    MsgBox "DPR Report Builder Saved.", vbOKOnly, "File Saved"
    bUnload = True
    Unload Me
End Sub

Private Sub cmdOneDrive_Click()
    sDir = "\\tsclient\C\Users\" & sUser & "\OneDrive - DPR Construction\Documents\"
    lblSave.visible = True
    lblSave.Caption = "Saving file to your local One Drive. Please wait......"
    Call Save_WB
    lblSave.visible = False
    If Range("rngIsTemp").Value = False Then Unload Me
    MsgBox "DPR Report Builder Saved.", vbOKOnly, "File Saved"
    bUnload = True
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 115
    Me.Left = Application.Left + 425
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Range("rngCP1").Value = False Then
        If bUnload = False Then
            Application.Quit
            ThisWorkbook.Close False
        End If
    End If
End Sub



