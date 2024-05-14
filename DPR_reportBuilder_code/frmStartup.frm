VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStartup 
   Caption         =   "DPR Construction"
   ClientHeight    =   3336
   ClientLeft      =   168
   ClientTop       =   636
   ClientWidth     =   6864
   OleObjectBlob   =   "frmStartup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim bUnload As Boolean
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Stop
Dim strFileToOpen As String
Dim lngCount As Long
    If optNew.Value = True Then
        bUnload = True
        Unload Me
        Call NewReport
    ElseIf optExisting.Value = True Then
        sUser = CStr(Environ("USERNAME"))
        Workbooks.Open FileOpen("\\tsclient\C\Users\" & sUser & "\OneDrive - DPR Construction\Documents\", "DPR Reporter", "*.xlsm")
                                 
        ThisWorkbook.Close False
        Unload Me
    End If
End Sub


Private Sub UserForm_Activate()
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 115
    Me.Left = Application.Left + 25
    lblVersion.Caption = Range("rngVersion").Value
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If bUnload = False Then
        ThisWorkbook.Close False
    End If
End Sub
