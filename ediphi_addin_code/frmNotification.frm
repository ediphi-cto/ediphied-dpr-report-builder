VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNotification 
   Caption         =   "ediphi"
   ClientHeight    =   2472
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6816
   OleObjectBlob   =   "frmNotification.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub setMsg(msg1str As String, msg2str As String)

    msg1.Caption = msg1str
    msg2.Caption = msg2str

End Sub

Sub closeMe()
    Unload Me
End Sub

