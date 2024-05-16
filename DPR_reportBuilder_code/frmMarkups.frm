VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMarkups 
   Caption         =   "DPR Report Builder"
   ClientHeight    =   5856
   ClientLeft      =   72
   ClientTop       =   468
   ClientWidth     =   8988
   OleObjectBlob   =   "frmMarkups.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMarkups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim F As Integer
Private listMainItem As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim C
    Set ows = Sheet0
    Set lObj = ows.ListObjects("tblTotals")
    For i = 1 To ListBox1.ListCount
        With lObj
            Set C = .ListColumns(1).DataBodyRange.Find(ListBox1.List(i - 1, 0), LookIn:=xlValues)
            If Not C Is Nothing Then
                If ListBox1.List(i - 1, 6) <> "" Then
                    .DataBodyRange(i, 7).Value = ListBox1.List(i - 1, 6)
                Else
                    .DataBodyRange(i, 7).Value = "Lower"
                End If
            End If
        End With
    Next i
    If lObj.ListRows.count > 0 Then
        Call clearAllAddons
        Call ReApplyAddons
        Call ExecSummary
    Else
        MsgBox "There are no Markups to manage.", vbOKOnly, "No Markups Found"
    End If
    Unload Me
End Sub

Private Sub ListBox1_Change()
    x = ListBox1.ListIndex
    If x = -1 Then Exit Sub
    If ListBox1.Selected(x) = True Then
        ListBox1.List(x, 6) = "Upper"
    Else
        ListBox1.List(x, 6) = "Lower"
    End If
End Sub

Private Sub UserForm_Activate()
    Set ows = Sheet0
    Set lObj = ows.ListObjects("tblTotals")
    If lObj.ListRows.count = 0 Then Unload Me
    i = 0
    With ListBox1
        For i = 1 To lObj.ListRows.count
            .AddItem
            .List(i - 1, 0) = lObj.DataBodyRange(i, 1)
            .List(i - 1, 1) = lObj.DataBodyRange(i, 2)
            .List(i - 1, 2) = lObj.DataBodyRange(i, 3)
            .List(i - 1, 3) = lObj.DataBodyRange(i, 4)
            .List(i - 1, 4) = FormatPercent(lObj.DataBodyRange(i, 5) / 100, 2)
            .List(i - 1, 5) = FormatCurrency(lObj.DataBodyRange(i, 6), 0)
            .List(i - 1, 6) = lObj.DataBodyRange(i, 7)
            .List(i - 1, 7) = lObj.DataBodyRange(i, 8)
            If lObj.DataBodyRange(i, 7) = "Upper" Then
                .Selected(i - 1) = True
            End If
        Next
    End With
End Sub
Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 115
    Me.Left = Application.Left + 25
End Sub

