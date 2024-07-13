VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHeadings 
   Caption         =   "DPR Report Builder"
   ClientHeight    =   3588
   ClientLeft      =   48
   ClientTop       =   360
   ClientWidth     =   10992
   OleObjectBlob   =   "frmHeadings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHeadings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim sHdr As String

Private Sub CheckBox1_Click()
    txtFontBold
End Sub

Private Sub CheckBox2_Click()
    txtFontBold
End Sub

Private Sub CheckBox3_Click()
    txtFontBold
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Set ows = ActiveSheet
    ows.Cells(1, 2).Value = txtMainHdr.Value
    Range("rngHeading1").Value = txtHeading1.Value
'    Range("rngHeading2").Value = txtHeading2.Value
    Range("rngHeading3").Value = txtHeading3.Value
    Range("rngSubHeading1").Value = txtSubHeading1.Value
    Range("rngSubHeading2").Value = txtSubHeading2.Value
    Range("rngSubHeading3").Value = txtSubHeading3.Value
    Range("rngSubHeading4").Value = txtSubHeading4.Value
    Range("rngSubHeading5").Value = txtSubHeading5.Value
    Unload Me
End Sub

Private Sub SpinButton1_Change()
    numFont.Value = SpinButton1.Value
End Sub

Private Sub SpinButton1_SpinDown()
    Call txtFontSize
End Sub

Private Sub SpinButton1_SpinUp()
    Call txtFontSize
End Sub

Private Sub txtMainHdr_Change()
    txtMainHdr.Value = StrConv(txtMainHdr.Value, vbUpperCase)
End Sub

Private Sub txtMainHdr_Enter()
    SpinButton1.Value = txtMainHdr.Font.Size
    CheckBox1.Value = txtMainHdr.Font.Bold
    CheckBox2.Value = txtMainHdr.Font.Italic
    sHdr = "txtMainHdr"
End Sub

Private Sub txtHeading1_Enter()
    SpinButton1.Value = txtHeading1.Font.Size
    CheckBox1.Value = txtHeading1.Font.Bold
    CheckBox2.Value = txtHeading1.Font.Italic
    sHdr = "txtHeading1"
End Sub

Private Sub txtHeading2_Enter()
'    SpinButton1.Value = txtHeading2.Font.Size
'    CheckBox1.Value = txtHeading2.Font.Bold
'    CheckBox2.Value = txtHeading2.Font.Italic
'    sHdr = "txtHeading2"
End Sub

Private Sub txtHeading3_Enter()
    SpinButton1.Value = txtHeading3.Font.Size
    CheckBox1.Value = txtHeading3.Font.Bold
    CheckBox2.Value = txtHeading3.Font.Italic
    sHdr = "txtHeading3"
End Sub

Private Sub txtSubHeading1_Enter()
    SpinButton1.Value = txtSubHeading1.Font.Size
    CheckBox1.Value = txtSubHeading1.Font.Bold
    CheckBox2.Value = txtSubHeading1.Font.Italic
    sHdr = "txtSubHeading1"
End Sub

Private Sub txtSubHeading2_Enter()
    SpinButton1.Value = txtSubHeading2.Font.Size
    CheckBox1.Value = txtSubHeading2.Font.Bold
    CheckBox2.Value = txtSubHeading2.Font.Italic
    sHdr = "txtSubHeading2"
End Sub

Private Sub txtSubHeading3_Enter()
    SpinButton1.Value = txtSubHeading3.Font.Size
    CheckBox1.Value = txtSubHeading3.Font.Bold
    CheckBox2.Value = txtSubHeading3.Font.Italic
    sHdr = "txtSubHeading3"
End Sub

Private Sub txtSubHeading4_Enter()
    SpinButton1.Value = txtSubHeading4.Font.Size
    CheckBox1.Value = txtSubHeading4.Font.Bold
    CheckBox2.Value = txtSubHeading4.Font.Italic
    sHdr = "txtSubHeading4"
End Sub

Private Sub txtSubHeading5_Enter()
    SpinButton1.Value = txtSubHeading5.Font.Size
    CheckBox1.Value = txtSubHeading5.Font.Bold
    CheckBox2.Value = txtSubHeading5.Font.Italic
    sHdr = "txtSubHeading5"
End Sub

Private Sub UserForm_Activate()
    Set ows = ActiveSheet
    txtMainHdr.Value = ows.Cells(1, 2).Value
    txtMainHdr.BackColor = ows.Cells(1, 2).Interior.Color
    txtMainHdr.ForeColor = ows.Cells(1, 2).Font.Color
    txtHeading1.Value = Range("rngHeading1").Value
'    txtHeading2.Value = Range("rngHeading2").Value
    txtHeading3.Value = Range("rngHeading3").Value
    txtSubHeading1.Value = Range("rngSubHeading1").Value
    txtSubHeading2.Value = Range("rngSubHeading2").Value
    txtSubHeading3.Value = Range("rngSubHeading3").Value
    txtSubHeading4.Value = Range("rngSubHeading4").Value
    txtSubHeading5.Value = Range("rngSubHeading5").Value
    
    txtMainHdr.Font.Size = ows.Cells(1, 2).Font.Size
    txtHeading1.Font.Size = ows.Shapes("txtHeading1").TextFrame.Characters.Font.Size
'    txtHeading2.Font.Size = ows.Shapes("txtHeading2").TextFrame.Characters.Font.Size
    txtHeading3.Font.Size = ows.Shapes("txtHeading3").TextFrame.Characters.Font.Size
    txtSubHeading1.Font.Size = ows.Shapes("txtSubHeading1").TextFrame.Characters.Font.Size
    txtSubHeading2.Font.Size = ows.Shapes("txtSubHeading2").TextFrame.Characters.Font.Size
    txtSubHeading3.Font.Size = ows.Shapes("txtSubHeading3").TextFrame.Characters.Font.Size
    txtSubHeading4.Font.Size = ows.Shapes("txtSubHeading4").TextFrame.Characters.Font.Size
    txtSubHeading5.Font.Size = ows.Shapes("txtSubHeading5").TextFrame.Characters.Font.Size
    
    txtMainHdr.Font.Bold = ows.Cells(1, 2).Font.Bold
    txtHeading1.Font.Bold = ows.Shapes("txtHeading1").TextFrame.Characters.Font.Bold
'    txtHeading2.Font.Bold = ows.Shapes("txtHeading2").TextFrame.Characters.Font.Bold
    txtHeading3.Font.Bold = ows.Shapes("txtHeading3").TextFrame.Characters.Font.Bold
    txtSubHeading1.Font.Bold = ows.Shapes("txtSubHeading1").TextFrame.Characters.Font.Bold
    txtSubHeading2.Font.Bold = ows.Shapes("txtSubHeading2").TextFrame.Characters.Font.Bold
    txtSubHeading3.Font.Bold = ows.Shapes("txtSubHeading3").TextFrame.Characters.Font.Bold
    txtSubHeading4.Font.Bold = ows.Shapes("txtSubHeading4").TextFrame.Characters.Font.Bold
    txtSubHeading5.Font.Bold = ows.Shapes("txtSubHeading5").TextFrame.Characters.Font.Bold
    
    txtMainHdr.Font.Italic = ows.Cells(1, 2).Font.Italic
    txtHeading1.Font.Italic = ows.Shapes("txtHeading1").TextFrame.Characters.Font.Italic
'    txtHeading2.Font.Italic = ows.Shapes("txtHeading2").TextFrame.Characters.Font.Italic
    txtHeading3.Font.Italic = ows.Shapes("txtHeading3").TextFrame.Characters.Font.Italic
    txtSubHeading1.Font.Italic = ows.Shapes("txtSubHeading1").TextFrame.Characters.Font.Italic
    txtSubHeading2.Font.Italic = ows.Shapes("txtSubHeading2").TextFrame.Characters.Font.Italic
    txtSubHeading3.Font.Italic = ows.Shapes("txtSubHeading3").TextFrame.Characters.Font.Italic
    txtSubHeading4.Font.Italic = ows.Shapes("txtSubHeading4").TextFrame.Characters.Font.Italic
    txtSubHeading5.Font.Italic = ows.Shapes("txtSubHeading5").TextFrame.Characters.Font.Italic
    
End Sub

Sub txtFontSize()
    On Error Resume Next
    For Each ows In ActiveWorkbook.Worksheets
        If sHdr = "txtMainHdr" Then
            ows.Cells(1, 2).Font.Size = numFont.Value
        Else
            With ows.Shapes(sHdr).TextFrame.Characters.Font
                .Size = numFont.Value
            End With
        End If
    Next ows
    Controls(sHdr).Font.Size = numFont.Value
    On Error GoTo 0
End Sub

Sub txtFontBold()
    On Error Resume Next
    For Each ows In ActiveWorkbook.Worksheets
        If sHdr = "txtMainHdr" Then
            ows.Cells(1, 2).Font.Bold = CheckBox1.Value
            ows.Cells(1, 2).Font.Italic = CheckBox2.Value
        Else
            With ows.Shapes(sHdr).TextFrame.Characters.Font
                .Bold = CheckBox1.Value
                .Italic = CheckBox2.Value
            End With
        End If
    Next ows
    Controls(sHdr).Font.Bold = CheckBox1.Value
    Controls(sHdr).Font.Italic = CheckBox2.Value
    sHdr = ""
    On Error GoTo 0
End Sub



Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 115
    Me.Left = Application.Left + 25
End Sub
