VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Progress Bar"
   ClientHeight    =   924
   ClientLeft      =   72
   ClientTop       =   2472
   ClientWidth     =   8208
   OleObjectBlob   =   "ProgressBar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private cFormTitle As String
Private cExcelStatusBar As Boolean
Private cTotalActions As Long
Private cActionNumber As Long
Private cStatusMessage As String
Private cBarWidth As Double
Private cPercentComplete As String

Private cFormShowStatus As Boolean
Private cTotalActionsSet As Boolean

Private cStartColourSet As Boolean
Private cEndColourSet As Boolean
Private cChangeColours As Boolean
Private cStartColour As XlRgbColor
Private cEndColour As XlRgbColor
Private cStartRed As Long, cEndRed As Long
Private cStartGreen As Long, cEndGreen As Long
Private cStartBlue As Long, cEndBlue As Long

Private Sub UserForm_Initialize()
Me.StartUpPosition = 0
Me.Top = Application.Top + 315
Me.Left = Application.Left + 475
cActionNumber = 0
cTotalActions = 4
cStatusMessage = "Ready"
cFormTitle = "Progress Bar"
cExcelStatusBar = False
cFormShowStatus = False
cTotalActionsSet = False
cPercentComplete = "0%"
cStartColourSet = False
cEndColourSet = False
cStartColour = rgbDodgerBlue
cChangeColours = False

Me.Title = cFormTitle
Me.StatusMessageBox.Caption = " " & cStatusMessage
Me.PercentIndicator.Caption = cPercentComplete
End Sub

Private Sub UserForm_Terminate()
Application.StatusBar = False
If Not Me.PercentIndicator = "100%" Then
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    End
End If
End Sub

Public Property Let Title(Value As String)

If cFormShowStatus Then
   ' Err.Raise 1, cFormTitle, "Set this Property before Running the Show Method."
        If Not Value = vbNullString Then
        cFormTitle = Value
        If Not Me Is Nothing Then
            DoEvents
            Me.Caption = cFormTitle
        End If
    End If
Else

    If Not Value = vbNullString Then
        cFormTitle = Value
        If Not Me Is Nothing Then
            DoEvents
            Me.Caption = cFormTitle
        End If
    End If
End If

End Property

Public Property Get Title() As String
Title = cFormTitle
End Property

Public Property Let ExcelStatusBar(Value As Boolean)

If cFormShowStatus Then
    Err.Raise 1, cFormTitle, "Set this Property before Running the Show Method."
Else
    cExcelStatusBar = Value
    If Value Then
        Application.DisplayStatusBar = True
    End If
End If
End Property

Public Property Get ExcelStatusBar() As Boolean
ExcelStatusBar = cExcelStatusBar
End Property

Public Property Let TotalActions(Value As Long)
If cFormShowStatus Then
    'Err.Raise 1, cFormTitle, "Set this Property before Running the Show Method."
    'If cTotalActionsSet Then
        'Err.Raise 4, cFormTitle, "TotalActions cannot be changed after it has been set."
    'Else
        cTotalActions = Value
        cTotalActionsSet = True
    'End If
Else
    If cTotalActionsSet Then
        Err.Raise 4, cFormTitle, "TotalActions cannot be changed after it has been set."
    Else
        cTotalActions = Value
        cTotalActionsSet = True
    End If
End If
   
End Property

Public Property Get TotalActions() As Long
TotalActions = cTotalActions
End Property

Public Property Let StartColour(Value As XlRgbColor)

If cFormShowStatus Then
    Err.Raise 1, cFormTitle, "Set this Property before Running the Show Method."
Else
    cStartColourSet = True
    cStartColour = Value
    cStartRed = GetPrimaryColour(cStartColour, "R")
    cStartGreen = GetPrimaryColour(cStartColour, "G")
    cStartBlue = GetPrimaryColour(cStartColour, "B")
End If

End Property

Public Property Let EndColour(Value As XlRgbColor)

If cFormShowStatus Then
    Err.Raise 1, cFormTitle, "Set this Property before Running the Show Method."
Else
    If Not cStartColourSet Then
        Err.Raise 8, cFormTitle, "Set StartColour First."
    Else
        cEndColourSet = True
        cEndColour = Value
        cEndRed = GetPrimaryColour(cEndColour, "R")
        cEndGreen = GetPrimaryColour(cEndColour, "G")
        cEndBlue = GetPrimaryColour(cEndColour, "B")
        cChangeColours = Not CBool(cStartColour = cEndColour)
    End If
End If

End Property

Public Property Let ActionNumber(Value As Long)

cActionNumber = Value

Call UpdateTheBar

End Property

Public Property Get ActionNumber() As Long
ActionNumber = cActionNumber
End Property

Public Property Let StatusMessage(Value As String)

cStatusMessage = Value
Call UpdateTheBar

End Property

Public Property Get StatusMessage() As String
StatusMessage = cStatusMessage
End Property

Public Sub ShowBar()
'If cFormShowStatus Then
'    Err.Raise 7, cFormTitle, "Progress Bar has already been Loaded."
'Else
    DoEvents
    cBarWidth = Me.ProgressBar.Width
    Me.ProgressBar.Width = 0
    Me.Caption = cFormTitle
    cFormShowStatus = True
    Me.ProgressBar.BackColor = cStartColour
    Me.ProgressBox.BorderColor = cStartColour
    Me.Show
    Me.Repaint
'End If
End Sub

Public Sub NextAction(Optional ByVal ProgressStatusMessage As String, _
    Optional ByVal ShowActionCount As Boolean = True)

cActionNumber = cActionNumber + 1
If ShowActionCount Then
    cStatusMessage = "Action " & cActionNumber & " of " _
        & cTotalActions & " | " & ProgressStatusMessage
Else
    cStatusMessage = ProgressStatusMessage
End If

Call UpdateTheBar
    
End Sub

Public Sub Complete(Optional ByVal WaitForSeconds As Long = 0, _
    Optional ByVal Prompt As String = "Complete")
    
Dim Counter As Long
If cFormShowStatus Then
    If cActionNumber < cTotalActions Then
        Err.Raise 6, cFormTitle, _
            "Run the Complete Method only after all the actions have been completed."
    Else
        If cExcelStatusBar Then
            Application.StatusBar = False
        End If

        If WaitForSeconds > 0 Then
            For Counter = WaitForSeconds To 1 Step -1
                DoEvents
                Me.StatusMessageBox.Caption = " " & Prompt & _
                    " | This Window will close in " & Counter & " " & _
                    IIf(Counter = 1, "second.", "seconds.")
                Application.Wait (Now() + TimeValue("00:00:01"))
            Next Counter
            Call Terminate
        Else
            DoEvents
            Me.StatusMessageBox.Caption = " " & Prompt
        End If
            
    End If
Else
    Err.Raise 5, cFormTitle, "Run the Show Method First"
End If
End Sub

Public Sub Terminate()

If cFormShowStatus Then
    If cFormShowStatus Then
        Me.Hide
        cFormShowStatus = False
        cTotalActionsSet = False
        cActionNumber = 0
        cTotalActions = 0
    End If
    If cExcelStatusBar Then
        Application.StatusBar = False
    End If
Else
    Err.Raise 5, cFormTitle, "Run the Show Method First"
End If
End Sub

Private Sub UpdateTheBar()

If cTotalActionsSet Then
    If cActionNumber > cTotalActions Then
        Err.Raise 3, cFormTitle, _
            "Current Action number is greater than Total Actions."
    Else
        If cFormShowStatus Then
            Call UpdateProgress
        Else
            Err.Raise 5, cFormTitle, "Run the Show Method First"
        End If
    End If
Else
    Err.Raise 2, cFormTitle, "Set TotalActions Property First."
End If

End Sub

Private Sub UpdateProgress()
Dim FractionComplete As Double
Dim ProgressPercent As String
Dim BarWidth As Double
Dim BarColour As XlRgbColor

FractionComplete = cActionNumber / cTotalActions
BarWidth = cBarWidth * FractionComplete
cPercentComplete = Format(FractionComplete * 100, "0") & "%"
DoEvents
Me.ProgressBar.Width = BarWidth
Me.PercentIndicator.Caption = cPercentComplete

If cChangeColours Then
    BarColour = RGB( _
        cStartRed + (cEndRed - cStartRed) * FractionComplete, _
        cStartGreen + (cEndGreen - cStartGreen) * FractionComplete, _
        cStartBlue + (cEndBlue - cStartBlue) * FractionComplete)
    Me.ProgressBar.BackColor = BarColour
End If

Me.StatusMessageBox.Caption = " " & cStatusMessage
If cExcelStatusBar Then
    Application.StatusBar = ProgressText(cActionNumber, cTotalActions) & _
        " | " & cPercentComplete & " | " & cStatusMessage
End If

Me.Repaint
End Sub

Private Function GetPrimaryColour(ByVal WhichColour As XlRgbColor, _
    ByVal RedBlueGreen As String) As Long
Dim HexString As String
HexString = CStr(Hex(WhichColour))
HexString = String(8 - Len(HexString), "0") & HexString

Select Case StrConv(RedBlueGreen, vbUpperCase)
    Case "R"
        HexString = "&H" & Mid(HexString, 7, 2)
    Case "G"
        HexString = "&H" & Mid(HexString, 5, 2)
    Case "B"
        HexString = "&H" & Mid(HexString, 3, 2)
    Case Else
        HexString = "-100"
End Select

GetPrimaryColour = CLng(HexString)
    
End Function

Function ProgressText(ByVal ActionNumber As Long, _
     ByVal TotalActions As Long, _
     Optional ByVal BarLength As Long = 15)
     
Dim BarComplete As Long
Dim BarInComplete As Long
Dim BarChar As String
Dim SpaceChar As String
Dim TempString As String

BarChar = ChrW(&H2589)
SpaceChar = ChrW(&H2000)

BarLength = Round(BarLength / 2, 0) * 2
BarComplete = Fix((ActionNumber * BarLength) / TotalActions)
BarInComplete = BarLength - BarComplete
ProgressText = String(BarComplete, BarChar) & String(BarInComplete, SpaceChar)
  
End Function
