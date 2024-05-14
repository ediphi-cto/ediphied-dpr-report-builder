VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReportPrint 
   Caption         =   "DPR Report Builder"
   ClientHeight    =   7224
   ClientLeft      =   72
   ClientTop       =   612
   ClientWidth     =   8268
   OleObjectBlob   =   "frmReportPrint.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReportPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim bPreview As Boolean
Dim nFlag As Boolean
Dim bLoading As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Sub shtSelect()
nFlag = True
    With lstSheets
        For C = 0 To .ListCount - 1
            If .Selected(C) = True Then
                Sheets(.List(C, 0)).Select nFlag
                Set ows = Sheets(.List(C, 0))
                nFlag = False
            End If
        Next C
    End With
End Sub

Private Sub cmdOK_Click()
    X = 0
    For i = 0 To lstSheets.ListCount - 1
        If lstSheets.Selected(i) = True Then
            X = X + 1
        End If
    Next i
    If X = 0 Then
        MsgBox "No sheets were selected." & vbCrLf & "Please make a selection to continue.", vbCritical, "Invalid Selection"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Call shtSelect
    If bPreview = False Then SheetFormatting
    ActiveWindow.SelectedSheets.PrintOut
    Sheets(sSht).Select
    Application.ScreenUpdating = True
    Unload Me
    Exit Sub
End Sub

Private Sub cmdPdf_Click()
    Dim fileName As String
    Dim fileSavename As String
    Dim ans
    X = 0
    sDir = CurDir()
    For i = 0 To lstSheets.ListCount - 1
        If lstSheets.Selected(i) = True Then
            X = X + 1
        End If
    Next i
    If X = 0 Then
        MsgBox "No sheets were selected." & vbCrLf & "Please make a selection to continue.", vbCritical, "Invalid Selection"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Call shtSelect
    SheetFormatting
    sUser = Environ("UserName")
    fileName = Range("rngProjectName").Value
    sDir = DriveName(ThisWorkbook.Path)
    fileName = sDir & "\Users\" & sUser & "\Documents\" & fileName
    fileSavename = Application.GetSaveAsFilename(fileName, FileFilter:="pdf Files (*.pdf), *.pdf")
    If fileSavename <> "False" Then
        bPdf = True
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileSavename, _
             Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    End If
    Sheets(sSht).Select
    Application.ScreenUpdating = True
    Unload Me
    Exit Sub
End Sub

Private Sub cmdPreview_Click()
    X = 0
    For i = 0 To lstSheets.ListCount - 1
        If lstSheets.Selected(i) = True Then
            X = X + 1
        End If
    Next i
    If X = 0 Then
        MsgBox "No sheets were selected." & vbCrLf & "Please make a selection to continue.", vbCritical, "Invalid Selection"
        Exit Sub
    End If
    Call shtSelect
    SheetFormatting
    Me.Hide
    ActiveWindow.SelectedSheets.PrintPreview
    Sheets(sSht).Select
    Exit Sub
End Sub

Private Sub cmdPrinter_Click()
    Application.Dialogs(xlDialogPrinterSetup).Show
    Me.txtPrinter.Value = Application.ActivePrinter
End Sub



Private Sub optDPR_Click()
    pic = "DPRLogo.25.png"
    Call loadLogo
End Sub

Private Sub optDPR_RQC_Click()
    pic = "DPR_RQC_25.png"
    Call loadLogo
End Sub

Private Sub optDPRHardin_Click()
    pic = "DPRClarkLogo.png"
    Call loadLogo
End Sub

Sub loadLogo()
Dim sImage As Shape
    For Each ows In ActiveWorkbook.Worksheets
        If ows.Name <> "EstData" And ows.Name <> "XMLTables" Then
           ows.PageSetup.LeftFooterPicture.fileName = sPth & pic
        End If
    Next ows
End Sub

Private Sub UserForm_Activate()
    bLoading = True
    optDPR_Click
    sSht = ActiveSheet.Name
    cmdCancel.Picture = Application.CommandBars.GetImageMso("PrintPreviewClose", 24, 24)
    cmdPreview.Picture = Application.CommandBars.GetImageMso("FilePrintPreview", 24, 24)
    cmdOK.Picture = Application.CommandBars.GetImageMso("FilePrint", 24, 24)
    cmdPdf.Picture = Application.CommandBars.GetImageMso("FileSaveAsPdfOrXps", 24, 24)
    Me.Caption = "Report Setup"
    txtPrinter.Value = Application.ActivePrinter
    i = 0
    With lstSheets
        .Clear
        For Each ows In ActiveWorkbook.Worksheets
            If ows.Name <> "EstData" And ows.Name <> "XMLTables" Then
                .AddItem
                .List(i, 0) = ows.Name
                .List(i, 1) = ows.Cells(1, 2).Value
                i = i + 1
            End If
        Next ows
    End With
    bLoading = False
End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 115
    Me.Left = Application.Left + 25
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Sheets(sSht).Select
End Sub

