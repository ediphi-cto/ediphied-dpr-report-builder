Attribute VB_Name = "modRibbon"
Option Explicit
Public Rib As IRibbonUI
Public MyTag As String
Public bsTag As String
Public rf As String

'************************************************************************************************************
'STANDARD RIBBON CONTROLS************************************************************************************
'************************************************************************************************************
Sub GetVisible(control As IRibbonControl, ByRef visible)
    If control.Tag Like MyTag Then
        visible = True
    Else
        visible = False
    End If
End Sub

Sub RefreshRibbon(Tag As String)
    MyTag = Tag
    If Rib Is Nothing Then
    Else
        Rib.Invalidate
    End If
End Sub

Sub RibbonUpdate()
    RefreshRibbon ("Lbl3")
End Sub

'Callback for customUI.onLoad
Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set Rib = ribbon
End Sub

'Callback for Group1 getLabel
Sub getLabel0(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Range("rngVersion").Value
End Sub

'Callback for Label1 getLabel
Sub getlabel1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Range("rngEstName").Value
End Sub

'Callback for Label2 getLabel
Sub getLabel2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Range("rngSubHeading3")
End Sub

'Callback for Label3 getLabel
Sub getLabel3(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "Last Update: " & Format(Range("rngXmlExportDate").Value, "mm/dd/yy") & "-" & Format(Range("rngXmlExportTime").Value, "h:mm:ss AM/PM")
End Sub

'Callback for backstage.onShow
'Sub OnShow(contextObject As Object)
'    'SheetFormatting
'End Sub

'GROUP 2----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Callback for g2b1 onAction
Sub g2b1a(control As IRibbonControl)
    frmReportLevel.Show
End Sub

Sub g2b1b(control As IRibbonControl)
   ' frmReportXTab.Show
End Sub

'Callback for g2b2 onAction
Sub g2b2(control As IRibbonControl)
    If ActiveSheet.name = "Executive Summary" Or ActiveSheet.name = "Systems Summary" Then
        Exit Sub
    Else
        frmReportFormat.Show
    End If
End Sub

Sub g2b1c(control As IRibbonControl)
    frmMarkups.Show
End Sub

Sub g2b3(control As IRibbonControl) 'Save to Local Drive
    FileSaveAs
End Sub

Sub g2b4(control As IRibbonControl)
    Dim fileName As String
    Dim fileSavename As String
    Dim ans
    sDir = CurDir()
    Application.ScreenUpdating = False
    Call SheetFormatting
    sUser = Environ("UserName")
    fileName = Range("rngProjectName").Value
    If sDir = "\\tsclient\C\Users\" & sUser & "\OneDrive - DPR Construction\Documents" Then
        fileName = "\\tsclient\C\Users\" & sUser & "\OneDrive - DPR Construction\Documents\" & fileName
    Else
        fileName = "C:\Users\" & sUser & "\OneDrive - DPR Construction\Documents\" & fileName
    End If
    fileSavename = Application.GetSaveAsFilename(fileName, FileFilter:="pdf Files (*.pdf), *.pdf")
    If fileSavename <> "False" Then
        bPdf = True
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileSavename, _
             Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    End If
End Sub

Sub g2b4a(control As IRibbonControl)
    MsgBox "This feature is currently not supported"
'    MsgBox "Data Refresh does not include Variance Reports." & vbCrLf & "Click OK to continue.", vbInformation, "Data Refresh"
'    Call Refresh_XML
End Sub

'GROUP 3----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Callback for g3b1 onAction
Sub g3b1(control As IRibbonControl)
    'frmProjects.Show (0)
End Sub

'Callback for g3b2 onAction
Sub g3b2(control As IRibbonControl)
   'refreshReports
End Sub

'Callback for g3b3 onAction
Sub g3b3(control As IRibbonControl)
    CopyToWb
End Sub

Sub btnSk2(control As IRibbonControl)
Dim rng As Range
    On Error GoTo errHndlr
    Set ows = ActiveSheet
    Set pt = ows.PivotTables(1)
    With pt
        .ColumnGrand = True
        .RowGrand = True
    End With
    Set rng = pt.DataBodyRange
    rng.Cells(rng.Rows.count, rng.Columns.count).ShowDetail = True
errHndlr:
    Exit Sub
End Sub

'GROUP 4----------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub g4b0(control As IRibbonControl)
    SheetFormatting
    PageSetup
End Sub

'Callback for g4b1 onAction
Sub g4b1(control As IRibbonControl)
    frmReportPrint.Show (0)
End Sub

'Callback for g4m1b1 onAction
Sub g4m1b1(control As IRibbonControl)
    Range("rngTheme").Value = "\\azr-corp-store\Estimators\DPR Reporter\Document Themes\DPR_ThemeLight.thmx"
    ActiveWorkbook.ApplyTheme ("\\azr-corp-store\Estimators\DPR Reporter\Document Themes\DPR_ThemeLight.thmx")
End Sub

'Callback for g4m1b2 onAction
Sub g4m1b2(control As IRibbonControl)
    Range("rngTheme").Value = "\\azr-corp-store\Estimators\DPR Reporter\Document Themes\DPR_ThemeDark.thmx"
    ActiveWorkbook.ApplyTheme ("\\azr-corp-store\Estimators\DPR Reporter\Document Themes\DPR_ThemeDark.thmx")
End Sub

'Callback for g4m1b3 onAction
Sub g4m1b3(control As IRibbonControl)
    Range("rngTheme").Value = "\\azr-corp-store\Estimators\DPR Reporter\Document Themes\DPR_Advanced_Tech.thmx"
    ActiveWorkbook.ApplyTheme ("\\azr-corp-store\Estimators\DPR Reporter\Document Themes\DPR_Advanced_Tech.thmx")
End Sub
    
'Callback for g4m1b4 onAction
Sub g4m1b4(control As IRibbonControl)
    Range("rngTheme").Value = "\\azr-corp-store\Estimators\DPR Reporter\Document Themes\DPR_Corporate_Office.thmx"
    ActiveWorkbook.ApplyTheme ("\\azr-corp-store\Estimators\DPR Reporter\Document Themes\DPR_Corporate_Office.thmx")
End Sub

'Callback for g4m1b5 onAction
Sub g4m1b5(control As IRibbonControl)
    Range("rngTheme").Value = "\\azr-corp-store\Estimators\DPR Reporter\Document Themes\DPR_Healthcare.thmx"
    ActiveWorkbook.ApplyTheme ("\\azr-corp-store\Estimators\DPR Reporter\Document Themes\DPR_Healthcare.thmx")
End Sub

'Callback for g4m1b6 onAction
Sub g4m1b6(control As IRibbonControl)
    Range("rngTheme").Value = "\\azr-corp-store\Estimators\DPR Reporter\Document Themes\DPR_Higher_Ed.thmx"
    ActiveWorkbook.ApplyTheme ("\\azr-corp-store\Estimators\DPR Reporter\Document Themes\DPR_Higher_Ed.thmx")
End Sub

'Callback for g4m1b7 onAction
Sub g4m1b7(control As IRibbonControl)
    Range("rngTheme").Value = "\\azr-corp-store\Estimators\DPR Reporter\Document Themes\DPR_Life_Sciences.thmx"
    ActiveWorkbook.ApplyTheme ("\\azr-corp-store\Estimators\DPR Reporter\Document Themes\DPR_Life_Sciences.thmx")
End Sub

'Callback for g4m1b8 onAction
Sub g4m1b8(control As IRibbonControl)
    Range("rngTheme").Value = "\\azr-corp-store\Estimators\DPR Reporter\Document Themes\DPR_Generic.thmx"
    ActiveWorkbook.ApplyTheme ("\\azr-corp-store\Estimators\DPR Reporter\Document Themes\DPR_Generic.thmx")
End Sub

'New Ribbon dropdown for currency
'Callback for customButton8 onAction
Sub Macro8(control As IRibbonControl)
    Range("rngNewCur_0").NumberFormat = Sheet1.ListObjects("tblCurrency").DataBodyRange(1, 2).NumberFormatLocal
    Range("rngNewCur_2").NumberFormat = Sheet1.ListObjects("tblCurrency").DataBodyRange(1, 3).NumberFormatLocal
    Call CurrencyFormat
End Sub

'Callback for customButton9 onAction
Sub Macro9(control As IRibbonControl)
    Range("rngNewCur_0").NumberFormat = Sheet1.ListObjects("tblCurrency").DataBodyRange(2, 2).NumberFormatLocal
    Range("rngNewCur_2").NumberFormat = Sheet1.ListObjects("tblCurrency").DataBodyRange(2, 3).NumberFormatLocal
    Call CurrencyFormat
End Sub

'Callback for customButton10 onAction
Sub Macro10(control As IRibbonControl)
    Range("rngNewCur_0").NumberFormat = Sheet1.ListObjects("tblCurrency").DataBodyRange(3, 2).NumberFormatLocal
    Range("rngNewCur_2").NumberFormat = Sheet1.ListObjects("tblCurrency").DataBodyRange(3, 3).NumberFormatLocal
    Call CurrencyFormat
End Sub

'Callback for customButton11 onAction
Sub Macro11(control As IRibbonControl)
    Range("rngNewCur_0").NumberFormat = Sheet1.ListObjects("tblCurrency").DataBodyRange(4, 2).NumberFormatLocal
    Range("rngNewCur_2").NumberFormat = Sheet1.ListObjects("tblCurrency").DataBodyRange(4, 3).NumberFormatLocal
    Call CurrencyFormat
End Sub

'Callback for customButton12 onAction
Sub Macro12(control As IRibbonControl)
    Range("rngNewCur_0").NumberFormat = Sheet1.ListObjects("tblCurrency").DataBodyRange(5, 2).NumberFormatLocal
    Range("rngNewCur_2").NumberFormat = Sheet1.ListObjects("tblCurrency").DataBodyRange(5, 3).NumberFormatLocal
    Call CurrencyFormat
End Sub

'Callback for customButton13 onAction
Sub Macro13(control As IRibbonControl)
    Range("rngNewCur_0").NumberFormat = Sheet1.ListObjects("tblCurrency").DataBodyRange(6, 2).NumberFormatLocal
    Range("rngNewCur_2").NumberFormat = Sheet1.ListObjects("tblCurrency").DataBodyRange(6, 3).NumberFormatLocal
    Call CurrencyFormat
End Sub

'Callback for customButton14 onAction
Sub Macro14(control As IRibbonControl)
    Range("rngNewCur_0").NumberFormat = Sheet1.ListObjects("tblCurrency").DataBodyRange(7, 2).NumberFormatLocal
    Range("rngNewCur_2").NumberFormat = Sheet1.ListObjects("tblCurrency").DataBodyRange(7, 3).NumberFormatLocal
    Call CurrencyFormat
End Sub

'Callback for customButton15 onAction
Sub Macro15(control As IRibbonControl)
    Range("rngNewCur_0").NumberFormat = Sheet1.ListObjects("tblCurrency").DataBodyRange(8, 2).NumberFormatLocal
    Range("rngNewCur_2").NumberFormat = Sheet1.ListObjects("tblCurrency").DataBodyRange(8, 3).NumberFormatLocal
    Call CurrencyFormat
End Sub

'Callback for customButton16 onAction
Sub Macro16(control As IRibbonControl)
    Range("rngNewCur_0").NumberFormat = Sheet1.ListObjects("tblCurrency").DataBodyRange(9, 2).NumberFormatLocal
    Range("rngNewCur_2").NumberFormat = Sheet1.ListObjects("tblCurrency").DataBodyRange(9, 3).NumberFormatLocal
    Call CurrencyFormat
End Sub

'Callback for getLabel Currency
Public Function getlabel10(control As IRibbonControl, ByRef Label)
    Label = Sheet1.ListObjects("tblCurrency").DataBodyRange(1, 1).Value
End Function

Public Function getlabel20(control As IRibbonControl, ByRef Label)
    Label = Sheet1.ListObjects("tblCurrency").DataBodyRange(2, 1).Value
End Function

Public Function getlabel30(control As IRibbonControl, ByRef Label)
    Label = Sheet1.ListObjects("tblCurrency").DataBodyRange(3, 1).Value
End Function

Public Function getlabel40(control As IRibbonControl, ByRef Label)
    Label = Sheet1.ListObjects("tblCurrency").DataBodyRange(4, 1).Value
End Function

Public Function getlabel50(control As IRibbonControl, ByRef Label)
    Label = Sheet1.ListObjects("tblCurrency").DataBodyRange(5, 1).Value
End Function

Public Function getlabel60(control As IRibbonControl, ByRef Label)
    Label = Sheet1.ListObjects("tblCurrency").DataBodyRange(6, 1).Value
End Function

Public Function getlabel70(control As IRibbonControl, ByRef Label)
    Label = Sheet1.ListObjects("tblCurrency").DataBodyRange(7, 1).Value
End Function

Public Function getlabel80(control As IRibbonControl, ByRef Label)
    Label = Sheet1.ListObjects("tblCurrency").DataBodyRange(8, 1).Value
End Function

Public Function getlabel81(control As IRibbonControl, ByRef Label)
    Label = Sheet1.ListObjects("tblCurrency").DataBodyRange(9, 1).Value
End Function
