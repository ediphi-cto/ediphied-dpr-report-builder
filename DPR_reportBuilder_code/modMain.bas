Attribute VB_Name = "modMain"
Option Explicit
#If Win64 Then
    Public Declare PtrSafe Function SetWindowPos _
        Lib "user32" ( _
            ByVal hwnd As Long, _
            ByVal hwndInsertAfter As Long, _
            ByVal X As Long, ByVal Y As Long, _
            ByVal cx As Long, ByVal cy As Long, _
            ByVal wFlags As Long) _
    As LongPtr
#Else
    Public Declare PtrSafe Function SetWindowPos _
        Lib "user32" ( _
            ByVal hwnd As Long, _
            ByVal hwndInsertAfter As Long, _
            ByVal X As Long, ByVal Y As Long, _
            ByVal cx As Long, ByVal cy As Long, _
            ByVal wFlags As Long) _
    As LongPtr
#End If

Public Const sPth As String = "\\azr-corp-store\Estimators\DPR Reporter\Modules\"
Public Const sConnStrP1 = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq="
Public Const sConnStrP2 = ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
Public Const sXpath As String = "/Estimate/ItemTable/Item"
Public Const sXpathSub As String = "/Estimate/SubcontractorTable/Subcontractor"

Public Const sSvr As String = "azr-corp-sqldb.database.windows.net"  'Dev = azr-corp-sqldb.database.windows.net
Public Const sDb As String = "DPR_MULTAPP"
Public Const sDbUser As String = "PrCUserApp"
Public Const sDbPass As String = "Q5a#6%HHnbr"

Public Const sFilter = "CSV File, *.csv"
Public rsNew As New ADODB.Recordset
Public objCon As Object
Public objRst As Object
Public objFrm As Object

Public owb As Workbook
Public ows As Worksheet
Public lObj As ListObject
Public lRow As ListRow
Public lCol As ListColumn
Public lastRow As Range
Public ptCache As PivotCache
Public pt As PivotTable
Public pf As PivotField
Public slcr As SlicerCache

Public sDir As String
Public sVarXML As String
Public sVarGuid As String
Public sRprtID As String
Public sConID As String
Public sEstID As String
Public sConnection As String
Public sConnID As String
Public sSql As String
Public sRegID As String
Public sRegTxt As String
Public sUser As String
Public sCompID As String
Public sJobUM As String
Public sGTLvl1 As String
Public sLvl  As String
Public sRprt As String
Public sXTRow As String
Public sMod As String
Public sFormula As String
Public sAxisX As String
Public sAxisY As String
Public cObj As ChartObject
Public cChr As Chart

Public bCkbAll As Boolean
Public bCkbSub As Boolean
Public bCkb0 As Boolean
Public bCkb1 As Boolean
Public bCkb2 As Boolean
Public bCkb3 As Boolean
Public bCkb4 As Boolean
Public bCkb5 As Boolean
Public sXpath0 As String
Public sXpath1 As String
Public sXpath2 As String
Public sXpath3 As String
Public sXpath4 As String
Public sXpath5 As String
Public sLUnit As String
Public sLvl0 As String
Public sLvl1 As String
Public sLvl2 As String
Public sLvl3 As String
Public sLvl4 As String
Public sLvl5 As String
Public sLvlSub As String
Public sLvl0Unit As String
Public sLvl0xNd As String
Public sLvl1xNd As String
Public sLvl2xNd As String
Public sLvl3xNd As String
Public sLvl4xNd As String
Public sLvl5xNd As String
Public sLvl0Item As String
Public sLvl1Item As String
Public sLvl2Item As String
Public sLvl3Item As String
Public sLvl4Item As String
Public sLvl5Item As String
Public sLvl0Code As String
Public sLvl1Code As String
Public sLvl2Code As String
Public sLvl3Code As String
Public sLvl4Code As String
Public sLvl5Code As String
Public sLvl0Name As String
Public sLvl1Name As String
Public sLvl2Name As String
Public sLvl3Name As String
Public sLvl4Name As String
Public sLvl5Name As String
Public sColName As String
Public sTblName As String
Public pic As String
Public sSht As String
Public sVal1 As String
Public sVal2 As String
Public sVal3 As String
Public sCol1 As String
Public sCol2 As String
Public sNewEst As String
Public sOldEst As String
Public sFilePath As String
Public sDuration As String
Public cmdPath As String
Public bConnect As Boolean
Public bAddon As Boolean
Public bPdf As Boolean
Public bPvt As Boolean
Public bAdd As Boolean
Public bHasData As Boolean
Public bCode As Boolean
Public bFileSave As Boolean

Public cl As Range
Public clLeft As Double
Public clTop As Double
Public clWidth As Double
Public clHeight As Double
Public myShape As Shape
Public lngCol As Long

Public iCount As Integer
Public iCol As Integer
Public iRow As Long
Public C As Long
Public i As Integer
Public r As Integer
Public X As Integer
Public Y As Integer
Public z As Integer
Public ptc As Integer
Public iAddRow As Long
Public iGTCol As Integer
Public iGTRow As Long
Public iLvl As Integer
Public dJobSz As Double

Sub LoadWorkbook()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    sSht = ActiveSheet.Name
    'Call LoadXMLTables
    'Call LaborHrProjection
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    frmReportLevel.Show
End Sub

Sub Save_WB()
Dim fileName As String
Dim fileSavename As String
    sUser = CStr(Environ("USERNAME"))
    fileName = FNameScrub(Range("rngProjectName").Value)
    fileName = sDir & fileName
    fileSavename = Application.GetSaveAsFilename(fileName, FileFilter:="xlsm Files (*.xlsm), *.xlsm")
    If fileSavename <> "False" Then
        Range("rngIsTemp").Value = True
        Range("rngCP1").Value = True
        ActiveWorkbook.SaveAs fileSavename
        bFileSave = True
    Else
        Exit Sub
    End If
End Sub


'Sub NewReport() 'Generate new report
'Dim wsh As Object
'Dim strPath As String
'Dim cmdLine As String
'Dim sWB As String
'Dim sXML As String
'Dim Ret, countOf
'Dim file As String, count As Long
'Dim waitOnReturn As Boolean: waitOnReturn = True
'Dim windowStyle As Integer: windowStyle = 1
'sDir = DriveName(ThisWorkbook.Path)
'sUser = CStr(Environ("USERNAME"))
'cmdPath = "C:\Users\" & sUser & "\AppData\Local\Temp\DPRReporter\"
'    Sheet2.Activate
'    If Dir(cmdPath, vbDirectory) = "" Then
'        MkDir cmdPath
'    End If
'    sWB = cmdPath & "DPR Report Builder for WinEst.xlsm"
'    sXML = cmdPath & "ReportTables.xml"
'    If Dir(sXML) <> "" Then
'        Kill sXML
'    End If
'    Set wsh = VBA.CreateObject("WScript.Shell")
'    strPath = Chr(34) & "C:\Program Files (x86)\WinEst\winest.exe" & Chr(34)
'    cmdLine = " /x /notallitems /emptyfields /tpl DPRTpl.xml " & Chr(34) & cmdPath & "ReportTables.xml" & Chr(34)
'    If Len(strPath) > 0 Then
'        wsh.Run strPath & cmdLine, windowStyle, waitOnReturn
'    Else
'        MsgBox "Path doesn't exist"
'    End If
'    If Dir(sWB) <> "" Then
'        Ret = IsWorkBookOpen(sWB)
'        If Ret = True Then
'            file = Dir$(cmdPath & "DPR Excel Report Builder*.xlsm")
'            Do Until file = ""
'                countOf = (countOf + 1)
'                file = Dir$()
'            Loop
'            sWB = cmdPath & "DPR Excel Report Builder-" & countOf & ".xlsm"
'        Else
'            Kill sWB
'        End If
'    End If
'    CleanTempFolder
'    Application.DisplayAlerts = False
'    ThisWorkbook.SaveAs sWB
'    Application.DisplayAlerts = True
'    Call LoadWorkbook
'End Sub
'
'Sub VarReport() 'Generate Variance XML file (Estimate 2)
'Dim fileData As String
'Dim fol As String
'Dim wsh As Object
'Dim strPath As String
'Dim sTempFldr As String
'Dim cmdLine As String
'Dim sWB As String
'Dim sXML As String
'Dim Ret, countOf
'Dim file As String, count As Long
'Dim waitOnReturn As Boolean: waitOnReturn = True
'Dim windowStyle As Integer: windowStyle = 1
'sDir = DriveName(ThisWorkbook.Path)
'sUser = CStr(Environ("USERNAME"))
'
''    fol = ThisWorkbook.Path & "\ReportData\"
'    sVarXML = GetFilenameFromPath(cmdPath)
'    fileData = FNameScrub(sVarXML)
'    Set wsh = VBA.CreateObject("WScript.Shell")
'    sTempFldr = "C:\Users\" & sUser & "\AppData\Local\Temp\DPRReporter\"
''Kill Temp Variance XML file
'    sXML = sTempFldr & "VarReportTables.xml"
'    If Dir(sXML) <> "" Then
'        Kill sXML
'    End If
'
'    strPath = Chr(34) & "C:\Program Files (x86)\WinEst\winest.exe" & Chr(34)
''1st command line: This will open the selected estimate. The command line doesn't always execute fully
'    cmdLine = "/uAdministrator /p /x /notallitems /emptyfields /tpl DPRTpl.xml " & Chr(34) & cmdPath & Chr(34) & " " & Chr(34) & sTempFldr & "VarReportTables.xml" & Chr(34)
'    If Len(strPath) > 0 Then
'        wsh.Run strPath & cmdLine, windowStyle, waitOnReturn
'    End If
''2nd command line:This will run the XML command again from the estimate that was opened above.
'    cmdLine = "/x /notallitems /emptyfields /tpl DPRTpl.xml " & Chr(34) & sTempFldr & "VarReportTables.xml" & Chr(34)
'    If Len(strPath) > 0 Then
'        wsh.Run strPath & cmdLine, windowStyle, waitOnReturn
'        Range("rngVarEstID").Value = fileData
'    Else
'        MsgBox "Path doesn't exist"
'    End If
'    Call xmlVarTotalsTable
'End Sub

Sub FileSaveAs()
Dim fileName As String
Dim fileSavename As String
Dim fileXML As String
Dim fileVarXML As String
Dim fileData As String
Dim fol As String
Dim ans
Dim fso As Object
    sDir = DriveName(ThisWorkbook.Path)
    sUser = CStr(Environ("USERNAME"))
    cmdPath = sDir & "\Users\" & sUser & "\AppData\Local\Temp\DPRReporter\"
    fileName = FNameScrub(Range("rngProjectName").Value)
    fileXML = cmdPath & "ReportTables.xml"
    fileVarXML = cmdPath & "VarReportTables.xml"
    fileData = "\" & Replace(FNameScrub(Range("rngEstimateID").Value), ".est", ".xml")
    fileName = "\\tsclient\C\Users\" & sUser & "\OneDrive - DPR Construction\Documents\" & fileName
    fileSavename = Application.GetSaveAsFilename(fileName, FileFilter:="xlsm Files (*.xlsm), *.xlsm")
    If fileSavename <> "False" Then
        Range("rngDataBase").Value = fileData
        Range("rngVarReport").Value = "\" & Replace(Range("rngVarEstID").Value, ".est", ".xml")
        Range("rngCP1").Value = True
        ActiveWorkbook.SaveAs fileSavename
    Else
        Exit Sub
    End If
    fol = ThisWorkbook.Path & "\ReportData"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(fol) Then
        fso.CreateFolder (fol)
    End If
    On Error Resume Next
    If bFileSave = False Then
        fileData = fol & fileData
        FileCopy fileXML, fileData
        FileCopy fileVarXML, fol & Range("rngVarReport").Value
    End If
    On Error GoTo 0
    MsgBox "Workbook Saved.", vbOKOnly, "DPR Reporter"
End Sub

'Sub ReLoadWorkbook()
'    sSht = ActiveSheet.Name
'    With Excel.Application
'       .ScreenUpdating = False
'       .EnableEvents = False
'       .Calculation = Excel.xlCalculationManual
'    End With
'    Call clearAllAddons
'
'
'    Call OLEDB_Text
'    Call RefreshPivots
'    Call scrubNotes
'    SummaryDetail
'    Call ReApplyAddons
'    ExecSummary
'    Sheet2.Activate
'ErrorFoundResetApplication:
'    With Excel.Application
'       .ScreenUpdating = True
'       .EnableEvents = True
'       .Calculation = Excel.xlCalculationAutomatic
'    End With
'End Sub

Sub SheetFormatting()
    On Error Resume Next
    Set ows = ActiveSheet
        If ows.CodeName = "Sheet2" Then
            ows.PageSetup.PrintArea = ows.Range("$B$1:$K$56").Address
        ElseIf ows.CodeName = "Sheet3" Then
            ows.PageSetup.PrintArea = ActualUsedRange(ows).Address
            ows.PageSetup.PrintTitleRows = "$1:$7"
        Else
            Set pt = ows.PivotTables(1)
            ows.PageSetup.PrintArea = ActualUsedRange(ows).Address
            If InStr(pt.Name, "XTab") Then
                ows.PageSetup.PrintTitleRows = "$1:$12"
            Else
                ows.PageSetup.PrintTitleRows = "$1:$13"
            End If
            ows.PageSetup.FitToPagesWide = 1
            ows.PageSetup.FitToPagesTall = False
            If InStr(pt.Name, "XTab") Or InStr(pt.Name, "ControlEstimate") Or InStr(pt.Name, "Variance") Then
                ows.PageSetup.Orientation = xlLandscape
            End If
        End If
    On Error GoTo 0
End Sub

Sub SheetFormatingAll()
    On Error Resume Next
    For Each ows In ActiveWorkbook.Worksheets
        If ows.CodeName = "Sheet2" Then
            ows.PageSetup.PrintArea = ows.Range("$B$1:$K$56").Address
        ElseIf ows.CodeName = "Sheet3" Then
            ows.PageSetup.PrintArea = ActualUsedRange(ows).Address
            ows.PageSetup.PrintTitleRows = "$1:$7"
        Else
            Set pt = ows.PivotTables(1)
            ows.PageSetup.PrintArea = ActualUsedRange(ows).Address
            If InStr(pt.Name, "XTab") Then
                ows.PageSetup.PrintTitleRows = "$1:$12"
            Else
                ows.PageSetup.PrintTitleRows = "$1:$13"
            End If
            ows.PageSetup.FitToPagesWide = 1
            ows.PageSetup.FitToPagesTall = False
            If InStr(pt.Name, "XTab") Or InStr(pt.Name, "ControlEstimate") Or InStr(pt.Name, "Variance") Then
                ows.PageSetup.Orientation = xlLandscape
            End If
        End If
    Next ows
    On Error GoTo 0
End Sub

Sub PageSetup()
    On Error GoTo HandleErrors
    Set ows = ActiveSheet
    pic = "DPRLogo.25.png"
    With ows.PageSetup
        .LeftFooterPicture.fileName = sPth & pic
        .LeftFooter = "&G"
        .CenterFooter = "Page &P of &N"
        .RightFooter = Range("rngEstName").text
    End With
    Application.PrintCommunication = False
    With ows.PageSetup
        .LeftMargin = Application.InchesToPoints(0.15)
        .RightMargin = Application.InchesToPoints(0.15)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.25)
        .FooterMargin = Application.InchesToPoints(0.15)
        If InStr(ows.Name, "Control Estimate") = 1 Or InStr(ows.Name, "Variance") = 1 Then
            .Orientation = xlLandscape
        Else
            .Orientation = xlPortrait
        End If
        .FitToPagesWide = 1
        .FitToPagesTall = False
     End With
    Application.PrintCommunication = True
ExitHere:
    Exit Sub
    
HandleErrors:
    Application.PrintCommunication = True
    Resume ExitHere
End Sub

Sub ResetSheetScroll()
Dim lLastRow As Long, lLastColumn As Long
Dim lRealLastRow As Long, lRealLastColumn As Long
    Set ows = ActiveSheet
    With ows.Range("A1").SpecialCells(xlCellTypeLastCell)
        lLastRow = .row
        lLastColumn = .Column
    End With
    lRealLastRow = ows.Cells.Find("*", ows.Range("A1"), xlFormulas, , xlByRows, xlPrevious).row
    lRealLastColumn = ows.Cells.Find("*", ows.Range("A1"), xlFormulas, , xlByColumns, xlPrevious).Column

    If lRealLastRow < lLastRow Then
        ows.Range(ows.Cells(lRealLastRow + 1, 1), ows.Cells(lLastRow, 1)).EntireRow.Delete
    End If
    If lRealLastColumn < lLastColumn Then
        ows.Range(ows.Cells(1, lRealLastColumn + 1), ows.Cells(1, lLastColumn)) _
        .EntireColumn.Delete
    End If
    ActiveSheet.UsedRange   'resets LastCell
End Sub

Sub CopyToWb()
Dim sTheme As String
Dim slcr As SlicerCache
Dim lvl As Integer
Dim C
    Set ows = ActiveSheet
    sTheme = Range("rngTheme").Value
    If ows.Name = "Executive Summary" Or ows.Name = "Systems Summary" Then
        ActiveSheet.Copy
        ActiveWorkbook.ApplyTheme (sTheme)
    Else
        Set pt = ows.PivotTables(1)
        sRprt = ows.PivotTables(1).Name
        Set lObj = Sheet0.ListObjects("tblRptTrack")
        Set C = lObj.ListColumns(1).DataBodyRange.Find(sRprt, LookIn:=xlValues)
        If Not C Is Nothing Then
            lvl = C.Offset(0, 1).Value
        End If
        iAddRow = ActualUsedRange(ows).Rows.count
        With pt.TableRange1
            iGTRow = .Cells(.Cells.count).row
            iCol = .Cells(.Cells.count).Column
        End With
        On Error Resume Next
        ActiveSheet.Range(Cells(iGTRow + 1, 2), Cells(iAddRow, iCol)).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues
        ActiveSheet.Copy
        ActiveSheet.PivotTables(1).PivotSelect "", xlDataAndLabel, True
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        ActiveSheet.Shapes.Range(Array("grpHeading")).Select
        Selection.OnAction = ""
        For Each slcr In ActiveWorkbook.SlicerCaches
            slcr.Delete
        Next slcr
        ActiveWorkbook.ApplyTheme (sTheme)
        If InStr(1, sRprt, "Level") > 0 Or InStr(1, sRprt, "Backup") > 0 Then
            Select Case lvl
                Case 1
                    ConvertFxL1
                Case 2
                    ConvertFxL2
                Case 3
                    ConvertFxL3
                Case 4
                    ConvertFxL4
                Case 5
                    ConvertFxL5
            End Select
        End If
        Range("A1").Select
        On Error GoTo 0
    End If
End Sub

Sub RefreshPivots()
Dim pt As PivotTable
Dim pc As PivotCache
    For Each ows In ActiveWorkbook.Worksheets
        For Each pt In ows.PivotTables
            With ows.PivotTables(1).PivotCache
                Set .Recordset = rsNew
                    .Refresh
                End With
            pt.PivotCache.MissingItemsLimit = xlMissingItemsNone
            pt.PivotCache.Refresh
        Next pt
    Next ows
End Sub

Sub FormatCntrlEst()
    Set ows = ActiveSheet
    Set pt = ows.PivotTables(1)
    With pt.RowRange
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlExpression, Formula1:=sFormula
        .FormatConditions(.FormatConditions.count).SetFirstPriority
        .FormatConditions(1).StopIfTrue = False
        With .FormatConditions(1).Font
            .Bold = False
            .Italic = True
            .TintAndShade = 0
        End With
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
    End With
End Sub

Sub FilterByIcon()
    Set ows = ActiveSheet
    Set pt = ows.PivotTables(1)
    With pt.TableRange1
        .AutoFilter
        .AutoFilter Field:=5, _
        Criteria1:=ActiveWorkbook.Colors(65535)
    End With
End Sub

Sub ReportTrack()
    Set ows = Sheet0
    Set lObj = ows.ListObjects("tblRptTrack")
    Set lRow = lObj.ListRows.Add
    With lRow
        .Range(1, 1).Value = sSht
        .Range(1, 2).Value = iLvl
        .Range(1, 3).Value = bCkbAll
        .Range(1, 4).Value = bCkbSub
        .Range(1, 5).Value = sXpath0
        .Range(1, 6).Value = bCkb0
        .Range(1, 7).Value = sLvl0xNd
        .Range(1, 8).Value = sLvl0Code
        .Range(1, 9).Value = sLvl0Item
        .Range(1, 10).Value = sXpath1
        .Range(1, 11).Value = bCkb1
        .Range(1, 12).Value = sLvl1xNd
        .Range(1, 13).Value = sLvl1Code
        .Range(1, 14).Value = sLvl1Item
        .Range(1, 15).Value = sXpath2
        .Range(1, 16).Value = bCkb2
        .Range(1, 17).Value = sLvl2xNd
        .Range(1, 18).Value = sLvl2Code
        .Range(1, 19).Value = sLvl2Item
        .Range(1, 20).Value = sXpath3
        .Range(1, 21).Value = bCkb3
        .Range(1, 22).Value = sLvl3xNd
        .Range(1, 23).Value = sLvl3Code
        .Range(1, 24).Value = sLvl3Item
        .Range(1, 25).Value = sXpath4
        .Range(1, 26).Value = bCkb4
        .Range(1, 27).Value = sLvl4xNd
        .Range(1, 28).Value = sLvl4Code
        .Range(1, 29).Value = sLvl4Item
        .Range(1, 30).Value = sXpath5
        .Range(1, 31).Value = bCkb5
        .Range(1, 32).Value = sLvl5xNd
        .Range(1, 33).Value = sLvl5Code
        .Range(1, 34).Value = sLvl5Item
    End With
End Sub

Sub clearStrings()
    sSht = ""
    iLvl = 0
    bCkbAll = False
    sXpath0 = ""
    bCkb0 = False
    sLvl0xNd = ""
    sLvl0Code = ""
    sLvl0Item = ""
    sXpath1 = ""
    bCkb1 = False
    sLvl1xNd = ""
    sLvl1Code = ""
    sLvl1Item = ""
    sXpath2 = ""
    bCkb2 = False
    sLvl2xNd = ""
    sLvl2Code = ""
    sLvl2Item = ""
    sXpath3 = ""
    bCkb3 = False
    sLvl3xNd = ""
    sLvl3Code = ""
    sLvl3Item = ""
    sXpath4 = ""
    bCkb4 = False
    sLvl4xNd = ""
    sLvl4Code = ""
    sLvl4Item = ""
    sXpath5 = ""
    bCkb5 = False
    sLvl5xNd = ""
    sLvl5Code = ""
    sLvl5Item = ""
End Sub

Sub ReportHeadings()
    frmHeadings.Show (0)
End Sub

Sub loadImage()
    Dim sDir As String
    Dim lResult As String
    Dim filters As String
    Dim fileName As Variant
    Dim cl As Range
    Dim clLeft As Double
    Dim clTop As Double
    Dim clWidth As Double
    Dim clHeight As Double

    Set cl = Range("$G$8:$K$22")
    clLeft = cl.Left
    clTop = cl.Top
    clHeight = cl.Height
    clWidth = cl.Width

    lResult = CurDir()
    sUser = Environ("UserName")
    
    If Left(LCase(lResult), 11) = "\\client\c$" Then
        sDir = "\\Client\c$\Users\" & sUser & "\"
    Else
        sDir = "C:\Users\" & sUser & "\"
    End If
    ChDir (sDir)
    
    filters = "Image Files,*.bmp;*.tif;*.jpg;*.png,BMP (*.bmp),*.bmp,PNG (*.png),*.png,TIFF (*.tif),*.tif,JPG (*.jpg),*.jpg"
    fileName = Application.GetOpenFilename(filters, 0, "Select Project Image", , False)
    If fileName = False Then Exit Sub
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, clLeft, clTop, clWidth, clHeight).Select
    Selection.ShapeRange.Fill.visible = msoFalse
    Selection.ShapeRange.Line.visible = msoFalse
    Selection.ShapeRange.Name = "ProjectImage"
    With Selection.ShapeRange.Fill
        .visible = msoTrue
        .UserPicture fileName
        .TextureTile = msoFalse
        .RotateWithObject = msoTrue
    End With
End Sub





