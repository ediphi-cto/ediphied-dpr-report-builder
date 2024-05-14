Attribute VB_Name = "modFunc"
Option Base 0
Option Explicit
Public svrPath As String

Sub CreateTempXML()
Dim fso As Object

sUser = CStr(Environ("USERNAME"))
cmdPath = "C:\Users\" & sUser & "\AppData\Local\Temp\DPRReporter\"
svrPath = "\\azr-corp-store\Estimating\WinEst\API\XML\" & Range("rngGUID").Value & ".xml"
    If Dir(cmdPath, vbDirectory) = "" Then
        MkDir cmdPath
    End If
    On Error GoTo errHndlr
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile svrPath, cmdPath & Range("rngGUID").Value & ".xml"
    Set fso = Nothing
    Exit Sub
errHndlr:
    MsgBox "Unable to connect to Server. Close this workbook and check your VPN connection and try again.", vbCritical, "Unable to connect"
    Exit Sub
    On Error GoTo 0
End Sub

Sub KillTempXML()
On Error Resume Next
    sUser = CStr(Environ("USERNAME"))
    cmdPath = "C:\Users\" & sUser & "\AppData\Local\Temp\DPRReporter\"
    Kill cmdPath & Range("rngGUID").Value & ".xml"
On Error GoTo 0
End Sub

Function GetWorkbookDirectory() As String

    Dim sPath As String
    Dim sOneDrive As String
    Dim iPos As Integer
    
    sPath = Application.ActiveWorkbook.Path
    
    ' Is this a OneDrive path?
    If Left(sPath, 6) = "https:" Then
        ' Find the start of the "local part" of the name
        iPos = InStr(sPath, "//")           ' Find start of URL hostname
        iPos = InStr(iPos + 2, sPath, "/")  ' Find end of URL hostname
        iPos = InStr(iPos + 1, sPath, "/")  ' Find start of local part
        iPos = InStr(iPos + 1, sPath, "/")  ' Find start of local part
        iPos = InStr(iPos + 1, sPath, "/")  ' Find start of local part
        ' Join that with the local location for OneDrive files
        sPath = Environ("OneDrive") & Mid(sPath, iPos)
        sPath = Replace(sPath, "/", Application.PathSeparator)
    End If
    
    GetWorkbookDirectory = sPath

End Function

Function FileOpen(initialFilename As String, _
  Optional sDesc As String = "Excel (*.xlsm)", _
  Optional sFilter As String = "*.xlsm") As String
  With Application.FileDialog(msoFileDialogOpen)
    .ButtonName = "&Open"
    .initialFilename = initialFilename
    .filters.Clear
    .filters.Add sDesc, sFilter, 1
    .Title = "File Open"
    .AllowMultiSelect = False
    If .Show = -1 Then FileOpen = .SelectedItems(1)
  End With
End Function

Function DriveName(sPath As String) As String
    DriveName = CreateObject("Scripting.FileSystemObject").GetDriveName(sPath)
End Function

Function Encode(strXML As String) As String
    Encode = Replace(Replace(Replace(strXML, "&", "&amp;"), "<", "&lt;"), ">", "&gt;")
End Function

Sub XL_SQL(ByVal sSht As String, ByVal sRprt As String, ByVal sSql As String)
    
    Set ows = Sheet0
    Set lObj = ows.ListObjects(1)
    Set lRow = lObj.ListRows.Add
    With lRow
        .Range(1, 1).Value = sSht
        .Range(1, 2).Value = sRprt
        .Range(1, 3).Value = sSql
    End With
End Sub

Sub CleanTempFolder()
    Dim file As String, count As Long
    Dim countOf, Ret
    On Error Resume Next
    sDir = DriveName(ThisWorkbook.Path)
    file = Dir$(cmdPath & "DPR Report Builder*.xlsm")
    Do Until file = ""
        Ret = IsWorkBookOpen(file)
        If Ret = True Then
            
        Else
            Kill (file)
        End If
        file = Dir$()
    Loop
    On Error GoTo 0
End Sub

Function IsWorkBookOpen(fileName As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open fileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else:
    End Select
End Function

Function GetFilenameFromPath(ByVal strPath As String) As String
' Returns the rightmost characters of a string upto but not including the rightmost '\'
' e.g. 'c:\winnt\win.ini' returns 'win.ini'

    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function

Function getPfName(pfName As String)
Set pt = ActiveSheet.PivotTables(1)
    For Each pf In pt.RowFields
        If pf.Position = i Then
            pfName = pf.SourceName
            Exit For
        End If
    Next pf
    rf = pfName
End Function

Function getCfName(cfName As String)
Set pt = ActiveSheet.PivotTables(1)
    For Each pf In pt.ColumnFields
        If pf.Position = 1 Then
            cfName = pf.SourceName
            Exit For
        End If
    Next pf
    rf = cfName
End Function

Function getPtCount() As Integer
    getPtCount = Range("RprtID").Value + 1
    Range("RprtID").Value = getPtCount
End Function

Function xmlPath() As String
    If Sheet1.Range("rngDataBase").Value = "" Or bRefresh = True Then
        xmlPath = "C:\Users\" & sUser & "\AppData\Local\Temp\DPRReporter\ReportTables.xml"
    Else
        xmlPath = Application.ThisWorkbook.Path & "\ReportData" & Range("rngDataBase").Value
    End If
End Function

'Function xmlPath() As String ' When using xml files with GUID's
'sUser = CStr(Environ("USERNAME"))
'Reload:
'    If Range("rngIsTemp").Value = False Then
'        cmdPath = "\\azr-corp-store\Estimating\WinEst\API\XML\"
'    Else
'        cmdPath = "C:\Users\" & sUser & "\AppData\Local\Temp\DPRReporter\"
'    End If
'    xmlPath = cmdPath & Range("rngGUID").Value & ".xml"
'    If FileExists(xmlPath) = False Then
'        CreateTempXML
'        GoTo Reload
'    End If
'End Function

Function xmlTemp() As String
    xmlTemp = "\\azr-corp-store\Estimating\WinEst\API\XML\" & Range("rngGUID").Value & ".xml"
End Function

Function FileExists(FilePath As String) As Boolean
Dim TestStr As String
    TestStr = ""
    On Error Resume Next
    TestStr = Dir(FilePath)
    On Error GoTo 0
    If TestStr = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function

Function xmlVarPath() As String
    If Sheet1.Range("rngVarReport").Value = "" Or bRefresh = True Then
        xmlVarPath = "C:\Users\" & sUser & "\AppData\Local\Temp\DPRReporter\VarReportTables.xml"
    Else
        xmlVarPath = Application.ThisWorkbook.Path & "\ReportData" & Range("rngVarReport").Value
    End If

'    xmlVarPath = "\\azr-corp-store\Estimating\WinEst\API\XML\" & Range("rngVarGUID").Value & ".xml"
'    If FileExists(xmlVarPath) = False Then
''        CreateTempXML
''        GoTo Reload
'    End If
End Function

Function GetWorksheetFromCodeName(CodeName As String) As Worksheet
    For Each ows In ThisWorkbook.Worksheets
        If StrComp(ows.CodeName, CodeName, vbTextCompare) = 0 Then
            Set GetWorksheetFromCodeName = ows
            Exit Function
        End If
    Next ows
End Function

Function ptExists() As Boolean
    On Error Resume Next
    ptExists = Not (IsError(ows.PivotTables(1)))
    On Error GoTo 0
End Function

Function colW() As Integer
    Select Case iLvl
        Case 1
            colW = 6
        Case 2
            colW = 2.88
        Case 3
            colW = 1.25
        Case 4
            colW = 1
        Case 5
            colW = 0.75
    End Select
End Function

Function rtfScrub()
Dim regexObject As RegExp
Dim rtfText As String
Dim X
    X = 0
    Set regexObject = New RegExp
    With regexObject
        .Pattern = "({\\)(.+?)(})|(\\)(.+?)(\b)"
        .IgnoreCase = True
        .Global = True
    End With
    On Error GoTo errHndlr
    pt.ManualUpdate = True
    With pt.RowFields("Comments")
        For i = 1 To .PivotItems.count
            If .PivotItems(i) <> "" Then
                rtfText = LTrim(Replace(Replace(Replace(regexObject.Replace(.PivotItems(i), ""), "}", ""), Chr(13), ""), Chr(10), ""))
                If rtfText = "" Then
                    .PivotItems(i).Value = Space(X) & "-"
                Else
                    .PivotItems(i).Value = Space(X) & rtfText
                End If
            End If
            
        Next
    End With
    pt.ManualUpdate = False
    On Error GoTo 0
    Exit Function
errHndlr:
    X = X + 1
    If X >= 20 Then Exit Function
    Resume
End Function

Function OLEDB_Text()
Dim odbcCn, cn
    For Each cn In ThisWorkbook.Connections
        If cn.Type = xlConnectionTypeOLEDB Then
            Set odbcCn = cn.OLEDBConnection
            sSql = odbcCn.CommandText
            sSql = Replace(sSql, sOldEst, sNewEst)
            odbcCn.CommandText = sSql
            odbcCn.Refresh
        End If
    Next
    For Each ows In ActiveWorkbook.Worksheets
        If bFieldItemExists("Comments") = True Then
            For Each pt In ows.PivotTables
              rtfScrub
            Next pt
        End If
    Next ows
End Function

Function scrubNotes()
    For Each ows In ActiveWorkbook.Worksheets
        If bFieldItemExists("Comments") = True Then
            For Each pt In ows.PivotTables
              rtfScrub
            Next pt
        End If
    Next ows
End Function

Function bFieldItemExists(strName As String) As Boolean
  Dim strTemp As String
  On Error Resume Next
  strTemp = ows.PivotTables(1).PivotFields(strName)
  If Err = 0 Then bFieldItemExists = True Else bFieldItemExists = False
End Function

Function ActualUsedRange(MySheet As Worksheet) As Range
    Dim FirstCell As Range, LastCell As Range
    On Error GoTo ErrorHandler
    
    With MySheet
        Set LastCell = .Cells(.Cells.Find(What:="*", SearchOrder:=xlRows, _
            SearchDirection:=xlPrevious, LookIn:=xlValues).row, _
            .Cells.Find(What:="*", SearchOrder:=xlByColumns, _
            SearchDirection:=xlPrevious, LookIn:=xlValues).Column)
        Set FirstCell = .Cells(.Cells.Find(What:="*", After:=LastCell, SearchOrder:=xlRows, _
            SearchDirection:=xlNext, LookIn:=xlValues).row, _
            .Cells.Find(What:="*", After:=LastCell, SearchOrder:=xlByColumns, _
            SearchDirection:=xlNext, LookIn:=xlValues).Column)
        Set ActualUsedRange = .Range(FirstCell, LastCell)
    End With
    Exit Function
ErrorHandler:
    Set ActualUsedRange = MySheet.Range("A1")
End Function

Function TableExists(ws As Worksheet, tblNam As String) As Boolean
Dim oTbl As ListObject
For Each oTbl In ws.ListObjects
    If oTbl.Name = tblNam Then
        TableExists = True
        Exit Function
    End If
Next oTbl
TableExists = False
End Function

Function LevelQty(ws As Worksheet, tblNam As String) As Boolean
Dim oTbl As ListObject
Dim X As Long

    Set oTbl = ws.ListObjects(tblNam)
    i = 0
    For X = 1 To oTbl.DataBodyRange.Rows.count
      If oTbl.ListColumns("LevelQuantity").DataBodyRange.Rows(X).Value <> "" Then
          i = i + 1
      End If
    Next X
    If i = 0 Then
        LevelQty = False
        Exit Function
    End If
    LevelQty = True
End Function

Function projScrub(sFName As String)
Dim filt, char As String
    filt = ""
    For i = 1 To Len(sFName)
        char = Mid(sFName, i, 1)
        Select Case char
        Case "A" To "z", "a" To "z", 0 To 9, "", "!", "?", "-", " ", "\", "/"
            filt = filt & char
        End Select
    Next i
    sFName = filt & " Procurement Matrix.xlsm"
End Function

Function FNameScrub(sFName As String) As String
Dim filt, char As String
    filt = ""
    For i = 1 To Len(sFName)
        char = Mid(sFName, i, 1)
        Select Case char
        Case "A" To "z", "a" To "z", 0 To 9, "", "!", "?", "-", " ", "."
            filt = filt & char
        End Select
    Next i
    FNameScrub = filt
End Function

Sub LastColumn()
'Dim LastColumn As Long
    Set ows = Sheet0
    With ows.UsedRange
        lngCol = .Columns(.Columns.count).Column
    End With
    MsgBox lngCol
End Sub

Function ParseDateTime(dt As String) As Date
   ParseDateTime = DateValue(Left(dt, 10)) + TimeValue(Replace(Mid(dt, 12, 8), "-", ":"))
End Function

Function StripChar(r As String) As String
    With CreateObject("vbscript.regexp"): .Pattern = "[A-Za-z]": .Global = True: StripChar = .Replace(r, ""): End With
End Function

