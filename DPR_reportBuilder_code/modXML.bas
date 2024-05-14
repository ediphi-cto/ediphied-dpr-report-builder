Attribute VB_Name = "modXML"
Option Explicit
Public xmp As XmlMap
Public lcNewCol As ListColumn
Public wbName As String
Public XDoc As New MSXML2.DOMDocument60
Public XDocV As New MSXML2.DOMDocument60
Public xmlNodeList As MSXML2.IXMLDOMNodeList
Public oXmlNode As IXMLDOMNode
Public nd, xNode
Public xLvl0, xLvl1, xLvl2, xLvl3, xLvl4, xLvl5, xLvlSub
Public nd0, nd1, nd2, nd3, nd4, nd5, ndSub
Public oTxt As String
Public q As Variant
Public Dict As Object
Public k As Variant
Public pth As String
Public sXML As String

Function fnUpdateXMLByTags(sLocal As String)
Dim strPath As String
Dim xPath As String
Dim cmdLine As String
Dim wsh As Object
Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1
Dim fso As Object
Dim NewFile As Object
Dim FullPath As String
Dim XMLFileText As String

    FolderFromPath (sLocal)
    xPath = FolderFromPath(sLocal) & "ReportData" & Sheet1.Range("rngDataBase").Value
    FullPath = "C:\Users\" & sUser & "\AppData\Local\Temp\DPRReporter\XMLReportPath.xml"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set NewFile = fso.CreateTextFile(FullPath, 1, 1)
    
    XMLFileText = ""
    XMLFileText = XMLFileText & "<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>" & vbNewLine
    XMLFileText = XMLFileText & "<Estimate>" & vbNewLine
    NewFile.Write (XMLFileText)
   
    XMLFileText = "<EstimateInfoTable>" & vbNewLine
    NewFile.Write (XMLFileText)
    
    XMLFileText = "<EstimateInfo>" & vbNewLine
    NewFile.Write (XMLFileText)
    
    XMLFileText = "<CustomText49>" & Encode(sLocal) & "</CustomText49>" & vbNewLine
    NewFile.Write (XMLFileText)
    
    XMLFileText = "<CustomLabel49>ReportPath</CustomLabel49>" & vbNewLine
    NewFile.Write (XMLFileText)
    
    XMLFileText = "<CustomText50>" & Encode(xPath) & "</CustomText50>" & vbNewLine
    NewFile.Write (XMLFileText)
    
    XMLFileText = "<CustomLabel50>XMLReportPath</CustomLabel50>" & vbNewLine
    NewFile.Write (XMLFileText)
    
    XMLFileText = "</EstimateInfo>" & vbNewLine
    NewFile.Write (XMLFileText)
    
    
    XMLFileText = "</EstimateInfoTable>" & vbNewLine
    NewFile.Write (XMLFileText)
    
    XMLFileText = "</Estimate>" & vbNewLine
    NewFile.Write (XMLFileText)
    
    NewFile.Close

    Set wsh = VBA.CreateObject("WScript.Shell")
    strPath = Chr(34) & "C:\Program Files (x86)\WinEst\winest.exe" & Chr(34)
    cmdLine = " /m " & Chr(34) & FullPath & Chr(34)
    If Len(strPath) > 0 Then
        wsh.Run strPath & cmdLine, windowStyle, waitOnReturn
    Else
        MsgBox "Path doesn't exist"
    End If
    
End Function
Sub ImportXmlMap() 'Imports XML schema from WInEst
    sDir = DriveName(ThisWorkbook.Path)
    sUser = CStr(Environ("USERNAME"))
    cmdPath = "C:\Users\" & sUser & "\AppData\Local\Temp\DPRReporter\"
    Set xmp = ActiveWorkbook.XmlMaps.Add(cmdPath & "ReportTables.xml")
    xmp.Name = "WinEstSchema"
    Exit Sub
End Sub

Sub ImportXmlFromFile() 'DataRefresh Imports XML into an existing XML Map
Dim ans
    Application.DisplayAlerts = False
    If ReadXMLSerial(xmlPath) = False Then
        ans = MsgBox("The XML file does not match the current XML data in this workbook." & vbCrLf & "Would you like to continue with the Refresh process?", vbYesNo, "Data Validation")
        If ans = vbNo Then Exit Sub
    End If
    Set xmp = ActiveWorkbook.XmlMaps(1)
    xmp.Import xmlPath
    Application.DisplayAlerts = True
End Sub

Function ReadXMLTable(sNode As String) As Boolean
    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    X = 0
    Set xNode = XDoc.SelectNodes(sNode)
    For Each nd In xNode
        If nd.SelectSingleNode("Name").text <> "(none)" Then
            X = X + 1
            If X >= 1 Then Exit For
        End If
    Next
    If X >= 1 Then
        ReadXMLTable = True
    Else
        ReadXMLTable = False
    End If
End Function

Function ReadXMLSerial(cPth As String) As Boolean
    Dim xEstimate As MSXML2.IXMLDOMNode
    Dim xEstInfoTable As MSXML2.IXMLDOMNode
    Dim xEstInfo As MSXML2.IXMLDOMNode
    Dim xNode As IXMLDOMNode
    Dim sSerial As String
    Dim lResult As Integer
    Dim dt, tm As Date
    
    Set XDoc = New MSXML2.DOMDocument60
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    
    Set xEstimate = XDoc.DocumentElement
    Set xEstInfoTable = xEstimate.FirstChild
    Set xEstInfo = xEstInfoTable.FirstChild
    Set xNode = xEstInfoTable.FirstChild
    
    dt = CDate(xNode.SelectSingleNode("XmlExportDate").text)
    tm = CDate(xNode.SelectSingleNode("XmlExportTime").text)
    sSerial = Format(dt, "yymmdd") & "-" & Format(tm, "hhmmss")
    lResult = StrComp(sSerial, cPth)
    If lResult = 0 Then
        ReadXMLSerial = True
    Else
        ReadXMLSerial = False
    End If
End Function

Function ReadXMLFile(cFile As String) As Boolean
    Dim xEstimate As MSXML2.IXMLDOMNode
    Dim xEstInfoTable As MSXML2.IXMLDOMNode
    Dim xEstInfo As MSXML2.IXMLDOMNode
    Dim xNode As IXMLDOMNode
    Dim sFile As String
    Dim lResult As Integer
  
    
    Set XDoc = New MSXML2.DOMDocument60
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    
    Set xEstimate = XDoc.DocumentElement
    Set xEstInfoTable = xEstimate.FirstChild
    Set xEstInfo = xEstInfoTable.FirstChild
    Set xNode = xEstInfoTable.FirstChild
    
    sFile = xNode.SelectSingleNode("XmlExportDate").text
    lResult = StrComp(sFile, cFile)
    If lResult = 1 Then
        ReadXMLFile = True
    Else
        ReadXMLFile = False
    End If
End Function

Sub LoadXMLTables()
    Call xmlProjectInfo
    Call xmlTotalsTable
    Call WBSDefinition
    Call UpdateAlternates
End Sub

Sub LoadEstXml() 'Run Temp xml file to check for Estimate GUID
    Dim wsh As Object
    Dim strPath As String
    Dim cmdLine As String
    Dim cmdPath As String
    Dim sWB As String
    
    Dim Ret, countOf
    Dim file As String, count As Long
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
        
    On Error Resume Next
'Run Command line to create XML file
    sUser = CStr(Environ("USERNAME"))
    cmdPath = "C:\Users\" & sUser & "\AppData\Local\Temp\DPRReporter\"
    If Dir(cmdPath, vbDirectory) = "" Then
        MkDir cmdPath
    End If
    sXML = cmdPath & "XMLProjSetup.xml"
    If Dir(sXML) <> "" Then
        Kill sXML
    End If

    Set wsh = VBA.CreateObject("WScript.Shell")
    strPath = Chr(34) & "C:\Program Files (x86)\WinEst\winest.exe" & Chr(34)
    cmdLine = " /x /notallitems /emptyfields /tpl XMLPathTpl.xml " & Chr(34) & cmdPath & "XMLProjSetup.xml" & Chr(34)
    If Len(strPath) > 0 Then
        wsh.Run strPath & cmdLine, windowStyle, waitOnReturn
    End If
    On Error GoTo 0
End Sub
Sub loadEST_GUID() 'Load estimate GUID - Terminate if False
    Dim XDoc As MSXML2.DOMDocument60
    Dim xEstimate As MSXML2.IXMLDOMNode
    Dim xEstInfoTable As MSXML2.IXMLDOMNode
    Dim xEstInfo As MSXML2.IXMLDOMNode
    Dim xNode As IXMLDOMNode
    
'Check if estimate has GUID
    Set XDoc = New MSXML2.DOMDocument60
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load sXML
    Set xEstimate = XDoc.DocumentElement
    Set xEstInfoTable = xEstimate.FirstChild
    Set xEstInfo = xEstInfoTable.FirstChild
    Set xNode = xEstInfoTable.FirstChild

    If Len(xNode.SelectSingleNode("CustomText50").text) = 36 Or Len(xNode.SelectSingleNode("CustomText50").text) = 38 Then
        Range("rngGUID").Value = xNode.SelectSingleNode("CustomText50").text                                'Estimate GUID
        Range("rngReportPath").Value = xNode.SelectSingleNode("CustomText49").text                          'Estimate File Path
        Call Get_XML
    Else
        MsgBox "The DPR Reporter requires key estimate and project data. Open the Project Setup form and fill in all the required fields to continue.", vbCritical, "Key estimate data missing"
        Application.Quit
        ThisWorkbook.Close False
    End If
End Sub

Sub Get_XML() 'Generate WinEst XML file to Server azr-corp-store
    Dim wsh As Object
    Dim strPath As String
    Dim cmdLine As String
    Dim sXML As String
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1

    sGUID = Range("rngGUID").Value
    cmdPath = "\\AZR-CORP-STORE\Estimating\WinEst\API\XML\"
    sXML = cmdPath & sGUID & ".xml"

    Set wsh = VBA.CreateObject("WScript.Shell")
    strPath = Chr(34) & "C:\Program Files (x86)\WinEst\winest.exe" & Chr(34)
'    cmdLine = " /x /notallitems /emptyfields /tpl EstTpl.xml " & Chr(34) & sXML & Chr(34)
    cmdLine = " /x /notallitems /emptyfields /tpl DPRTpl.xml " & Chr(34) & sXML & Chr(34)
    If Len(strPath) > 0 Then
        wsh.Run strPath & cmdLine, windowStyle, waitOnReturn
    Else
        MsgBox "Path doesn't exist"
    End If
End Sub

Sub check_XMLDate() 'Check Export dates form XML
    Dim XDoc As MSXML2.DOMDocument60
    Dim xEstimate As MSXML2.IXMLDOMNode
    Dim xEstInfoTable As MSXML2.IXMLDOMNode
    Dim xEstInfo As MSXML2.IXMLDOMNode
    Dim xNode As IXMLDOMNode
    
'Check if estimate has GUID
    Set XDoc = New MSXML2.DOMDocument60
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    Set xEstimate = XDoc.DocumentElement
    Set xEstInfoTable = xEstimate.FirstChild
    Set xEstInfo = xEstInfoTable.FirstChild
    Set xNode = xEstInfoTable.FirstChild

    If CDate(xNode.SelectSingleNode("XmlExportDate").text) <> Range("rngXmlExportDate").Value Or _
       CDate(xNode.SelectSingleNode("XmlExportTime").text) <> FormatDateTime(Range("rngXmlExportTime").Value, vbLongTime) Then
        Call refreshReports
    Else
        MsgBox "The estimate data file matches the data in the DPR Reporter." & vbCrLf & "To update the data file, open this estimate in WinEst and run the Data Refresh from the Project Setup, then run the refresh from the DPR Reporter again.", vbOKOnly, "Refresh Estimate Data"
        Exit Sub
    End If
End Sub


Sub xmlProjectInfo()
    Dim xEstimate As MSXML2.IXMLDOMNode
    Dim xEstInfoTable As MSXML2.IXMLDOMNode
    Dim xEstInfo As MSXML2.IXMLDOMNode
    Dim xNode As IXMLDOMNode
    Dim sSerial As String
    Dim lResult As Integer
    Dim dt, tm As Date

    Set XDoc = New MSXML2.DOMDocument60
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    
    Set xEstimate = XDoc.DocumentElement
    Set xEstInfoTable = xEstimate.FirstChild
    Set xEstInfo = xEstInfoTable.FirstChild
    Set xNode = xEstInfoTable.FirstChild
    Set ows = Sheet1
    
    With ows
        On Error Resume Next
        .Range("rngXmlExportDate").Value = CDate(xNode.SelectSingleNode("XmlExportDate").text)
        .Range("rngXmlExportTime").Value = CDate(xNode.SelectSingleNode("XmlExportTime").text)
        .Range("rngEstSerialNo").Value = Format(.Range("rngXmlExportDate").Value, "yymmdd") _
            & "-" & Format(.Range("rngXmlExportTime").Value, "hhmmss")
        If bRefresh = False Then .Range("rngEstimateID").Value = xNode.SelectSingleNode("FileName").text
        .Range("rngEstName").Value = xNode.SelectSingleNode("CustomText11").text
        .Range("rngEstNum").Value = xNode.SelectSingleNode("CustomText12").text
        If xNode.SelectSingleNode("CustomText13").text <> "" Then
            .Range("rngEstDate").Value = CDate(xNode.SelectSingleNode("CustomText13").text)
        End If
        .Range("rngEstType").Value = xNode.SelectSingleNode("EstimateType").text
        .Range("rngEstStatus").Value = xNode.SelectSingleNode("EstimateStatus").text
        .Range("rngProjectName").Value = xNode.SelectSingleNode("ProjectName").text
        .Range("rngProjectAddress").Value = xNode.SelectSingleNode("CustomText1").text
        .Range("rngProjectCityStateZip").Value = xNode.SelectSingleNode("CustomText2").text
        .Range("rngProjectClient").Value = xNode.SelectSingleNode("CustomText22").text
        .Range("rngProjectArchitect").Value = xNode.SelectSingleNode("CustomText23").text
        .Range("rngProjectMEPEngineer").Value = xNode.SelectSingleNode("CustomText24").text
        .Range("rngEstimator").Value = xNode.SelectSingleNode("ProjectEstimator").text
        If xNode.SelectSingleNode("ProjectJobSize").text <> "" Then
            .Range("rngJobSize").Value = CDbl(xNode.SelectSingleNode("ProjectJobSize").text)
        End If
        .Range("rngJobUnitName").Value = xNode.SelectSingleNode("ProjectJobUnit").text
        .Range("rngProjectStartDate").Value = xNode.SelectSingleNode("ProjectStartDate").text
        .Range("rngProjectDuration").Value = xNode.SelectSingleNode("ProjectDuration").text
        .Range("rngJobNo").Value = xNode.SelectSingleNode("ProjectCode").text
        .Range("rngProjectType").Value = xNode.SelectSingleNode("CustomText4").text
        .Range("rngRegion").Value = xNode.SelectSingleNode("CustomText5").text
        .Range("rngHeading1").Value = .Range("rngProjectName").Value
        '.Range("rngHeading2").value = xNode.SelectSingleNode("CustomText2").Text
        .Range("rngHeading3").Value = xNode.SelectSingleNode("CustomText2").text
        .Range("rngSubHeading1").Value = "Estimate: " & xNode.SelectSingleNode("CustomText11").text
        .Range("rngSubHeading2").Value = "Project No.: " & .Range("rngJobNo").Value
        .Range("rngSubHeading3").Value = "Estimate No.: " & xNode.SelectSingleNode("CustomText12").text
        .Range("rngSubHeading4").Value = "Date: " & Format(.Range("rngEstDate").Value, "mmmm dd, yyyy")
        .Range("rngSubHeading5").Value = "Construction Area: " & Format(.Range("rngJobSize").Value, "#,##0") & " " & xNode.SelectSingleNode("ProjectJobUnit").text
        On Error GoTo 0
    End With
End Sub

Sub xmlTotalsTable() ' Build Markups(Addons) Table  **Initial Setup**
Dim arr As Variant
    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath

    Set xNode = XDoc.SelectNodes("/Estimate/TotalsPageTable/TotalsPage")
    On Error Resume Next
    ReDim arr(1 To XDoc.SelectNodes("/Estimate/TotalsPageTable/TotalsPage").Length, 1 To 11)
    For Each nd In xNode
        If nd.SelectSingleNode("Class").text <> "SUBTOTAL" Then
            If nd.SelectSingleNode("IsDeleted").text <> True Then
                If nd.SelectSingleNode("IsInactive").text <> True Then
                    C = C + 1
                    arr(C, 1) = nd.SelectSingleNode("Identity").text
                    arr(C, 2) = CDbl(nd.SelectSingleNode("SortOrder").text)
                    arr(C, 3) = nd.SelectSingleNode("Class").text
                    arr(C, 4) = nd.SelectSingleNode("Name").text
                    arr(C, 5) = CDbl(nd.SelectSingleNode("Percent").text)
                    arr(C, 6) = CDbl(nd.SelectSingleNode("Amount").text)
                    arr(C, 7) = "Upper"
                    Select Case nd.SelectSingleNode("Class").text
                        Case "NET"
                            arr(C, 8) = 1
                        Case "GROSS"
                            arr(C, 8) = 2
                        Case "TAX"
                            arr(C, 8) = 3
                        Case "BOND"
                            arr(C, 8) = 4
                        Case Else
                            arr(C, 8) = 5
                    End Select
                    arr(C, 9) = nd.SelectSingleNode("JobCostCategory").text
                    arr(C, 10) = nd.SelectSingleNode("JobCost").text
                    arr(C, 11) = JCName(nd.SelectSingleNode("JobCost").text)
                    If nd.SelectSingleNode("Lump").text > 0 Then
                        arr(C, 6) = CDbl(nd.SelectSingleNode("Lump").text)
                    End If
                End If
            End If
        End If
    Next nd
    On Error GoTo 0
    If C = 0 Then Exit Sub
    Set ows = Sheet0
    ows.Range("AS2").Resize(C, 11) = arr
    Set lObj = ows.ListObjects("tblTotals")
    With lObj
        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=Range("tblTotals[PrimarySort]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Sort.SortFields.Add Key:=Range("tblTotals[SortOrder]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
End Sub

Sub xmlVarTotalsTable() ' Build Markups(Addons) Table  **For Variance Report**
Dim arr As Variant
Dim C
    Set owb = ActiveWorkbook
    wbName = owb.Name

    XDocV.async = False
    XDocV.validateOnParse = False
    XDocV.Load xmlVarPath

    Set xNode = XDocV.SelectNodes("/Estimate/TotalsPageTable/TotalsPage")
    On Error Resume Next
    Set ows = Sheet0
    Set lObj = ows.ListObjects("tblTotals")
    For Each nd In xNode
        If nd.SelectSingleNode("Class").text <> "SUBTOTAL" Then
            If nd.SelectSingleNode("IsDeleted").text <> True Then
                If nd.SelectSingleNode("IsInactive").text <> True Then
                    Set C = lObj.ListColumns("Name").DataBodyRange.Find(nd.SelectSingleNode("Name").text, LookIn:=xlValues)
                    i = C.row - 1
                    If Not C Is Nothing Then
                        lObj.DataBodyRange(i, 13).Value = nd.SelectSingleNode("Amount").text
                    Else
                        Set lRow = lObj.ListRows.Add
                        lRow.Range(1, 1).Value = nd.SelectSingleNode("Identity").text
                        lRow.Range(1, 2).Value = CDbl(nd.SelectSingleNode("SortOrder").text)
                        lRow.Range(1, 3).Value = nd.SelectSingleNode("Class").text
                        lRow.Range(1, 4).Value = nd.SelectSingleNode("Name").text
                        lRow.Range(1, 5).Value = CDbl(nd.SelectSingleNode("Percent").text)
                        lRow.Range(1, 6).Value = 0
                        lRow.Range(1, 7).Value = "Upper"
                        Select Case nd.SelectSingleNode("Class").text
                            Case "NET"
                                lRow.Range(1, 8).Value = 1
                            Case "GROSS"
                                lRow.Range(1, 8).Value = 2
                            Case "TAX"
                                lRow.Range(1, 8).Value = 3
                            Case "BOND"
                                lRow.Range(1, 8).Value = 4
                            Case Else
                                lRow.Range(1, 8).Value = 5
                        End Select
                        lRow.Range(1, 9).Value = nd.SelectSingleNode("JobCostCategory").text
                        lRow.Range(1, 10).Value = nd.SelectSingleNode("JobCost").text
                        lRow.Range(1, 11).Value = JCName(nd.SelectSingleNode("JobCost").text)
                        lRow.Range(1, 13).Value = CDbl(nd.SelectSingleNode("Amount").text)
                        If nd.SelectSingleNode("Lump").text > 0 Then
                            lRow.Range(1, 13).Value = CDbl(nd.SelectSingleNode("Lump").text)
                        End If
                    End If
                End If
            End If
        End If
    Next nd
    On Error GoTo 0
'    If c = 0 Then Exit Sub
    With lObj
        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=Range("tblTotals[PrimarySort]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Sort.SortFields.Add Key:=Range("tblTotals[SortOrder]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
End Sub

Function JCName(nVal As String) As String
    Set xLvlSub = XDoc.SelectNodes("/Estimate/JobCostTable/JobCost[Code='" & nVal & "']")
    For Each nd0 In xLvlSub
        JCName = nd0.SelectSingleNode("Name").text
    Next nd0
End Function

Sub WBSDefinition()
Dim C
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    Set xNode = XDoc.SelectNodes("/Estimate/WbsTableNameTable/WbsTableName")
    
    Set ows = Sheet0
    Set lObj = ows.ListObjects("tblWBSMaster")
    With lObj
        For Each nd In xNode
        Set C = .ListColumns("Index").DataBodyRange.Find(nd.SelectSingleNode("Index").text, LookIn:=xlValues)
            If Not C Is Nothing Then
                i = C.row - 1
                .DataBodyRange(i, 2).Value = nd.SelectSingleNode("Name").text
            End If
        Next nd
    End With
    If bRefresh = False Then Call WbsHasData
End Sub

Sub WbsHasData() 'Checks WBS Tables for data
    Set ows = Sheet0
    Set lObj = ows.ListObjects("tblWBSMaster")
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    For i = 1 To lObj.DataBodyRange.Rows.count
        sXpath1 = lObj.DataBodyRange.Cells(i, 5).Value
        If XDoc.SelectNodes(sXpath1).Length > 1 Then
            lObj.DataBodyRange.Cells(i, 8).Value = True
        Else
            lObj.DataBodyRange.Cells(i, 8).Value = False
        End If
    Next i
End Sub

Sub xmlTotalsTableRefresh() ' Refresh Markups(Addons) Table  **Initial Setup**
Dim C
    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set xNode = XDoc.SelectNodes("/Estimate/TotalsPageTable/TotalsPage")
    Set ows = Sheet0
    Set lObj = ows.ListObjects("tblTotals")
    If lObj.Range.Rows.count <= 2 Then Exit Sub
    With lObj
        .ListColumns.Add(12).Name = "Updated"
        For Each nd In xNode
            Set C = .ListColumns("Identity").DataBodyRange.Find(nd.SelectSingleNode("Identity").text, LookIn:=xlValues)
            If Not C Is Nothing Then
                If nd.SelectSingleNode("IsDeleted").text <> True Then
                    If nd.SelectSingleNode("IsInactive").text <> True Then
                        i = C.row - 1
                        .DataBodyRange(i, 4).Value = nd.SelectSingleNode("Name").text
                        .DataBodyRange(i, 5).Value = CDbl(nd.SelectSingleNode("Percent").text)
                        .DataBodyRange(i, 6).Value = CDbl(nd.SelectSingleNode("Amount").text)
                        .DataBodyRange(i, 9).Value = nd.SelectSingleNode("JobCostCategory").text
                        .DataBodyRange(i, 10).Value = nd.SelectSingleNode("JobCost").text
                        .DataBodyRange(i, 11).Value = JCName(nd.SelectSingleNode("JobCost").text)
                        .DataBodyRange(i, 12).Value = "U"
                    Else
                        i = C.row - 1
                        .ListRows(i).Delete
                    End If
                Else
                    i = C.row - 1
                    .ListRows(i).Delete
                End If
            Else
                If nd.SelectSingleNode("Class").text <> "SUBTOTAL" Then
                    If nd.SelectSingleNode("IsDeleted").text <> True Then
                        If nd.SelectSingleNode("IsInactive").text <> True Then
                            Set lRow = lObj.ListRows.Add
                            lRow.Range(1, 1).Value = nd.SelectSingleNode("Identity").text
                            lRow.Range(1, 2).Value = CDbl(nd.SelectSingleNode("SortOrder").text)
                            lRow.Range(1, 3).Value = nd.SelectSingleNode("Class").text
                            lRow.Range(1, 4).Value = nd.SelectSingleNode("Name").text
                            lRow.Range(1, 5).Value = CDbl(nd.SelectSingleNode("Percent").text)
                            lRow.Range(1, 6).Value = CDbl(nd.SelectSingleNode("Amount").text)
                            lRow.Range(1, 7).Value = "Lower"
                            Select Case nd.SelectSingleNode("Class").text
                                Case "NET"
                                    lRow.Range(1, 8).Value = 1
                                Case "GROSS"
                                    lRow.Range(1, 8).Value = 2
                                Case "TAX"
                                    lRow.Range(1, 8).Value = 3
                                Case "BOND"
                                    lRow.Range(1, 8).Value = 4
                                Case Else
                                    lRow.Range(1, 8).Value = 5
                            End Select
                            lRow.Range(1, 9).Value = nd.SelectSingleNode("JobCostCategory").text
                            lRow.Range(1, 10).Value = nd.SelectSingleNode("JobCost").text
                            lRow.Range(1, 11).Value = JCName(nd.SelectSingleNode("JobCost").text)
                            lRow.Range(1, 12).Value = "N"
                        End If
                    End If
                End If
            End If
        Next nd
    End With
'**Check for deleted items
    With lObj
        For i = .ListRows.count To 1 Step -1
            If .DataBodyRange(i, 12).Value = "" Then
                .ListRows(i).Delete
            End If
        Next i
        .ListColumns(12).Delete
    End With
'**Primary Sort
    With lObj
        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=Range("tblTotals[PrimarySort]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Sort.SortFields.Add Key:=Range("tblTotals[SortOrder]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
End Sub

Sub UpdateAlternates()
       Dim oXmlDoc As DOMDocument60
       Dim oXmlNode As IXMLDOMNode
    
       Set oXmlDoc = New DOMDocument60
       oXmlDoc.async = False
       oXmlDoc.Load xmlPath

        Set xNode = oXmlDoc.SelectNodes("/Estimate/AlternateTable/Alternate")
        
        For Each nd In xNode
            Select Case nd.SelectSingleNode("Status").text
                Case 0
                    nd.SelectSingleNode("Status").text = "Pending"
                Case 1
                    nd.SelectSingleNode("Status").text = "Approved"
                Case 2
                    nd.SelectSingleNode("Status").text = "Denied"
                Case Else
             End Select
        Next
        oXmlDoc.Save xmlPath
End Sub

Sub WBSTest()
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    Dim sSearch As String
    sSearch = "_EquipOper(med)"
    Set xNode = XDoc.SelectNodes("/Estimate/LaborTable/Labor[Trade= '" & sSearch & "']")
        For Each nd In xNode
               Debug.Print nd.SelectSingleNode("Rate").text
        Next nd
 End Sub
