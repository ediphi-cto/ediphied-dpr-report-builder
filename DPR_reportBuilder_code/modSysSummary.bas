Attribute VB_Name = "modSysSummary"
Option Explicit

Sub SummaryDetail()
Dim rSys As Range
Dim sFnd As String
Dim C, firstAddress
    Set owb = ActiveWorkbook
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set xNode = XDoc.SelectNodes(sXpath) 'Item Level
    sJobUM = Range("rngJobUnitName").Value
    Set lObj = Sheet0.ListObjects("tblRptTrack")
    sXpath1 = lObj.DataBodyRange(1, 10).Value
    bCkb1 = lObj.DataBodyRange(1, 11).Value
    sLvl1xNd = lObj.DataBodyRange(1, 12).Value
    sLvl1Code = lObj.DataBodyRange(1, 13).Value
    
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
'Read data from Items table
    For Each nd In xNode
        Set xLvl1 = XDoc.SelectNodes(sXpath1 & "[Index=" & nd.SelectSingleNode(sLvl1xNd).text & "]")
        For Each nd1 In xLvl1
            If bCkb1 = True Then
                If sLvl1xNd = "Division" Then
                    sLvl1 = Left(nd1.SelectSingleNode(sLvl1Code).text, 2) & "-" & nd1.SelectSingleNode("Name").text
                Else
                    sLvl1 = nd1.SelectSingleNode(sLvl1Code).text & "-" & nd1.SelectSingleNode("Name").text
                End If
            Else
                sLvl1 = nd1.SelectSingleNode("Name").text
            End If
            sCode1 = nd1.SelectSingleNode(sLvl1Code).text
        Next nd1
        
'Add to data collection
        oTxt = sLvl1 & sCode1
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, Val(CDbl(nd.SelectSingleNode("GrandTotal").text)))
        Else
            q = Dict.Item(oTxt)
            q(2) = q(2) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
    Next nd
'Load data collection to sheet
    Set ows = Sheet3
    ows.Activate
'    ows.Range(ows.Range("SysStart").Offset(1, 0), ows.Range("SysEnd").Offset(-2, 0)).EntireRow.ClearContents
    If ows.Range(ows.Range("SysStart").Offset(1, 0), ows.Range("SysEnd").Offset(-2, 0)).Rows.count > 2 Then
        ows.Range(ows.Range("SysStart").Offset(1, 0), ows.Range("SysEnd").Offset(-2, 0)).EntireRow.Delete
    End If
    
    ReDim dataArray(1 To Dict.count, 1 To 3)
    On Error Resume Next
    C = 0
    For Each k In Dict.Keys
        C = C + 1
        ows.Range("sysEnd").Offset(-1, 0).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
        X = ows.Range("sysEnd").Offset(-2, 1).row
        If InStr(1, sLvl1Code, "Name") = 0 Then
            ows.Cells(X, 2).Value = Dict.Item(k)(0)
        End If
        ows.Cells(X, 3).Value = Dict.Item(k)(1)
        ows.Cells(X, 7).Formula = "=IFERROR(RC[1]/rngJobSize,0)"
        ows.Cells(X, 8).Value = Dict.Item(k)(2)
    Next k
    On Error GoTo 0
'Sort Summary
    ows.Sort.SortFields.Clear
    ows.Sort.SortFields.Add Key:=Range(ows.Range("SysStart").Offset(1, 0), ows.Range("SysEnd").Offset(-2, 0)), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ows.Sort.SortFields.Add Key:=Range(ows.Range("SysStart").Offset(1, 1), ows.Range("SysEnd").Offset(-2, 1)), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ows.Sort
        .SetRange Range(ows.Range("SysStart").Offset(1, 0), ows.Range("SysEnd").Offset(-2, 6))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
   
'Set Total formula
    r = Range(Range("SysStart"), Range("SysEnd")).count
    ows.Range("SysEnd").Offset(0, 6).FormulaR1C1 = "=SUM(R[-" & r - 1 & "]C:R[-1]C)"
    ows.Range("SysEnd").Offset(0, 6).NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    Call Addons
    Exit Sub
errHndlr:
    Exit Sub
End Sub



