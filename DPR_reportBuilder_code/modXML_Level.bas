Attribute VB_Name = "modXML_Level"
Public sCode0, sCode1, sCode2, sCode3, sCode4, sCode5 As String
Public dVal0 As Double
Sub xmlLevel1() 'Build 1 level array
    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
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
    'Check Box Subcontractor
        If bCkbSub = True Then
            Set xLvlSub = XDoc.SelectNodes(sXpathSub & "[Index=" & nd.SelectSingleNode("Subcontractor").text & "]")
            For Each ndSub In xLvlSub
                sLvlSub = "(" & ndSub.SelectSingleNode("Name").text & ") "
            Next ndSub
        End If
        
    'Check Box Notes
        If bCkbAll = True Then
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl1xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl1xNd).text)
        End If
    'Add to data collection
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                nd.SelectSingleNode("Description").text, _
                                sLvlSub & nd.SelectSingleNode("ItemNote").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                nd.SelectSingleNode("TakeoffUnit").text, _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)))
        Else
            q = Dict.Item(oTxt)
            q(6) = q(6) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(8) = q(8) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
    Next nd
    If Dict.count > 0 Then
        ReDim dataArray(1 To Dict.count, 1 To 10)
        On Error Resume Next
        C = 0
        For Each k In Dict.Keys
            C = C + 1
            dataArray(C, 1) = Dict.Item(k)(0)
            dataArray(C, 2) = Dict.Item(k)(1)
            dataArray(C, 3) = Dict.Item(k)(2)
            dataArray(C, 4) = Dict.Item(k)(3)
            dataArray(C, 5) = Dict.Item(k)(4)
            dataArray(C, 6) = Dict.Item(k)(5)
            dataArray(C, 7) = Dict.Item(k)(6)
            dataArray(C, 8) = Dict.Item(k)(7)
            dataArray(C, 9) = Dict.Item(k)(8) / Dict.Item(k)(6)
            dataArray(C, 10) = Dict.Item(k)(8)
        Next k
        On Error GoTo 0
        Set rsNew = ADOCopyArrayIntoRecordset(argArray:=dataArray)
        If bRefresh = False Then
            Call Create_PivotTable_ODBC_MO
        End If
    Else
        MsgBox "Unable to create this report. Make sure the WBS codes you have selected have values associated.", vbCritical, "Unable To Create Report"
    End If
End Sub

Sub xmlLevel2() 'Build 2 level array

    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
    'Set Level One
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
    'Set Level 2
        Set xLvl2 = XDoc.SelectNodes(sXpath2 & "[Index=" & nd.SelectSingleNode(sLvl2xNd).text & "]")
        For Each nd2 In xLvl2
            If bCkb2 = True Then
                If sLvl2xNd = "Division" Then
                    sLvl2 = Left(nd2.SelectSingleNode(sLvl2Code).text, 2) & "-" & nd2.SelectSingleNode("Name").text
                Else
                    sLvl2 = nd2.SelectSingleNode(sLvl2Code).text & "-" & nd2.SelectSingleNode("Name").text
                End If
            Else
                sLvl2 = nd2.SelectSingleNode("Name").text
            End If
            sCode2 = nd2.SelectSingleNode(sLvl2Code).text
        Next nd2
    'Check Box Subcontractor
        If bCkbSub = True Then
            Set xLvlSub = XDoc.SelectNodes(sXpathSub & "[Index=" & nd.SelectSingleNode("Subcontractor").text & "]")
            For Each ndSub In xLvlSub
                sLvlSub = "(" & ndSub.SelectSingleNode("Name").text & ") "
            Next ndSub
        End If
    'Check Box Notes
        If bCkbAll = True Then
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl2xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) & CInt(nd.SelectSingleNode(sLvl2xNd).text)
        End If
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                nd.SelectSingleNode("Description").text, _
                                sLvlSub & nd.SelectSingleNode("ItemNote").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                nd.SelectSingleNode("TakeoffUnit").text, _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)))
        Else
            q = Dict.Item(oTxt)
            q(8) = q(8) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(10) = q(10) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
    Next nd
    If Dict.count > 0 Then
        ReDim dataArray(1 To Dict.count, 1 To 12)
        On Error Resume Next
        C = 0
        For Each k In Dict.Keys
            C = C + 1
            dataArray(C, 1) = Dict.Item(k)(0)
            dataArray(C, 2) = Dict.Item(k)(1)
            dataArray(C, 3) = Dict.Item(k)(2)
            dataArray(C, 4) = Dict.Item(k)(3)
            dataArray(C, 5) = Dict.Item(k)(4)
            dataArray(C, 6) = Dict.Item(k)(5)
            dataArray(C, 7) = Dict.Item(k)(6)
            dataArray(C, 8) = Dict.Item(k)(7)
            dataArray(C, 9) = Dict.Item(k)(8)
            dataArray(C, 10) = Dict.Item(k)(9)
            dataArray(C, 11) = Dict.Item(k)(10) / Dict.Item(k)(8)
            dataArray(C, 12) = Dict.Item(k)(10)
        Next k
        On Error GoTo 0
        Set rsNew = ADOCopyArrayIntoRecordset(argArray:=dataArray)
        If bRefresh = False Then
            Call Create_PivotTable_ODBC_MO
        End If
    Else
        MsgBox "Unable to create this report. Make sure the WBS codes you have selected have values associated.", vbCritical, "Unable To Create Report"
    End If
End Sub

Sub xmlLevel3() 'Build 3 level array

    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
    'Set Level One
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
    'Set Level 2
        Set xLvl2 = XDoc.SelectNodes(sXpath2 & "[Index=" & nd.SelectSingleNode(sLvl2xNd).text & "]")
        For Each nd2 In xLvl2
            If bCkb2 = True Then
                If sLvl2xNd = "Division" Then
                    sLvl2 = Left(nd2.SelectSingleNode(sLvl2Code).text, 2) & "-" & nd2.SelectSingleNode("Name").text
                Else
                    sLvl2 = nd2.SelectSingleNode(sLvl2Code).text & "-" & nd2.SelectSingleNode("Name").text
                End If
            Else
                sLvl2 = nd2.SelectSingleNode("Name").text
            End If
            sCode2 = nd2.SelectSingleNode(sLvl2Code).text
        Next nd2
    'Set Level 3
        Set xLvl3 = XDoc.SelectNodes(sXpath3 & "[Index=" & nd.SelectSingleNode(sLvl3xNd).text & "]")
        For Each nd3 In xLvl3
            If bCkb3 = True Then
                If sLvl3xNd = "Division" Then
                    sLvl3 = Left(nd3.SelectSingleNode(sLvl3Code).text, 2) & "-" & nd3.SelectSingleNode("Name").text
                Else
                    sLvl3 = nd3.SelectSingleNode(sLvl3Code).text & "-" & nd3.SelectSingleNode("Name").text
                End If
            Else
                sLvl3 = nd3.SelectSingleNode("Name").text
            End If
            sCode3 = nd3.SelectSingleNode(sLvl3Code).text
        Next nd3
    'Check Box Subcontractor
        If bCkbSub = True Then
            Set xLvlSub = XDoc.SelectNodes(sXpathSub & "[Index=" & nd.SelectSingleNode("Subcontractor").text & "]")
            For Each ndSub In xLvlSub
                sLvlSub = "(" & ndSub.SelectSingleNode("Name").text & ") "
            Next ndSub
        End If
    'Check Box Notes
        If bCkbAll = True Then
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl2xNd).text) & CInt(nd.SelectSingleNode(sLvl3xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) & CInt(nd.SelectSingleNode(sLvl2xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl3xNd).text)
        End If
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                sCode3, sLvl3, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                nd.SelectSingleNode("Description").text, _
                                sLvlSub & nd.SelectSingleNode("ItemNote").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                nd.SelectSingleNode("TakeoffUnit").text, _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)))
        Else
            q = Dict.Item(oTxt)
            q(10) = q(10) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(12) = q(12) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
    Next nd
    If Dict.count > 0 Then
        ReDim dataArray(1 To Dict.count, 1 To 14)
        On Error Resume Next
        C = 0
        For Each k In Dict.Keys
            C = C + 1
            dataArray(C, 1) = Dict.Item(k)(0)
            dataArray(C, 2) = Dict.Item(k)(1)
            dataArray(C, 3) = Dict.Item(k)(2)
            dataArray(C, 4) = Dict.Item(k)(3)
            dataArray(C, 5) = Dict.Item(k)(4)
            dataArray(C, 6) = Dict.Item(k)(5)
            dataArray(C, 7) = Dict.Item(k)(6)
            dataArray(C, 8) = Dict.Item(k)(7)
            dataArray(C, 9) = Dict.Item(k)(8)
            dataArray(C, 10) = Dict.Item(k)(9)
            dataArray(C, 11) = Dict.Item(k)(10)
            dataArray(C, 12) = Dict.Item(k)(11)
            dataArray(C, 13) = Dict.Item(k)(12) / Dict.Item(k)(10)
            dataArray(C, 14) = Dict.Item(k)(12)
        Next k
        On Error GoTo 0
        Set rsNew = ADOCopyArrayIntoRecordset(argArray:=dataArray)
        If bRefresh = False Then
            Call Create_PivotTable_ODBC_MO
        End If
    Else
        MsgBox "Unable to create this report. Make sure the WBS codes you have selected have values associated.", vbCritical, "Unable To Create Report"
    End If
End Sub

Sub xmlLevel4() 'Build 4 level array
    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
    'Set Level One
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
    'Set Level 2
        Set xLvl2 = XDoc.SelectNodes(sXpath2 & "[Index=" & nd.SelectSingleNode(sLvl2xNd).text & "]")
        For Each nd2 In xLvl2
            If bCkb2 = True Then
                If sLvl2xNd = "Division" Then
                    sLvl2 = Left(nd2.SelectSingleNode(sLvl2Code).text, 2) & "-" & nd2.SelectSingleNode("Name").text
                Else
                    sLvl2 = nd2.SelectSingleNode(sLvl2Code).text & "-" & nd2.SelectSingleNode("Name").text
                End If
            Else
                sLvl2 = nd2.SelectSingleNode("Name").text
            End If
            sCode2 = nd2.SelectSingleNode(sLvl2Code).text
        Next nd2
    'Set Level 3
        Set xLvl3 = XDoc.SelectNodes(sXpath3 & "[Index=" & nd.SelectSingleNode(sLvl3xNd).text & "]")
        For Each nd3 In xLvl3
            If bCkb3 = True Then
                If sLvl3xNd = "Division" Then
                    sLvl3 = Left(nd3.SelectSingleNode(sLvl3Code).text, 2) & "-" & nd3.SelectSingleNode("Name").text
                Else
                    sLvl3 = nd3.SelectSingleNode(sLvl3Code).text & "-" & nd3.SelectSingleNode("Name").text
                End If
            Else
                sLvl3 = nd3.SelectSingleNode("Name").text
            End If
            sCode3 = nd3.SelectSingleNode(sLvl3Code).text
        Next nd3
    'Set Level 4
        Set xLvl4 = XDoc.SelectNodes(sXpath4 & "[Index=" & nd.SelectSingleNode(sLvl4xNd).text & "]")
        For Each nd4 In xLvl4
            If bCkb4 = True Then
                If sLvl4xNd = "Division" Then
                    sLvl4 = Left(nd4.SelectSingleNode(sLvl4Code).text, 2) & "-" & nd4.SelectSingleNode("Name").text
                Else
                    sLvl4 = nd4.SelectSingleNode(sLvl4Code).text & "-" & nd4.SelectSingleNode("Name").text
                End If
            Else
                sLvl4 = nd4.SelectSingleNode("Name").text
            End If
            sCode4 = nd4.SelectSingleNode(sLvl4Code).text
        Next nd4
    'Check Box Subcontractor
        If bCkbSub = True Then
            Set xLvlSub = XDoc.SelectNodes(sXpathSub & "[Index=" & nd.SelectSingleNode("Subcontractor").text & "]")
            For Each ndSub In xLvlSub
                sLvlSub = "(" & ndSub.SelectSingleNode("Name").text & ") "
            Next ndSub
        End If
    'Check Box Notes
        If bCkbAll = True Then
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl2xNd).text) & CInt(nd.SelectSingleNode(sLvl3xNd).text) & CInt(nd.SelectSingleNode(sLvl4xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) & CInt(nd.SelectSingleNode(sLvl2xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl3xNd).text) & CInt(nd.SelectSingleNode(sLvl4xNd).text)
        End If
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                sCode3, sLvl3, _
                                sCode4, sLvl4, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                nd.SelectSingleNode("Description").text, _
                                sLvlSub & nd.SelectSingleNode("ItemNote").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                nd.SelectSingleNode("TakeoffUnit").text, _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)))
        Else
            q = Dict.Item(oTxt)
            q(12) = q(12) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(14) = q(14) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
    Next nd
    If Dict.count > 0 Then
        ReDim dataArray(1 To Dict.count, 1 To 16)
        On Error Resume Next
        C = 0
        For Each k In Dict.Keys
            C = C + 1
            dataArray(C, 1) = Dict.Item(k)(0)
            dataArray(C, 2) = Dict.Item(k)(1)
            dataArray(C, 3) = Dict.Item(k)(2)
            dataArray(C, 4) = Dict.Item(k)(3)
            dataArray(C, 5) = Dict.Item(k)(4)
            dataArray(C, 6) = Dict.Item(k)(5)
            dataArray(C, 7) = Dict.Item(k)(6)
            dataArray(C, 8) = Dict.Item(k)(7)
            dataArray(C, 9) = Dict.Item(k)(8)
            dataArray(C, 10) = Dict.Item(k)(9)
            dataArray(C, 11) = Dict.Item(k)(10)
            dataArray(C, 12) = Dict.Item(k)(11)
            dataArray(C, 13) = Dict.Item(k)(12)
            dataArray(C, 14) = Dict.Item(k)(13)
            dataArray(C, 15) = Dict.Item(k)(14) / Dict.Item(k)(12)
            dataArray(C, 16) = Dict.Item(k)(14)
        Next k
        On Error GoTo 0
        Set rsNew = ADOCopyArrayIntoRecordset(argArray:=dataArray)
        If bRefresh = False Then
            Call Create_PivotTable_ODBC_MO
        End If
    Else
        MsgBox "Unable to create this report. Make sure the WBS codes you have selected have values associated.", vbCritical, "Unable To Create Report"
    End If
End Sub

Sub xmlLevel5() 'Build 5 level array

    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
    'Set Level One
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
    'Set Level 2
        Set xLvl2 = XDoc.SelectNodes(sXpath2 & "[Index=" & nd.SelectSingleNode(sLvl2xNd).text & "]")
        For Each nd2 In xLvl2
            If bCkb2 = True Then
                If sLvl2xNd = "Division" Then
                    sLvl2 = Left(nd2.SelectSingleNode(sLvl2Code).text, 2) & "-" & nd2.SelectSingleNode("Name").text
                Else
                    sLvl2 = nd2.SelectSingleNode(sLvl2Code).text & "-" & nd2.SelectSingleNode("Name").text
                End If
            Else
                sLvl2 = nd2.SelectSingleNode("Name").text
            End If
            sCode2 = nd2.SelectSingleNode(sLvl2Code).text
        Next nd2
    'Set Level 3
        Set xLvl3 = XDoc.SelectNodes(sXpath3 & "[Index=" & nd.SelectSingleNode(sLvl3xNd).text & "]")
        For Each nd3 In xLvl3
            If bCkb3 = True Then
                If sLvl3xNd = "Division" Then
                    sLvl3 = Left(nd3.SelectSingleNode(sLvl3Code).text, 2) & "-" & nd3.SelectSingleNode("Name").text
                Else
                    sLvl3 = nd3.SelectSingleNode(sLvl3Code).text & "-" & nd3.SelectSingleNode("Name").text
                End If
            Else
                sLvl3 = nd3.SelectSingleNode("Name").text
            End If
            sCode3 = nd3.SelectSingleNode(sLvl3Code).text
        Next nd3
    'Set Level 4
        Set xLvl4 = XDoc.SelectNodes(sXpath4 & "[Index=" & nd.SelectSingleNode(sLvl4xNd).text & "]")
        For Each nd4 In xLvl4
            If bCkb4 = True Then
                If sLvl4xNd = "Division" Then
                    sLvl4 = Left(nd4.SelectSingleNode(sLvl4Code).text, 2) & "-" & nd4.SelectSingleNode("Name").text
                Else
                    sLvl4 = nd4.SelectSingleNode(sLvl4Code).text & "-" & nd4.SelectSingleNode("Name").text
                End If
            Else
                sLvl4 = nd4.SelectSingleNode("Name").text
            End If
            sCode4 = nd4.SelectSingleNode(sLvl4Code).text
        Next nd4
    'Set Level 5
        Set xLvl5 = XDoc.SelectNodes(sXpath5 & "[Index=" & nd.SelectSingleNode(sLvl5xNd).text & "]")
        For Each nd5 In xLvl5
            If bCkb5 = True Then
                If sLvl5xNd = "Division" Then
                    sLvl5 = Left(nd5.SelectSingleNode(sLvl5Code).text, 2) & "-" & nd5.SelectSingleNode("Name").text
                Else
                    sLvl5 = nd5.SelectSingleNode(sLvl5Code).text & "-" & nd5.SelectSingleNode("Name").text
                End If
            Else
                sLvl5 = nd5.SelectSingleNode("Name").text
            End If
            sCode5 = nd5.SelectSingleNode(sLvl5Code).text
        Next nd5
    'Check Box Subcontractor
        If bCkbSub = True Then
            Set xLvlSub = XDoc.SelectNodes(sXpathSub & "[Index=" & nd.SelectSingleNode("Subcontractor").text & "]")
            For Each ndSub In xLvlSub
                sLvlSub = "(" & ndSub.SelectSingleNode("Name").text & ") "
            Next ndSub
        End If
    'Check Box Notes
        If bCkbAll = True Then
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl2xNd).text) & CInt(nd.SelectSingleNode(sLvl3xNd).text) & CInt(nd.SelectSingleNode(sLvl4xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl5xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) & CInt(nd.SelectSingleNode(sLvl2xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl3xNd).text) & CInt(nd.SelectSingleNode(sLvl4xNd).text) & CInt(nd.SelectSingleNode(sLvl5xNd).text)
        End If
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                sCode3, sLvl3, _
                                sCode4, sLvl4, _
                                sCode5, sLvl5, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                nd.SelectSingleNode("Description").text, _
                                sLvlSub & nd.SelectSingleNode("ItemNote").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                nd.SelectSingleNode("TakeoffUnit").text, _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)))
        Else
            q = Dict.Item(oTxt)
            q(14) = q(14) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(16) = q(16) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
    Next nd
    If Dict.count > 0 Then
        ReDim dataArray(1 To Dict.count, 1 To 18)
        On Error Resume Next
        C = 0
        For Each k In Dict.Keys
            C = C + 1
            dataArray(C, 1) = Dict.Item(k)(0)
            dataArray(C, 2) = Dict.Item(k)(1)
            dataArray(C, 3) = Dict.Item(k)(2)
            dataArray(C, 4) = Dict.Item(k)(3)
            dataArray(C, 5) = Dict.Item(k)(4)
            dataArray(C, 6) = Dict.Item(k)(5)
            dataArray(C, 7) = Dict.Item(k)(6)
            dataArray(C, 8) = Dict.Item(k)(7)
            dataArray(C, 9) = Dict.Item(k)(8)
            dataArray(C, 10) = Dict.Item(k)(9)
            dataArray(C, 11) = Dict.Item(k)(10)
            dataArray(C, 12) = Dict.Item(k)(11)
            dataArray(C, 13) = Dict.Item(k)(12)
            dataArray(C, 14) = Dict.Item(k)(13)
            dataArray(C, 15) = Dict.Item(k)(14)
            dataArray(C, 16) = Dict.Item(k)(15)
            dataArray(C, 17) = Dict.Item(k)(16) / Dict.Item(k)(14)
            dataArray(C, 18) = Dict.Item(k)(16)
        Next k
        On Error GoTo 0
        Set rsNew = ADOCopyArrayIntoRecordset(argArray:=dataArray)
        If bRefresh = False Then
            Call Create_PivotTable_ODBC_MO
        End If
    Else
        MsgBox "Unable to create this report. Make sure the WBS codes you have selected have values associated.", vbCritical, "Unable To Create Report"
    End If
End Sub

Private Function ADOCopyArrayIntoRecordset(argArray As Variant) As ADODB.RecordSet 'Create data recordset for pivot cache
Dim rsADO As ADODB.RecordSet
Dim lngR As Long
Dim lngC As Long

    Set rsADO = New ADODB.RecordSet
    For i = 1 To iLvl
        Select Case i
            Case 1
                sLvl1Code = sLvl1Code & i
                rsADO.Fields.Append sLvl1Code, adVariant
                rsADO.Fields.Append sLvl1Item, adVariant
            Case 2
                sLvl2Code = sLvl2Code & i
                rsADO.Fields.Append sLvl2Code, adVariant
                rsADO.Fields.Append sLvl2Item, adVariant
            Case 3
                sLvl3Code = sLvl3Code & i
                rsADO.Fields.Append sLvl3Code, adVariant
                rsADO.Fields.Append sLvl3Item, adVariant
            Case 4
                sLvl4Code = sLvl4Code & i
                rsADO.Fields.Append sLvl4Code, adVariant
                rsADO.Fields.Append sLvl4Item, adVariant
            Case 5
                sLvl5Code = sLvl5Code & i
                rsADO.Fields.Append sLvl5Code, adVariant
                rsADO.Fields.Append sLvl5Item, adVariant
        End Select
    Next i
    rsADO.Fields.Append "SortOrder", adVariant
    rsADO.Fields.Append "ItemCode", adVariant
    rsADO.Fields.Append "Description", adVariant
    rsADO.Fields.Append "ItemNote", adVariant
    rsADO.Fields.Append "TakeoffQty", adVariant
    rsADO.Fields.Append "TakeoffUnit", adVariant
    rsADO.Fields.Append "UnitPrice", adVariant
    rsADO.Fields.Append "GrandTotal", adVariant
    rsADO.Open

    For lngR = 1 To UBound(argArray, 1)
    rsADO.AddNew
       For lngC = 1 To UBound(argArray, 2)
       rsADO.Fields(lngC - 1).Value = argArray(lngR, lngC)
       Next lngC
       rsADO.MoveNext
       lngC = 0
    Next lngR
    rsADO.MoveFirst
    Set ADOCopyArrayIntoRecordset = rsADO
    Set rsADO = Nothing
End Function

'*********************************************************
'Special FB code added for WBS 14 3-level sort. 05-19-2020
'*********************************************************

Sub xmlLevel3FB() 'Build 3 level array with WBS 14 FB
    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
    'Set Level One thru 3
    If nd.SelectSingleNode(sLvl1xNd).text = 0 Then GoTo errNextNd
    Level_FB (nd.SelectSingleNode(sLvl1xNd).text)
    'Check Box Subcontractor
        If bCkbSub = True Then
            Set xLvlSub = XDoc.SelectNodes(sXpathSub & "[Index=" & nd.SelectSingleNode("Subcontractor").text & "]")
            For Each ndSub In xLvlSub
                sLvlSub = "(" & ndSub.SelectSingleNode("Name").text & ") "
            Next ndSub
        End If
    'Check Box Notes
        If bCkbAll = True Then
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl1xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl1xNd).text)
        End If
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                sCode3, sLvl3, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                nd.SelectSingleNode("Description").text, _
                                sLvlSub & nd.SelectSingleNode("ItemNote").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                nd.SelectSingleNode("TakeoffUnit").text, _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)))
        Else
            q = Dict.Item(oTxt)
            q(10) = q(10) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(12) = q(12) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
errNextNd:
    Next nd
    If Dict.count > 0 Then
        ReDim dataArray(1 To Dict.count, 1 To 14)
        On Error Resume Next
        C = 0
        For Each k In Dict.Keys
            C = C + 1
            dataArray(C, 1) = Dict.Item(k)(0)
            dataArray(C, 2) = Dict.Item(k)(1)
            dataArray(C, 3) = Dict.Item(k)(2)
            dataArray(C, 4) = Dict.Item(k)(3)
            dataArray(C, 5) = Dict.Item(k)(4)
            dataArray(C, 6) = Dict.Item(k)(5)
            dataArray(C, 7) = Dict.Item(k)(6)
            dataArray(C, 8) = Dict.Item(k)(7)
            dataArray(C, 9) = Dict.Item(k)(8)
            dataArray(C, 10) = Dict.Item(k)(9)
            dataArray(C, 11) = Dict.Item(k)(10)
            dataArray(C, 12) = Dict.Item(k)(11)
            dataArray(C, 13) = Dict.Item(k)(12) / Dict.Item(k)(10)
            dataArray(C, 14) = Dict.Item(k)(12)
        Next k
        On Error GoTo 0
        Set rsNew = ADOCopyArrayIntoRecordset(argArray:=dataArray)
        If bRefresh = False Then
            Call Create_PivotTable_ODBC_MO
        End If
    Else
        MsgBox "Unable to create this report. Make sure the WBS codes you have selected have values associated.", vbCritical, "Unable To Create Report"
    End If
End Sub

Sub xmlLevel4FB() 'Build 4 level array with WBS 14
    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
    'Set Level One
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
    'Set Level 2 thru 4
        If nd.SelectSingleNode(sLvl2xNd).text = 0 Then GoTo errNextNd
        Level_FB2 (nd.SelectSingleNode(sLvl2xNd).text)
    'Check Box Subcontractor
        If bCkbSub = True Then
            Set xLvlSub = XDoc.SelectNodes(sXpathSub & "[Index=" & nd.SelectSingleNode("Subcontractor").text & "]")
            For Each ndSub In xLvlSub
                sLvlSub = "(" & ndSub.SelectSingleNode("Name").text & ") "
            Next ndSub
        End If
    'Check Box Notes
        If bCkbAll = True Then
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl2xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) & CInt(nd.SelectSingleNode(sLvl2xNd).text)
        End If
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                sCode3, sLvl3, _
                                sCode4, sLvl4, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                nd.SelectSingleNode("Description").text, _
                                sLvlSub & nd.SelectSingleNode("ItemNote").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                nd.SelectSingleNode("TakeoffUnit").text, _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)))
        Else
            q = Dict.Item(oTxt)
            q(12) = q(12) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(14) = q(14) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
errNextNd:
    Next nd
    If Dict.count > 0 Then
        ReDim dataArray(1 To Dict.count, 1 To 16)
        On Error Resume Next
        C = 0
        For Each k In Dict.Keys
            C = C + 1
            dataArray(C, 1) = Dict.Item(k)(0)
            dataArray(C, 2) = Dict.Item(k)(1)
            dataArray(C, 3) = Dict.Item(k)(2)
            dataArray(C, 4) = Dict.Item(k)(3)
            dataArray(C, 5) = Dict.Item(k)(4)
            dataArray(C, 6) = Dict.Item(k)(5)
            dataArray(C, 7) = Dict.Item(k)(6)
            dataArray(C, 8) = Dict.Item(k)(7)
            dataArray(C, 9) = Dict.Item(k)(8)
            dataArray(C, 10) = Dict.Item(k)(9)
            dataArray(C, 11) = Dict.Item(k)(10)
            dataArray(C, 12) = Dict.Item(k)(11)
            dataArray(C, 13) = Dict.Item(k)(12)
            dataArray(C, 14) = Dict.Item(k)(13)
            dataArray(C, 15) = Dict.Item(k)(14) / Dict.Item(k)(12)
            dataArray(C, 16) = Dict.Item(k)(14)
        Next k
        On Error GoTo 0
        Set rsNew = ADOCopyArrayIntoRecordset(argArray:=dataArray)
        If bRefresh = False Then
            Call Create_PivotTable_ODBC_MO
        End If
    Else
        MsgBox "Unable to create this report. Make sure the WBS codes you have selected have values associated.", vbCritical, "Unable To Create Report"
    End If
End Sub

Sub xmlLevel5FB() 'Build 5 level array with WBS 14

    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
    'Set Level One
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
    'Set Level 2
        Set xLvl2 = XDoc.SelectNodes(sXpath2 & "[Index=" & nd.SelectSingleNode(sLvl2xNd).text & "]")
        For Each nd2 In xLvl2
            If bCkb2 = True Then
                If sLvl2xNd = "Division" Then
                    sLvl2 = Left(nd2.SelectSingleNode(sLvl2Code).text, 2) & "-" & nd2.SelectSingleNode("Name").text
                Else
                    sLvl2 = nd2.SelectSingleNode(sLvl2Code).text & "-" & nd2.SelectSingleNode("Name").text
                End If
            Else
                sLvl2 = nd2.SelectSingleNode("Name").text
            End If
            sCode2 = nd2.SelectSingleNode(sLvl2Code).text
        Next nd2
    'Set Level 3 thru 5
        If nd.SelectSingleNode(sLvl3xNd).text = 0 Then GoTo errNextNd
        Level_FB3 (nd.SelectSingleNode(sLvl3xNd).text)
    'Check Box Subcontractor
        If bCkbSub = True Then
            Set xLvlSub = XDoc.SelectNodes(sXpathSub & "[Index=" & nd.SelectSingleNode("Subcontractor").text & "]")
            For Each ndSub In xLvlSub
                sLvlSub = "(" & ndSub.SelectSingleNode("Name").text & ") "
            Next ndSub
        End If
    'Check Box Notes
        If bCkbAll = True Then
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl2xNd).text) & CInt(nd.SelectSingleNode(sLvl3xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) & CInt(nd.SelectSingleNode(sLvl2xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl3xNd).text)
        End If
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                sCode3, sLvl3, _
                                sCode4, sLvl4, _
                                sCode5, sLvl5, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                nd.SelectSingleNode("Description").text, _
                                sLvlSub & nd.SelectSingleNode("ItemNote").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                nd.SelectSingleNode("TakeoffUnit").text, _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)))
        Else
            q = Dict.Item(oTxt)
            q(14) = q(14) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(16) = q(16) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
errNextNd:
    Next nd
    If Dict.count > 0 Then
        ReDim dataArray(1 To Dict.count, 1 To 18)
        On Error Resume Next
        C = 0
        For Each k In Dict.Keys
            C = C + 1
            dataArray(C, 1) = Dict.Item(k)(0)
            dataArray(C, 2) = Dict.Item(k)(1)
            dataArray(C, 3) = Dict.Item(k)(2)
            dataArray(C, 4) = Dict.Item(k)(3)
            dataArray(C, 5) = Dict.Item(k)(4)
            dataArray(C, 6) = Dict.Item(k)(5)
            dataArray(C, 7) = Dict.Item(k)(6)
            dataArray(C, 8) = Dict.Item(k)(7)
            dataArray(C, 9) = Dict.Item(k)(8)
            dataArray(C, 10) = Dict.Item(k)(9)
            dataArray(C, 11) = Dict.Item(k)(10)
            dataArray(C, 12) = Dict.Item(k)(11)
            dataArray(C, 13) = Dict.Item(k)(12)
            dataArray(C, 14) = Dict.Item(k)(13)
            dataArray(C, 15) = Dict.Item(k)(14)
            dataArray(C, 16) = Dict.Item(k)(15)
            dataArray(C, 17) = Dict.Item(k)(16) / Dict.Item(k)(14)
            dataArray(C, 18) = Dict.Item(k)(16)
        Next k
        On Error GoTo 0
        Set rsNew = ADOCopyArrayIntoRecordset(argArray:=dataArray)
        If bRefresh = False Then
            Call Create_PivotTable_ODBC_MO
        End If
    Else
        MsgBox "Unable to create this report. Make sure the WBS codes you have selected have values associated.", vbCritical, "Unable To Create Report"
    End If
End Sub

Function Level_FB(sIndex As String)
    Set xLvl3 = XDoc.SelectNodes(sXpath1 & "[Index=" & sIndex & "]")
    For Each nd1 In xLvl3
        sLvl3 = nd1.SelectSingleNode("Name").text
        sCode3 = nd1.SelectSingleNode(sLvl1Code).text
    Next nd1
    Set xLvl1 = XDoc.SelectNodes(sXpath1 & "[HierCode =" & Left(sCode3, 3) & "]")
    For Each nd2 In xLvl1
        sLvl1 = nd2.SelectSingleNode(sLvl1Code).text & " " & nd2.SelectSingleNode("Name").text
        sCode1 = nd2.SelectSingleNode(sLvl1Code).text
    Next nd2
    Set xLvl2 = XDoc.SelectNodes(sXpath1 & "[HierCode =" & Left(sCode3, 6) & "]")
    For Each nd3 In xLvl2
        sLvl2 = nd3.SelectSingleNode("Name").text
        sCode2 = nd3.SelectSingleNode(sLvl1Code).text
    Next nd3
    sCode2 = Mid(sCode2, 1, 3) & "." & Mid(sCode2, 4, 3)
    sLvl2 = sCode2 & " " & sLvl2
    sCode3 = Mid(sCode3, 1, 3) & "." & Mid(sCode3, 4, 3) & "." & Mid(sCode3, 7, 3)
    sLvl3 = sCode3 & " " & sLvl3
End Function

Function Level_FB2(sIndex As String)
    Set xLvl4 = XDoc.SelectNodes(sXpath2 & "[Index=" & sIndex & "]")
    For Each nd2 In xLvl4
        sLvl4 = nd2.SelectSingleNode("Name").text
        sCode4 = nd2.SelectSingleNode(sLvl2Code).text
    Next nd2
    Set xLvl2 = XDoc.SelectNodes(sXpath2 & "[HierCode =" & Left(sCode4, 3) & "]")
    For Each nd3 In xLvl2
        sLvl2 = nd3.SelectSingleNode(sLvl2Code).text & " " & nd3.SelectSingleNode("Name").text
        sCode2 = nd3.SelectSingleNode(sLvl2Code).text
    Next nd3
    Set xLvl3 = XDoc.SelectNodes(sXpath2 & "[HierCode =" & Left(sCode4, 6) & "]")
    For Each nd4 In xLvl3
        sLvl3 = nd4.SelectSingleNode("Name").text
        sCode3 = nd4.SelectSingleNode(sLvl2Code).text
    Next nd4
    sCode3 = Mid(sCode4, 1, 3) & "." & Mid(sCode4, 4, 3)
    sLvl3 = sCode3 & " " & sLvl3
    sCode4 = Mid(sCode4, 1, 3) & "." & Mid(sCode4, 4, 3) & "." & Mid(sCode4, 7, 3)
    sLvl4 = sCode4 & " " & sLvl4
End Function

Function Level_FB3(sIndex As String)
    Set xLvl5 = XDoc.SelectNodes(sXpath3 & "[Index=" & sIndex & "]")
    For Each nd3 In xLvl5
        sLvl5 = nd3.SelectSingleNode("Name").text
        sCode5 = nd3.SelectSingleNode(sLvl3Code).text
    Next nd3
    Set xLvl3 = XDoc.SelectNodes(sXpath3 & "[HierCode =" & Left(sCode5, 3) & "]")
    For Each nd4 In xLvl3
        sLvl3 = nd4.SelectSingleNode(sLvl3Code).text & " " & nd4.SelectSingleNode("Name").text
        sCode3 = nd4.SelectSingleNode(sLvl3Code).text
    Next nd4
    Set xLvl4 = XDoc.SelectNodes(sXpath3 & "[HierCode =" & Left(sCode5, 6) & "]")
    For Each nd5 In xLvl4
        sLvl4 = nd5.SelectSingleNode("Name").text
        sCode4 = nd5.SelectSingleNode(sLvl3Code).text
    Next nd5
    sCode4 = Mid(sCode5, 1, 3) & "." & Mid(sCode5, 4, 3)
    sLvl4 = sCode4 & " " & sLvl4
    sCode5 = Mid(sCode5, 1, 3) & "." & Mid(sCode5, 4, 3) & "." & Mid(sCode5, 7, 3)
    sLvl5 = sCode5 & " " & sLvl5
End Function
