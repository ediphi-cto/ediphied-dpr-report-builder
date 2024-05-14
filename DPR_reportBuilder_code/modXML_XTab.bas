Attribute VB_Name = "modXML_XTab"
Sub xmlXTabLevel1() 'Build 1 level array for XTab report


    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    dJobSz = Range("rngJobSize").Value
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
    'Set Column Level
        Set xLvl0 = XDoc.SelectNodes(sXpath0 & "[Index=" & nd.SelectSingleNode(sLvl0xNd).text & "]")
        For Each nd0 In xLvl0
            If bCkb0 = True Then
                If sLvl0xNd = "Division" Then
                    sLvl0 = Left(nd0.SelectSingleNode(sLvl0Code).text, 2) & "-" & nd0.SelectSingleNode("Name").text
                Else
                    sLvl0 = nd0.SelectSingleNode(sLvl0Code).text & "-" & nd0.SelectSingleNode("Name").text
                End If
            Else
                sLvl0 = nd0.SelectSingleNode("Name").text
            End If
            If nd0.SelectSingleNode("Unit").text <> "" Then
                sLUnit = nd0.SelectSingleNode("Unit").text
            Else
                sLUnit = Range("rngJobUnitName").Value
            End If
            sCode0 = nd0.SelectSingleNode(sLvl0Code).text
            dVal0 = CDbl(nd0.SelectSingleNode("Quantity").text)
        Next nd0
    'Set Level 1
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
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl0xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl1xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl0xNd).text) & CInt(nd.SelectSingleNode(sLvl1xNd).text)
        End If
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode0, sLvl0, _
                                sCode1, sLvl1, _
                                nd.SelectSingleNode("Description").text, _
                                Val(dVal0), _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)), _
                                sLUnit)
        Else
            q = Dict.Item(oTxt)
            q(6) = q(6) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
    Next nd
     
    ReDim dataArray(1 To Dict.count, 1 To 9)
    On Error Resume Next
    C = 0
    For Each k In Dict.Keys
        C = C + 1
        dataArray(C, 1) = Dict.Item(k)(0) 'lvl0code
        dataArray(C, 2) = Dict.Item(k)(1) 'lvl0item
        dataArray(C, 3) = Dict.Item(k)(2) 'lvl1code
        dataArray(C, 4) = Dict.Item(k)(3) 'lvl1item
        dataArray(C, 5) = Dict.Item(k)(4) 'description
        dataArray(C, 6) = Dict.Item(k)(5) & " " & Dict.Item(k)(7) 'levelquantity
        dataArray(C, 7) = Dict.Item(k)(6) 'amount
        dataArray(C, 8) = Dict.Item(k)(6) / Dict.Item(k)(5) 'cost/unit
        dataArray(C, 9) = Dict.Item(k)(6) / dJobSz 'cost/sf
    Next k
    On Error GoTo 0
    Set rsNew = ADOCopyArrayIntoRecordsetXtab(argArray:=dataArray)
End Sub

Sub xmlXTabLevel2() 'Build 2 level array for XTab report


    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    dJobSz = Range("rngJobSize").Value
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
    'Set Column Level
        Set xLvl0 = XDoc.SelectNodes(sXpath0 & "[Index=" & nd.SelectSingleNode(sLvl0xNd).text & "]")
        For Each nd0 In xLvl0
            If bCkb0 = True Then
                If sLvl0xNd = "Division" Then
                    sLvl0 = Left(nd0.SelectSingleNode(sLvl0Code).text, 2) & "-" & nd0.SelectSingleNode("Name").text
                Else
                    sLvl0 = nd0.SelectSingleNode(sLvl0Code).text & "-" & nd0.SelectSingleNode("Name").text
                End If
            Else
                sLvl0 = nd0.SelectSingleNode("Name").text
            End If
            If nd0.SelectSingleNode("Unit").text <> "" Then
                sLUnit = nd0.SelectSingleNode("Unit").text
            Else
                sLUnit = Range("rngJobUnitName").Value
            End If
            sCode0 = nd0.SelectSingleNode(sLvl0Code).text
            dVal0 = CDbl(nd0.SelectSingleNode("Quantity").text)
        Next nd0
    'Set Level 1
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
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl0xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl1xNd).text) & CInt(nd.SelectSingleNode(sLvl2xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl0xNd).text) & CInt(nd.SelectSingleNode(sLvl1xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl2xNd).text)
        End If
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode0, sLvl0, _
                                sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                nd.SelectSingleNode("Description").text, _
                                Val(dVal0), _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)), _
                                sLUnit)
        Else
            q = Dict.Item(oTxt)
            q(8) = q(8) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
    Next nd
     
    ReDim dataArray(1 To Dict.count, 1 To 11)
    On Error Resume Next
    C = 0
    For Each k In Dict.Keys
        C = C + 1
        dataArray(C, 1) = Dict.Item(k)(0) 'lvl0code
        dataArray(C, 2) = Dict.Item(k)(1) 'lvl0item
        dataArray(C, 3) = Dict.Item(k)(2) 'lvl1code
        dataArray(C, 4) = Dict.Item(k)(3) 'lvl1item
        dataArray(C, 5) = Dict.Item(k)(4) 'lvl2code
        dataArray(C, 6) = Dict.Item(k)(5) 'lvl2item
        dataArray(C, 7) = Dict.Item(k)(6) 'description
        dataArray(C, 8) = Dict.Item(k)(7) & " " & Dict.Item(k)(9) 'levelquantity
        dataArray(C, 9) = Dict.Item(k)(8) 'amount
        dataArray(C, 10) = Dict.Item(k)(8) / Dict.Item(k)(7) 'cost/unit
        dataArray(C, 11) = Dict.Item(k)(8) / dJobSz 'cost/sf
    Next k
    On Error GoTo 0
    Set rsNew = ADOCopyArrayIntoRecordsetXtab(argArray:=dataArray)
End Sub

Sub xmlXTabLevel3() 'Build 3 level array for XTab report

    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    
    dJobSz = Range("rngJobSize").Value
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
    'Set Column Level
        Set xLvl0 = XDoc.SelectNodes(sXpath0 & "[Index=" & nd.SelectSingleNode(sLvl0xNd).text & "]")
        For Each nd0 In xLvl0
            If bCkb0 = True Then
                If sLvl0xNd = "Division" Then
                    sLvl0 = Left(nd0.SelectSingleNode(sLvl0Code).text, 2) & "-" & nd0.SelectSingleNode("Name").text
                Else
                    sLvl0 = nd0.SelectSingleNode(sLvl0Code).text & "-" & nd0.SelectSingleNode("Name").text
                End If
            Else
                sLvl0 = nd0.SelectSingleNode("Name").text
            End If
            If nd0.SelectSingleNode("Unit").text <> "" Then
                sLUnit = nd0.SelectSingleNode("Unit").text
            Else
                sLUnit = Range("rngJobUnitName").Value
            End If
            sCode0 = nd0.SelectSingleNode(sLvl0Code).text
            dVal0 = CDbl(nd0.SelectSingleNode("Quantity").text)
        Next nd0
    'Set Level 1
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
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl0xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl1xNd).text) & CInt(nd.SelectSingleNode(sLvl2xNd).text) & CInt(nd.SelectSingleNode(sLvl3xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl0xNd).text) & CInt(nd.SelectSingleNode(sLvl1xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl2xNd).text) & CInt(nd.SelectSingleNode(sLvl3xNd).text)
        End If
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode0, sLvl0, _
                                sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                sCode3, sLvl3, _
                                nd.SelectSingleNode("Description").text, _
                                Val(dVal0), _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)), _
                                sLUnit)
        Else
            q = Dict.Item(oTxt)
            q(10) = q(10) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
    Next nd
     
    ReDim dataArray(1 To Dict.count, 1 To 13)
    On Error Resume Next
    C = 0
    For Each k In Dict.Keys
        C = C + 1
        dataArray(C, 1) = Dict.Item(k)(0) 'lvl0code
        dataArray(C, 2) = Dict.Item(k)(1) 'lvl0item
        dataArray(C, 3) = Dict.Item(k)(2) 'lvl1code
        dataArray(C, 4) = Dict.Item(k)(3) 'lvl1item
        dataArray(C, 5) = Dict.Item(k)(4) 'lvl2code
        dataArray(C, 6) = Dict.Item(k)(5) 'lvl2item
        dataArray(C, 7) = Dict.Item(k)(6) 'lvl3code
        dataArray(C, 8) = Dict.Item(k)(7) 'lvl3item
        dataArray(C, 9) = Dict.Item(k)(8) 'description
        dataArray(C, 10) = Dict.Item(k)(9) & " " & Dict.Item(k)(11) 'levelquantity
        dataArray(C, 11) = Dict.Item(k)(10) 'amount
        dataArray(C, 12) = Dict.Item(k)(10) / Dict.Item(k)(9) 'cost/unit
        dataArray(C, 13) = Dict.Item(k)(10) / dJobSz 'cost/sf
    Next k
    On Error GoTo 0
    Set rsNew = ADOCopyArrayIntoRecordsetXtab(argArray:=dataArray)
End Sub

Sub xmlXTabLevel4() 'Build 4 level array for XTab report

    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    
    dJobSz = Range("rngJobSize").Value
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
    'Set Column Level
        Set xLvl0 = XDoc.SelectNodes(sXpath0 & "[Index=" & nd.SelectSingleNode(sLvl0xNd).text & "]")
        For Each nd0 In xLvl0
            If bCkb0 = True Then
                If sLvl0xNd = "Division" Then
                    sLvl0 = Left(nd0.SelectSingleNode(sLvl0Code).text, 2) & "-" & nd0.SelectSingleNode("Name").text
                Else
                    sLvl0 = nd0.SelectSingleNode(sLvl0Code).text & "-" & nd0.SelectSingleNode("Name").text
                End If
            Else
                sLvl0 = nd0.SelectSingleNode("Name").text
            End If
            If nd0.SelectSingleNode("Unit").text <> "" Then
                sLUnit = nd0.SelectSingleNode("Unit").text
            Else
                sLUnit = Range("rngJobUnitName").Value
            End If
            sCode0 = nd0.SelectSingleNode(sLvl0Code).text
            dVal0 = CDbl(nd0.SelectSingleNode("Quantity").text)
        Next nd0
    'Set Level 1
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
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl0xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl1xNd).text) & CInt(nd.SelectSingleNode(sLvl2xNd).text) & CInt(nd.SelectSingleNode(sLvl3xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl4xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl0xNd).text) & CInt(nd.SelectSingleNode(sLvl1xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl2xNd).text) & CInt(nd.SelectSingleNode(sLvl3xNd).text) & CInt(nd.SelectSingleNode(sLvl4xNd).text)
        End If
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode0, sLvl0, _
                                sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                sCode3, sLvl3, _
                                sCode4, sLvl4, _
                                nd.SelectSingleNode("Description").text, _
                                Val(dVal0), _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)), _
                                sLUnit)
        Else
            q = Dict.Item(oTxt)
            q(12) = q(12) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
    Next nd
    ReDim dataArray(1 To Dict.count, 1 To 15)
    On Error Resume Next
    C = 0
    For Each k In Dict.Keys
        C = C + 1
        dataArray(C, 1) = Dict.Item(k)(0) 'lvl0code
        dataArray(C, 2) = Dict.Item(k)(1) 'lvl0item
        dataArray(C, 3) = Dict.Item(k)(2) 'lvl1code
        dataArray(C, 4) = Dict.Item(k)(3) 'lvl1item
        dataArray(C, 5) = Dict.Item(k)(4) 'lvl2code
        dataArray(C, 6) = Dict.Item(k)(5) 'lvl2item
        dataArray(C, 7) = Dict.Item(k)(6) 'lvl3code
        dataArray(C, 8) = Dict.Item(k)(7) 'lvl3item
        dataArray(C, 9) = Dict.Item(k)(8) 'lvl4code
        dataArray(C, 10) = Dict.Item(k)(9) 'lvl4item
        dataArray(C, 11) = Dict.Item(k)(10) 'description
        dataArray(C, 12) = Dict.Item(k)(11) & " " & Dict.Item(k)(13) 'levelquantity
        dataArray(C, 13) = Dict.Item(k)(12) 'amount
        dataArray(C, 14) = Dict.Item(k)(12) / Dict.Item(k)(11) 'cost/unit
        dataArray(C, 15) = Dict.Item(k)(12) / dJobSz 'cost/sf
    Next k
    On Error GoTo 0
    Set rsNew = ADOCopyArrayIntoRecordsetXtab(argArray:=dataArray)
End Sub


Sub xmlXTabLevel5() 'Build 5 level array for XTab report


    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    
    dJobSz = Range("rngJobSize").Value
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
    'Set Column Level
        Set xLvl0 = XDoc.SelectNodes(sXpath0 & "[Index=" & nd.SelectSingleNode(sLvl0xNd).text & "]")
        For Each nd0 In xLvl0
            If bCkb0 = True Then
                If sLvl0xNd = "Division" Then
                    sLvl0 = Left(nd0.SelectSingleNode(sLvl0Code).text, 2) & "-" & nd0.SelectSingleNode("Name").text
                Else
                    sLvl0 = nd0.SelectSingleNode(sLvl0Code).text & "-" & nd0.SelectSingleNode("Name").text
                End If
            Else
                sLvl0 = nd0.SelectSingleNode("Name").text
            End If
            If nd0.SelectSingleNode("Unit").text <> "" Then
                sLUnit = nd0.SelectSingleNode("Unit").text
            Else
                sLUnit = Range("rngJobUnitName").Value
            End If
            sCode0 = nd0.SelectSingleNode(sLvl0Code).text
            dVal0 = CDbl(nd0.SelectSingleNode("Quantity").text)
        Next nd0
    'Set Level 1
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
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl0xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl1xNd).text) & CInt(nd.SelectSingleNode(sLvl2xNd).text) & CInt(nd.SelectSingleNode(sLvl3xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl4xNd).text) & CInt(nd.SelectSingleNode(sLvl5xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl0xNd).text) & CInt(nd.SelectSingleNode(sLvl1xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl2xNd).text) & CInt(nd.SelectSingleNode(sLvl3xNd).text) & CInt(nd.SelectSingleNode(sLvl4xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl5xNd).text)
        End If
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode0, sLvl0, _
                                sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                sCode3, sLvl3, _
                                sCode4, sLvl4, _
                                sCode5, sLvl5, _
                                nd.SelectSingleNode("Description").text, _
                                Val(dVal0), _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)), _
                                sLUnit)
        Else
            q = Dict.Item(oTxt)
            q(14) = q(14) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
    Next nd
     
    ReDim dataArray(1 To Dict.count, 1 To 17)
    On Error Resume Next
    C = 0
    For Each k In Dict.Keys
        C = C + 1
        dataArray(C, 1) = Dict.Item(k)(0) 'lvl0code
        dataArray(C, 2) = Dict.Item(k)(1) 'lvl0item
        dataArray(C, 3) = Dict.Item(k)(2) 'lvl1code
        dataArray(C, 4) = Dict.Item(k)(3) 'lvl1item
        dataArray(C, 5) = Dict.Item(k)(4) 'lvl2code
        dataArray(C, 6) = Dict.Item(k)(5) 'lvl2item
        dataArray(C, 7) = Dict.Item(k)(6) 'lvl3code
        dataArray(C, 8) = Dict.Item(k)(7) 'lvl3item
        dataArray(C, 9) = Dict.Item(k)(8) 'lvl4code
        dataArray(C, 10) = Dict.Item(k)(9) 'lvl4item
        dataArray(C, 11) = Dict.Item(k)(10) 'lvl5code
        dataArray(C, 12) = Dict.Item(k)(11) 'lvl5item
        dataArray(C, 13) = Dict.Item(k)(12) 'description
        dataArray(C, 14) = Dict.Item(k)(13) & " " & Dict.Item(k)(15) 'levelquantity
        dataArray(C, 15) = Dict.Item(k)(14) 'amount
        dataArray(C, 16) = Dict.Item(k)(14) / Dict.Item(k)(13) 'cost/unit
        dataArray(C, 17) = Dict.Item(k)(14) / dJobSz 'cost/sf
    Next k
    On Error GoTo 0
    Set rsNew = ADOCopyArrayIntoRecordsetXtab(argArray:=dataArray)
End Sub

Private Function ADOCopyArrayIntoRecordsetXtab(argArray As Variant) As ADODB.Recordset 'Create data recordset for pivot cache
Dim rsADO As ADODB.Recordset
Dim lngR As Long
Dim lngC As Long

    Set rsADO = New ADODB.Recordset
    For i = 0 To iLvl
        Select Case i
            Case 0
                sLvl0Code = sLvl0Code & i
                rsADO.Fields.Append sLvl0Code, adVariant
                rsADO.Fields.Append sLvl0Item, adVariant
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
    rsADO.Fields.Append "Description", adVariant
    rsADO.Fields.Append "LevelQuantity", adVariant
    rsADO.Fields.Append "Amount", adVariant
    rsADO.Fields.Append "Cost/Unit", adVariant
    rsADO.Fields.Append "Cost/SF", adVariant
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
    Set ADOCopyArrayIntoRecordsetXtab = rsADO
    Set rsADO = Nothing
End Function

'*********************************************************
'Special FB code added for WBS 14 3-level sort. 05-19-2020
'*********************************************************

Sub xmlXTabLevel3FB() 'Build 3 level array for XTab report

    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    
    dJobSz = Range("rngJobSize").Value
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
    'Set Column Level
        Set xLvl0 = XDoc.SelectNodes(sXpath0 & "[Index=" & nd.SelectSingleNode(sLvl0xNd).text & "]")
        For Each nd0 In xLvl0
            If bCkb0 = True Then
                If sLvl0xNd = "Division" Then
                    sLvl0 = Left(nd0.SelectSingleNode(sLvl0Code).text, 2) & "-" & nd0.SelectSingleNode("Name").text
                Else
                    sLvl0 = nd0.SelectSingleNode(sLvl0Code).text & "-" & nd0.SelectSingleNode("Name").text
                End If
            Else
                sLvl0 = nd0.SelectSingleNode("Name").text
            End If
            If nd0.SelectSingleNode("Unit").text <> "" Then
                sLUnit = nd0.SelectSingleNode("Unit").text
            Else
                sLUnit = Range("rngJobUnitName").Value
            End If
            sCode0 = nd0.SelectSingleNode(sLvl0Code).text
            dVal0 = CDbl(nd0.SelectSingleNode("Quantity").text)
        Next nd0
    'Set Level 1 thru 3
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
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl0xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl1xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl0xNd).text) & CInt(nd.SelectSingleNode(sLvl1xNd).text)
        End If
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode0, sLvl0, _
                                sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                sCode3, sLvl3, _
                                nd.SelectSingleNode("Description").text, _
                                Val(dVal0), _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)), _
                                sLUnit)
        Else
            q = Dict.Item(oTxt)
            q(10) = q(10) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
errNextNd:
    Next nd
     
    ReDim dataArray(1 To Dict.count, 1 To 13)
    On Error Resume Next
    C = 0
    For Each k In Dict.Keys
        C = C + 1
        dataArray(C, 1) = Dict.Item(k)(0) 'lvl0code
        dataArray(C, 2) = Dict.Item(k)(1) 'lvl0item
        dataArray(C, 3) = Dict.Item(k)(2) 'lvl1code
        dataArray(C, 4) = Dict.Item(k)(3) 'lvl1item
        dataArray(C, 5) = Dict.Item(k)(4) 'lvl2code
        dataArray(C, 6) = Dict.Item(k)(5) 'lvl2item
        dataArray(C, 7) = Dict.Item(k)(6) 'lvl3code
        dataArray(C, 8) = Dict.Item(k)(7) 'lvl3item
        dataArray(C, 9) = Dict.Item(k)(8) 'description
        dataArray(C, 10) = Dict.Item(k)(9) & " " & Dict.Item(k)(11) 'levelquantity
        dataArray(C, 11) = Dict.Item(k)(10) 'amount
        dataArray(C, 12) = Dict.Item(k)(10) / Dict.Item(k)(9) 'cost/unit
        dataArray(C, 13) = Dict.Item(k)(10) / dJobSz 'cost/sf
    Next k
    On Error GoTo 0
    Set rsNew = ADOCopyArrayIntoRecordsetXtab(argArray:=dataArray)
End Sub

Sub xmlXTabLevel4FB() 'Build 4 level array for XTab report

    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    
    dJobSz = Range("rngJobSize").Value
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
    'Set Column Level
        Set xLvl0 = XDoc.SelectNodes(sXpath0 & "[Index=" & nd.SelectSingleNode(sLvl0xNd).text & "]")
        For Each nd0 In xLvl0
            If bCkb0 = True Then
                If sLvl0xNd = "Division" Then
                    sLvl0 = Left(nd0.SelectSingleNode(sLvl0Code).text, 2) & "-" & nd0.SelectSingleNode("Name").text
                Else
                    sLvl0 = nd0.SelectSingleNode(sLvl0Code).text & "-" & nd0.SelectSingleNode("Name").text
                End If
            Else
                sLvl0 = nd0.SelectSingleNode("Name").text
            End If
            If nd0.SelectSingleNode("Unit").text <> "" Then
                sLUnit = nd0.SelectSingleNode("Unit").text
            Else
                sLUnit = Range("rngJobUnitName").Value
            End If
            sCode0 = nd0.SelectSingleNode(sLvl0Code).text
            dVal0 = CDbl(nd0.SelectSingleNode("Quantity").text)
        Next nd0
    'Set Level 1
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
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl0xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl1xNd).text) & CInt(nd.SelectSingleNode(sLvl2xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl0xNd).text) & CInt(nd.SelectSingleNode(sLvl1xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl2xNd).text)
        End If
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode0, sLvl0, _
                                sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                sCode3, sLvl3, _
                                sCode4, sLvl4, _
                                nd.SelectSingleNode("Description").text, _
                                Val(dVal0), _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)), _
                                sLUnit)
        Else
            q = Dict.Item(oTxt)
            q(12) = q(12) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
errNextNd:
    Next nd
    ReDim dataArray(1 To Dict.count, 1 To 15)
    On Error Resume Next
    C = 0
    For Each k In Dict.Keys
        C = C + 1
        dataArray(C, 1) = Dict.Item(k)(0) 'lvl0code
        dataArray(C, 2) = Dict.Item(k)(1) 'lvl0item
        dataArray(C, 3) = Dict.Item(k)(2) 'lvl1code
        dataArray(C, 4) = Dict.Item(k)(3) 'lvl1item
        dataArray(C, 5) = Dict.Item(k)(4) 'lvl2code
        dataArray(C, 6) = Dict.Item(k)(5) 'lvl2item
        dataArray(C, 7) = Dict.Item(k)(6) 'lvl3code
        dataArray(C, 8) = Dict.Item(k)(7) 'lvl3item
        dataArray(C, 9) = Dict.Item(k)(8) 'lvl4code
        dataArray(C, 10) = Dict.Item(k)(9) 'lvl4item
        dataArray(C, 11) = Dict.Item(k)(10) 'description
        dataArray(C, 12) = Dict.Item(k)(11) & " " & Dict.Item(k)(13) 'levelquantity
        dataArray(C, 13) = Dict.Item(k)(12) 'amount
        dataArray(C, 14) = Dict.Item(k)(12) / Dict.Item(k)(11) 'cost/unit
        dataArray(C, 15) = Dict.Item(k)(12) / dJobSz 'cost/sf
    Next k
    On Error GoTo 0
    Set rsNew = ADOCopyArrayIntoRecordsetXtab(argArray:=dataArray)
End Sub

Sub xmlXTabLevel5FB() 'Build 5 level array for XTab report

    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    
    dJobSz = Range("rngJobSize").Value
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
    'Set Column Level
        Set xLvl0 = XDoc.SelectNodes(sXpath0 & "[Index=" & nd.SelectSingleNode(sLvl0xNd).text & "]")
        For Each nd0 In xLvl0
            If bCkb0 = True Then
                If sLvl0xNd = "Division" Then
                    sLvl0 = Left(nd0.SelectSingleNode(sLvl0Code).text, 2) & "-" & nd0.SelectSingleNode("Name").text
                Else
                    sLvl0 = nd0.SelectSingleNode(sLvl0Code).text & "-" & nd0.SelectSingleNode("Name").text
                End If
            Else
                sLvl0 = nd0.SelectSingleNode("Name").text
            End If
            If nd0.SelectSingleNode("Unit").text <> "" Then
                sLUnit = nd0.SelectSingleNode("Unit").text
            Else
                sLUnit = Range("rngJobUnitName").Value
            End If
            sCode0 = nd0.SelectSingleNode(sLvl0Code).text
            dVal0 = CDbl(nd0.SelectSingleNode("Quantity").text)
        Next nd0
    'Set Level 1
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
        If nd.SelectSingleNode(sLvl5xNd).text = 0 Then GoTo errNextNd
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
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl0xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl1xNd).text) & CInt(nd.SelectSingleNode(sLvl2xNd).text) & CInt(nd.SelectSingleNode(sLvl3xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl0xNd).text) & CInt(nd.SelectSingleNode(sLvl1xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl2xNd).text) & CInt(nd.SelectSingleNode(sLvl3xNd).text)
        End If
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode0, sLvl0, _
                                sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                sCode3, sLvl3, _
                                sCode4, sLvl4, _
                                sCode5, sLvl5, _
                                nd.SelectSingleNode("Description").text, _
                                Val(dVal0), _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)), _
                                sLUnit)
        Else
            q = Dict.Item(oTxt)
            q(14) = q(14) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
errNextNd:
    Next nd
     
    ReDim dataArray(1 To Dict.count, 1 To 17)
    On Error Resume Next
    C = 0
    For Each k In Dict.Keys
        C = C + 1
        dataArray(C, 1) = Dict.Item(k)(0) 'lvl0code
        dataArray(C, 2) = Dict.Item(k)(1) 'lvl0item
        dataArray(C, 3) = Dict.Item(k)(2) 'lvl1code
        dataArray(C, 4) = Dict.Item(k)(3) 'lvl1item
        dataArray(C, 5) = Dict.Item(k)(4) 'lvl2code
        dataArray(C, 6) = Dict.Item(k)(5) 'lvl2item
        dataArray(C, 7) = Dict.Item(k)(6) 'lvl3code
        dataArray(C, 8) = Dict.Item(k)(7) 'lvl3item
        dataArray(C, 9) = Dict.Item(k)(8) 'lvl4code
        dataArray(C, 10) = Dict.Item(k)(9) 'lvl4item
        dataArray(C, 11) = Dict.Item(k)(10) 'lvl5code
        dataArray(C, 12) = Dict.Item(k)(11) 'lvl5item
        dataArray(C, 13) = Dict.Item(k)(12) 'description
        dataArray(C, 14) = Dict.Item(k)(13) & " " & Dict.Item(k)(15) 'levelquantity
        dataArray(C, 15) = Dict.Item(k)(14) 'amount
        dataArray(C, 16) = Dict.Item(k)(14) / Dict.Item(k)(13) 'cost/unit
        dataArray(C, 17) = Dict.Item(k)(14) / dJobSz 'cost/sf
    Next k
    On Error GoTo 0
    Set rsNew = ADOCopyArrayIntoRecordsetXtab(argArray:=dataArray)
End Sub

