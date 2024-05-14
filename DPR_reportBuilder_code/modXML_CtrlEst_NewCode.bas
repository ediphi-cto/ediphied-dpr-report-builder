Attribute VB_Name = "modXML_CtrlEst_NewCode"
Dim sCat10 As String
Dim sCat20 As String
Dim sCat30 As String
Dim sCat40 As String
Dim sCat50 As String
Dim sCat51 As String
Dim sCat52 As String
Dim sCat60 As String
Dim sCat61 As String
Dim sCat62 As String
Dim sCat70 As String
Dim dval As Double

Function xNodeVal(nVal As String) As String
    Set xLvlSub = XDoc.SelectNodes("/Estimate/JobCostCategoryTable/JobCostCategory[Index=" & nVal & "]")
    For Each nd0 In xLvlSub
        xNodeVal = nd0.SelectSingleNode("Code").text
    Next nd0
End Function

Sub xmlCtrlEst1() 'Build 1 level array
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare

    sCat10 = 0
    sCat20 = 0
    sCat30 = 0
    sCat40 = 0
    sCat50 = 0
    sCat51 = 0
    sCat52 = 0
    sCat60 = 0
    sCat61 = 0
    sCat62 = 0
    sCat70 = 0
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
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
        If bCkbAll = True Then
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl1xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl1xNd).text)
        End If
'       Check if UserText6 code applied
        If nd.SelectSingleNode("UserTotal").text > 0 And nd.SelectSingleNode("JobCostCategoryUser").text = "" Then
            CodeChk = "*~*" & nd.SelectSingleNode("Description").text
            bCode = True
        Else
            CodeChk = nd.SelectSingleNode("Description").text
        End If
        For i = 1 To 6
            If i = 1 Then
                dval = nd.SelectSingleNode("LaborTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryLabor").text)
            ElseIf i = 2 Then
                dval = nd.SelectSingleNode("MatTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryMat").text)
            ElseIf i = 3 Then
                dval = nd.SelectSingleNode("SubconTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategorySubcon").text)
            ElseIf i = 4 Then
                dval = nd.SelectSingleNode("EquipTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryEquip").text)
            ElseIf i = 5 Then
                dval = nd.SelectSingleNode("OtherTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryOther").text)
            ElseIf i = 6 Then
                dval = nd.SelectSingleNode("UserTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryUser").text)
            End If
            Select Case nVal
                Case 10: sCat10 = sCat10 + dval
                Case 20: sCat20 = sCat20 + dval
                Case 30: sCat30 = sCat30 + dval
                Case 40: sCat40 = sCat40 + dval
                Case 50: sCat50 = sCat50 + dval
                Case 51: sCat51 = sCat51 + dval
                Case 52: sCat52 = sCat52 + dval
                Case 60: sCat60 = sCat60 + dval
                Case 61: sCat61 = sCat61 + dval
                Case 62: sCat62 = sCat62 + dval
                Case 70: sCat70 = sCat70 + dval
                Case Else
            End Select
            dval = 0
        Next i
        On Error Resume Next
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                CodeChk, nd.SelectSingleNode("ItemNote").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                nd.SelectSingleNode("TakeoffUnit").text, _
                                Val(CDbl(nd.SelectSingleNode("LaborHours").text)), _
                                Val(CDbl(sCat10)), _
                                Val(CDbl(sCat20)), _
                                Val(CDbl(sCat30)), _
                                Val(CDbl(sCat40)), _
                                Val(CDbl(sCat50)), _
                                Val(CDbl(sCat51)), _
                                Val(CDbl(sCat52)), _
                                Val(CDbl(sCat60)), _
                                Val(CDbl(sCat61)), _
                                Val(CDbl(sCat62)), _
                                Val(CDbl(sCat70)), _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)))
        Else
            q = Dict.Item(oTxt)
            q(6) = q(6) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(8) = q(8) + Val(CDbl(nd.SelectSingleNode("LaborHours").text))
            q(9) = q(9) + Val(CDbl(sCat10))
            q(10) = q(10) + Val(CDbl(sCat20))
            q(11) = q(11) + Val(CDbl(sCat30))
            q(12) = q(12) + Val(CDbl(sCat40))
            q(13) = q(13) + Val(CDbl(sCat50))
            q(14) = q(14) + Val(CDbl(sCat51))
            q(15) = q(15) + Val(CDbl(sCat52))
            q(16) = q(16) + Val(CDbl(sCat60))
            q(17) = q(17) + Val(CDbl(sCat61))
            q(18) = q(18) + Val(CDbl(sCat62))
            q(19) = q(19) + Val(CDbl(sCat70))
            q(20) = q(20) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
        sCat10 = 0
        sCat20 = 0
        sCat30 = 0
        sCat40 = 0
        sCat50 = 0
        sCat51 = 0
        sCat52 = 0
        sCat60 = 0
        sCat61 = 0
        sCat62 = 0
        sCat70 = 0
     Next nd
     
    ReDim dataArray(1 To Dict.count, 1 To 21)
    On Error Resume Next
    C = 0
    For Each k In Dict.Keys
        C = C + 1
        dataArray(C, 1) = Dict.Item(k)(0) 'Lvl1Code
        dataArray(C, 2) = Dict.Item(k)(1) 'Lvl1Desc
        dataArray(C, 3) = Dict.Item(k)(2) 'SortOrder
        dataArray(C, 4) = Dict.Item(k)(3) 'Index
        dataArray(C, 5) = Dict.Item(k)(4) 'Desc
        dataArray(C, 6) = Dict.Item(k)(5) 'Note
        dataArray(C, 7) = Dict.Item(k)(6) 'TakeoffQty
        dataArray(C, 8) = Dict.Item(k)(7) 'TakeoffUnit
        If Dict.Item(k)(9) > 0 Then dataArray(C, 9) = Dict.Item(k)(8) 'Labor Hours
        dataArray(C, 10) = Dict.Item(k)(9) 'Labor10
        dataArray(C, 11) = Dict.Item(k)(10) 'Material20
        dataArray(C, 12) = Dict.Item(k)(11) 'Equip30
        dataArray(C, 13) = Dict.Item(k)(12) 'Other40
        dataArray(C, 14) = Dict.Item(k)(13) 'Sub50
        dataArray(C, 15) = Dict.Item(k)(14) 'DPR Est 51
        dataArray(C, 16) = Dict.Item(k)(15) 'DPR Cont 52
        dataArray(C, 17) = Dict.Item(k)(16) 'Owner Allow 60
        dataArray(C, 18) = Dict.Item(k)(17) 'ConstCont 61
        dataArray(C, 19) = Dict.Item(k)(18) 'OwnerCont 62
        dataArray(C, 20) = Dict.Item(k)(19) 'OH&P 70
        For i = 10 To 20
            dataArray(C, 21) = dataArray(C, 21) + dataArray(C, i) 'Total
        Next i
    Next k
    On Error GoTo 0
    Set rsNew = ADOCopyArrayIntoRecordsetCEst(argArray:=dataArray)
End Sub
Sub xmlCtrlEst2() 'Build 2 level array
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare

    sCat10 = 0
    sCat20 = 0
    sCat30 = 0
    sCat40 = 0
    sCat50 = 0
    sCat51 = 0
    sCat52 = 0
    sCat60 = 0
    sCat61 = 0
    sCat62 = 0
    sCat70 = 0
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
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
        If bCkbAll = True Then
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl2xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) & CInt(nd.SelectSingleNode(sLvl2xNd).text)
        End If
    'Check if UserText6 code applied
        If nd.SelectSingleNode("UserTotal").text > 0 And nd.SelectSingleNode("JobCostCategoryUser").text = "" Then
            CodeChk = "*~*" & nd.SelectSingleNode("Description").text
            bCode = True
        Else
            CodeChk = nd.SelectSingleNode("Description").text
        End If
        For i = 1 To 6
            If i = 1 Then
                dval = nd.SelectSingleNode("LaborTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryLabor").text)
            ElseIf i = 2 Then
                dval = nd.SelectSingleNode("MatTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryMat").text)
            ElseIf i = 3 Then
                dval = nd.SelectSingleNode("SubconTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategorySubcon").text)
            ElseIf i = 4 Then
                dval = nd.SelectSingleNode("EquipTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryEquip").text)
            ElseIf i = 5 Then
                dval = nd.SelectSingleNode("OtherTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryOther").text)
            ElseIf i = 6 Then
                dval = nd.SelectSingleNode("UserTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryUser").text)
            End If
            Select Case nVal
                Case 10: sCat10 = sCat10 + dval
                Case 20: sCat20 = sCat20 + dval
                Case 30: sCat30 = sCat30 + dval
                Case 40: sCat40 = sCat40 + dval
                Case 50: sCat50 = sCat50 + dval
                Case 51: sCat51 = sCat51 + dval
                Case 52: sCat52 = sCat52 + dval
                Case 60: sCat60 = sCat60 + dval
                Case 61: sCat61 = sCat61 + dval
                Case 62: sCat62 = sCat62 + dval
                Case 70: sCat70 = sCat70 + dval
                Case Else
            End Select
            dval = 0
        Next i
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                CodeChk, _
                                nd.SelectSingleNode("ItemNote").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                nd.SelectSingleNode("TakeoffUnit").text, _
                                Val(CDbl(nd.SelectSingleNode("LaborHours").text)), _
                                Val(CDbl(sCat10)), _
                                Val(CDbl(sCat20)), _
                                Val(CDbl(sCat30)), _
                                Val(CDbl(sCat40)), _
                                Val(CDbl(sCat50)), _
                                Val(CDbl(sCat51)), _
                                Val(CDbl(sCat52)), _
                                Val(CDbl(sCat60)), _
                                Val(CDbl(sCat61)), _
                                Val(CDbl(sCat62)), _
                                Val(CDbl(sCat70)), _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)))
        Else
            q = Dict.Item(oTxt)
            q(8) = q(8) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(10) = q(10) + Val(CDbl(nd.SelectSingleNode("LaborHours").text))
            q(11) = q(11) + Val(CDbl(sCat10))
            q(12) = q(12) + Val(CDbl(sCat20))
            q(13) = q(13) + Val(CDbl(sCat30))
            q(14) = q(14) + Val(CDbl(sCat40))
            q(15) = q(15) + Val(CDbl(sCat50))
            q(16) = q(16) + Val(CDbl(sCat51))
            q(17) = q(17) + Val(CDbl(sCat52))
            q(18) = q(18) + Val(CDbl(sCat60))
            q(19) = q(19) + Val(CDbl(sCat61))
            q(20) = q(20) + Val(CDbl(sCat62))
            q(21) = q(21) + Val(CDbl(sCat70))
            q(22) = q(22) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
        sCat10 = 0
        sCat20 = 0
        sCat30 = 0
        sCat40 = 0
        sCat50 = 0
        sCat51 = 0
        sCat52 = 0
        sCat60 = 0
        sCat61 = 0
        sCat62 = 0
        sCat70 = 0
    Next nd

    ReDim dataArray(1 To Dict.count, 1 To 23)
    On Error Resume Next
    C = 0
    For Each k In Dict.Keys
        C = C + 1
        dataArray(C, 1) = Dict.Item(k)(0) 'Lvl1Code
        dataArray(C, 2) = Dict.Item(k)(1) 'Lvl1Desc
        dataArray(C, 3) = Dict.Item(k)(2) 'Lvl2Code
        dataArray(C, 4) = Dict.Item(k)(3) 'Lvl2Desc
        dataArray(C, 5) = Dict.Item(k)(4) 'SortOrder
        dataArray(C, 6) = Dict.Item(k)(5) 'Index
        dataArray(C, 7) = Dict.Item(k)(6) 'Desc
        dataArray(C, 8) = Dict.Item(k)(7) 'Note
        dataArray(C, 9) = Dict.Item(k)(8) 'TakeoffQty
        dataArray(C, 10) = Dict.Item(k)(9) 'TakeoffUnit
        If Dict.Item(k)(11) > 0 Then dataArray(C, 11) = Dict.Item(k)(10) 'Labor Hours
        dataArray(C, 12) = Dict.Item(k)(11) 'Labor10
        dataArray(C, 13) = Dict.Item(k)(12) 'Material20
        dataArray(C, 14) = Dict.Item(k)(13) 'Equip30
        dataArray(C, 15) = Dict.Item(k)(14) 'Other40
        dataArray(C, 16) = Dict.Item(k)(15) 'Sub50
        dataArray(C, 17) = Dict.Item(k)(16) 'DPR Est 51
        dataArray(C, 18) = Dict.Item(k)(17) 'DPR Cont 52
        dataArray(C, 19) = Dict.Item(k)(18) 'Owner Allow 60
        dataArray(C, 20) = Dict.Item(k)(19) 'ConstCont 61
        dataArray(C, 21) = Dict.Item(k)(20) 'OwnerCont 62
        dataArray(C, 22) = Dict.Item(k)(21) 'OH&P 70
        For i = 12 To 22
            dataArray(C, 23) = dataArray(C, 23) + dataArray(C, i) 'Total
        Next i
    Next k
    On Error GoTo 0
    Set rsNew = ADOCopyArrayIntoRecordsetCEst(argArray:=dataArray)
End Sub

Sub xmlCtrlEst3() 'Build 3 level array
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare

    sCat10 = 0
    sCat20 = 0
    sCat30 = 0
    sCat40 = 0
    sCat50 = 0
    sCat51 = 0
    sCat52 = 0
    sCat60 = 0
    sCat61 = 0
    sCat62 = 0
    sCat70 = 0
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
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
        If bCkbAll = True Then
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl2xNd).text) & CInt(nd.SelectSingleNode(sLvl3xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) & CInt(nd.SelectSingleNode(sLvl2xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl3xNd).text)
            
        End If
    'Check if UserText6 code applied
        If nd.SelectSingleNode("UserTotal").text > 0 And nd.SelectSingleNode("JobCostCategoryUser").text = "" Then
            CodeChk = "*~*" & nd.SelectSingleNode("Description").text
            bCode = True
        Else
            CodeChk = nd.SelectSingleNode("Description").text
        End If
        For i = 1 To 6
            If i = 1 Then
                dval = nd.SelectSingleNode("LaborTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryLabor").text)
            ElseIf i = 2 Then
                dval = nd.SelectSingleNode("MatTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryMat").text)
            ElseIf i = 3 Then
                dval = nd.SelectSingleNode("SubconTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategorySubcon").text)
            ElseIf i = 4 Then
                dval = nd.SelectSingleNode("EquipTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryEquip").text)
            ElseIf i = 5 Then
                dval = nd.SelectSingleNode("OtherTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryOther").text)
            ElseIf i = 6 Then
                dval = nd.SelectSingleNode("UserTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryUser").text)
            End If
            Select Case nVal
                Case 10: sCat10 = sCat10 + dval
                Case 20: sCat20 = sCat20 + dval
                Case 30: sCat30 = sCat30 + dval
                Case 40: sCat40 = sCat40 + dval
                Case 50: sCat50 = sCat50 + dval
                Case 51: sCat51 = sCat51 + dval
                Case 52: sCat52 = sCat52 + dval
                Case 60: sCat60 = sCat60 + dval
                Case 61: sCat61 = sCat61 + dval
                Case 62: sCat62 = sCat62 + dval
                Case 70: sCat70 = sCat70 + dval
                Case Else
            End Select
            dval = 0
        Next i
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, sCode2, sLvl2, sCode3, sLvl3, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                CodeChk, _
                                nd.SelectSingleNode("ItemNote").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                nd.SelectSingleNode("TakeoffUnit").text, _
                                Val(CDbl(nd.SelectSingleNode("LaborHours").text)), _
                                Val(CDbl(sCat10)), _
                                Val(CDbl(sCat20)), _
                                Val(CDbl(sCat30)), _
                                Val(CDbl(sCat40)), _
                                Val(CDbl(sCat50)), _
                                Val(CDbl(sCat51)), _
                                Val(CDbl(sCat52)), _
                                Val(CDbl(sCat60)), _
                                Val(CDbl(sCat61)), _
                                Val(CDbl(sCat62)), _
                                Val(CDbl(sCat70)), _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)))
        Else
            q = Dict.Item(oTxt)
            q(10) = q(10) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(12) = q(12) + Val(CDbl(nd.SelectSingleNode("LaborHours").text))
            q(13) = q(13) + Val(CDbl(sCat10))
            q(14) = q(14) + Val(CDbl(sCat20))
            q(15) = q(15) + Val(CDbl(sCat30))
            q(16) = q(16) + Val(CDbl(sCat40))
            q(17) = q(17) + Val(CDbl(sCat50))
            q(18) = q(18) + Val(CDbl(sCat51))
            q(19) = q(19) + Val(CDbl(sCat52))
            q(20) = q(20) + Val(CDbl(sCat60))
            q(21) = q(21) + Val(CDbl(sCat61))
            q(22) = q(22) + Val(CDbl(sCat62))
            q(23) = q(23) + Val(CDbl(sCat70))
            q(24) = q(24) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
        sCat10 = 0
        sCat20 = 0
        sCat30 = 0
        sCat40 = 0
        sCat50 = 0
        sCat51 = 0
        sCat52 = 0
        sCat60 = 0
        sCat61 = 0
        sCat62 = 0
        sCat70 = 0
    Next nd

    ReDim dataArray(1 To Dict.count, 1 To 25)
    On Error Resume Next
    C = 0
    For Each k In Dict.Keys
        C = C + 1
        dataArray(C, 1) = Dict.Item(k)(0) 'Lvl1Code
        dataArray(C, 2) = Dict.Item(k)(1) 'Lvl1Desc
        dataArray(C, 3) = Dict.Item(k)(2) 'Lvl2Code
        dataArray(C, 4) = Dict.Item(k)(3) 'Lvl2Desc
        dataArray(C, 5) = Dict.Item(k)(4) 'Lvl3Code
        dataArray(C, 6) = Dict.Item(k)(5) 'Lvl3Desc
        dataArray(C, 7) = Dict.Item(k)(6) 'SortOrder
        dataArray(C, 8) = Dict.Item(k)(7) 'Index
        dataArray(C, 9) = Dict.Item(k)(8) 'Desc
        dataArray(C, 10) = Dict.Item(k)(9) 'Note
        dataArray(C, 11) = Dict.Item(k)(10) 'TakeoffQty
        dataArray(C, 12) = Dict.Item(k)(11) 'TakeoffUnit
        If Dict.Item(k)(13) > 0 Then dataArray(C, 13) = Dict.Item(k)(12) 'Labor Hours
        dataArray(C, 14) = Dict.Item(k)(13) 'Labor10
        dataArray(C, 15) = Dict.Item(k)(14) 'Material20
        dataArray(C, 16) = Dict.Item(k)(15) 'Equip30
        dataArray(C, 17) = Dict.Item(k)(16) 'Other40
        dataArray(C, 18) = Dict.Item(k)(17) 'Sub50
        dataArray(C, 19) = Dict.Item(k)(18) 'DPR Est 51
        dataArray(C, 20) = Dict.Item(k)(19) 'DPR Cont 52
        dataArray(C, 21) = Dict.Item(k)(20) 'Owner Allow 60
        dataArray(C, 22) = Dict.Item(k)(21) 'ConstCont 61
        dataArray(C, 23) = Dict.Item(k)(22) 'OwnerCont 62
        dataArray(C, 24) = Dict.Item(k)(23) 'OH&P 70
        For i = 14 To 24
            dataArray(C, 25) = dataArray(C, 25) + dataArray(C, i) 'Total
        Next i
    Next k
    On Error GoTo 0
    Set rsNew = ADOCopyArrayIntoRecordsetCEst(argArray:=dataArray)
End Sub

Sub xmlCtrlEst4() 'Build 4 level array
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare

    sCat10 = 0
    sCat20 = 0
    sCat30 = 0
    sCat40 = 0
    sCat50 = 0
    sCat51 = 0
    sCat52 = 0
    sCat60 = 0
    sCat61 = 0
    sCat62 = 0
    sCat70 = 0
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
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
        If bCkbAll = True Then
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl2xNd).text) & CInt(nd.SelectSingleNode(sLvl3xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl4xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) & CInt(nd.SelectSingleNode(sLvl2xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl3xNd).text) & CInt(nd.SelectSingleNode(sLvl4xNd).text)
        End If
    'Check if UserText6 code applied
        If nd.SelectSingleNode("UserTotal").text > 0 And nd.SelectSingleNode("JobCostCategoryUser").text = "" Then
            CodeChk = "*~*" & nd.SelectSingleNode("Description").text
            bCode = True
        Else
            CodeChk = nd.SelectSingleNode("Description").text
        End If
        For i = 1 To 6
            If i = 1 Then
                dval = nd.SelectSingleNode("LaborTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryLabor").text)
            ElseIf i = 2 Then
                dval = nd.SelectSingleNode("MatTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryMat").text)
            ElseIf i = 3 Then
                dval = nd.SelectSingleNode("SubconTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategorySubcon").text)
            ElseIf i = 4 Then
                dval = nd.SelectSingleNode("EquipTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryEquip").text)
            ElseIf i = 5 Then
                dval = nd.SelectSingleNode("OtherTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryOther").text)
            ElseIf i = 6 Then
                dval = nd.SelectSingleNode("UserTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryUser").text)
            End If
            Select Case nVal
                Case 10: sCat10 = sCat10 + dval
                Case 20: sCat20 = sCat20 + dval
                Case 30: sCat30 = sCat30 + dval
                Case 40: sCat40 = sCat40 + dval
                Case 50: sCat50 = sCat50 + dval
                Case 51: sCat51 = sCat51 + dval
                Case 52: sCat52 = sCat52 + dval
                Case 60: sCat60 = sCat60 + dval
                Case 61: sCat61 = sCat61 + dval
                Case 62: sCat62 = sCat62 + dval
                Case 70: sCat70 = sCat70 + dval
                Case Else
            End Select
            dval = 0
        Next i
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, sCode2, sLvl2, sCode3, sLvl3, sCode4, sLvl4, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                CodeChk, _
                                nd.SelectSingleNode("ItemNote").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                nd.SelectSingleNode("TakeoffUnit").text, _
                                Val(CDbl(nd.SelectSingleNode("LaborHours").text)), _
                                Val(CDbl(sCat10)), _
                                Val(CDbl(sCat20)), _
                                Val(CDbl(sCat30)), _
                                Val(CDbl(sCat40)), _
                                Val(CDbl(sCat50)), _
                                Val(CDbl(sCat51)), _
                                Val(CDbl(sCat52)), _
                                Val(CDbl(sCat60)), _
                                Val(CDbl(sCat61)), _
                                Val(CDbl(sCat62)), _
                                Val(CDbl(sCat70)), _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)))
        Else
            q = Dict.Item(oTxt)
            q(12) = q(12) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(14) = q(14) + Val(CDbl(nd.SelectSingleNode("LaborHours").text))
            q(15) = q(15) + Val(CDbl(sCat10))
            q(16) = q(16) + Val(CDbl(sCat20))
            q(17) = q(17) + Val(CDbl(sCat30))
            q(18) = q(18) + Val(CDbl(sCat40))
            q(19) = q(19) + Val(CDbl(sCat50))
            q(20) = q(20) + Val(CDbl(sCat51))
            q(21) = q(21) + Val(CDbl(sCat52))
            q(22) = q(22) + Val(CDbl(sCat60))
            q(23) = q(23) + Val(CDbl(sCat61))
            q(24) = q(24) + Val(CDbl(sCat62))
            q(25) = q(25) + Val(CDbl(sCat70))
            q(26) = q(26) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
        sCat10 = 0
        sCat20 = 0
        sCat30 = 0
        sCat40 = 0
        sCat50 = 0
        sCat51 = 0
        sCat52 = 0
        sCat60 = 0
        sCat61 = 0
        sCat62 = 0
        sCat70 = 0
    Next nd

    ReDim dataArray(1 To Dict.count, 1 To 27)
    On Error Resume Next
    C = 0
    For Each k In Dict.Keys
        C = C + 1
        dataArray(C, 1) = Dict.Item(k)(0) 'Lvl1Code
        dataArray(C, 2) = Dict.Item(k)(1) 'Lvl1Desc
        dataArray(C, 3) = Dict.Item(k)(2) 'Lvl2Code
        dataArray(C, 4) = Dict.Item(k)(3) 'Lvl2Desc
        dataArray(C, 5) = Dict.Item(k)(4) 'Lvl3Code
        dataArray(C, 6) = Dict.Item(k)(5) 'Lvl3Desc
        dataArray(C, 7) = Dict.Item(k)(6) 'Lvl4Code
        dataArray(C, 8) = Dict.Item(k)(7) 'Lvl4Desc
        dataArray(C, 9) = Dict.Item(k)(8) 'SortOrder
        dataArray(C, 10) = Dict.Item(k)(9) 'Index
        dataArray(C, 11) = Dict.Item(k)(10) 'Desc
        dataArray(C, 12) = Dict.Item(k)(11) 'Note
        dataArray(C, 13) = Dict.Item(k)(12) 'TakeoffQty
        dataArray(C, 14) = Dict.Item(k)(13) 'TakeoffUnit
        If Dict.Item(k)(15) > 0 Then dataArray(C, 15) = Dict.Item(k)(14) 'Labor Hours
        dataArray(C, 16) = Dict.Item(k)(15) 'Labor10
        dataArray(C, 17) = Dict.Item(k)(16) 'Material20
        dataArray(C, 18) = Dict.Item(k)(17) 'Equip30
        dataArray(C, 19) = Dict.Item(k)(18) 'Other40
        dataArray(C, 20) = Dict.Item(k)(19) 'Sub50
        dataArray(C, 21) = Dict.Item(k)(20) 'DPR Est 51
        dataArray(C, 22) = Dict.Item(k)(21) 'DPR Cont 52
        dataArray(C, 23) = Dict.Item(k)(22) 'Owner Allow 60
        dataArray(C, 24) = Dict.Item(k)(23) 'ConstCont 61
        dataArray(C, 25) = Dict.Item(k)(24) 'OwnerCont 62
        dataArray(C, 26) = Dict.Item(k)(25) 'OH&P 70
        For i = 16 To 26
            dataArray(C, 27) = dataArray(C, 27) + dataArray(C, i) 'Total
        Next i
    Next k
    On Error GoTo 0
    Set rsNew = ADOCopyArrayIntoRecordsetCEst(argArray:=dataArray)
End Sub

Sub xmlCtrlEst5() 'Build 5 level array
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare

    sCat10 = 0
    sCat20 = 0
    sCat30 = 0
    sCat40 = 0
    sCat50 = 0
    sCat51 = 0
    sCat52 = 0
    sCat60 = 0
    sCat61 = 0
    sCat62 = 0
    sCat70 = 0
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
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
        If bCkbAll = True Then
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl2xNd).text) & CInt(nd.SelectSingleNode(sLvl3xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl4xNd).text) & CInt(nd.SelectSingleNode(sLvl5xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) & CInt(nd.SelectSingleNode(sLvl2xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl3xNd).text) & CInt(nd.SelectSingleNode(sLvl4xNd).text) & CInt(nd.SelectSingleNode(sLvl5xNd).text)
        End If
    'Check if UserText6 code applied
        If nd.SelectSingleNode("UserTotal").text > 0 And nd.SelectSingleNode("JobCostCategoryUser").text = "" Then
            CodeChk = "*~*" & nd.SelectSingleNode("Description").text
            bCode = True
        Else
            CodeChk = nd.SelectSingleNode("Description").text
        End If
        For i = 1 To 6
            If i = 1 Then
                dval = nd.SelectSingleNode("LaborTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryLabor").text)
            ElseIf i = 2 Then
                dval = nd.SelectSingleNode("MatTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryMat").text)
            ElseIf i = 3 Then
                dval = nd.SelectSingleNode("SubconTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategorySubcon").text)
            ElseIf i = 4 Then
                dval = nd.SelectSingleNode("EquipTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryEquip").text)
            ElseIf i = 5 Then
                dval = nd.SelectSingleNode("OtherTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryOther").text)
            ElseIf i = 6 Then
                dval = nd.SelectSingleNode("UserTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryUser").text)
            End If
            Select Case nVal
                Case 10: sCat10 = sCat10 + dval
                Case 20: sCat20 = sCat20 + dval
                Case 30: sCat30 = sCat30 + dval
                Case 40: sCat40 = sCat40 + dval
                Case 50: sCat50 = sCat50 + dval
                Case 51: sCat51 = sCat51 + dval
                Case 52: sCat52 = sCat52 + dval
                Case 60: sCat60 = sCat60 + dval
                Case 61: sCat61 = sCat61 + dval
                Case 62: sCat62 = sCat62 + dval
                Case 70: sCat70 = sCat70 + dval
                Case Else
            End Select
            dval = 0
        Next i
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, sCode2, sLvl2, _
                                sCode3, sLvl3, sCode4, sLvl4, sCode5, sLvl5, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                CodeChk, _
                                nd.SelectSingleNode("ItemNote").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                nd.SelectSingleNode("TakeoffUnit").text, _
                                Val(CDbl(nd.SelectSingleNode("LaborHours").text)), _
                                Val(CDbl(sCat10)), _
                                Val(CDbl(sCat20)), _
                                Val(CDbl(sCat30)), _
                                Val(CDbl(sCat40)), _
                                Val(CDbl(sCat50)), _
                                Val(CDbl(sCat51)), _
                                Val(CDbl(sCat52)), _
                                Val(CDbl(sCat60)), _
                                Val(CDbl(sCat61)), _
                                Val(CDbl(sCat62)), _
                                Val(CDbl(sCat70)), _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)))
        Else
            q = Dict.Item(oTxt)
            q(14) = q(14) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(16) = q(16) + Val(CDbl(nd.SelectSingleNode("LaborHours").text))
            q(17) = q(17) + Val(CDbl(sCat10))
            q(18) = q(18) + Val(CDbl(sCat20))
            q(19) = q(19) + Val(CDbl(sCat30))
            q(20) = q(20) + Val(CDbl(sCat40))
            q(21) = q(21) + Val(CDbl(sCat50))
            q(22) = q(22) + Val(CDbl(sCat51))
            q(23) = q(23) + Val(CDbl(sCat52))
            q(24) = q(24) + Val(CDbl(sCat60))
            q(25) = q(25) + Val(CDbl(sCat61))
            q(26) = q(26) + Val(CDbl(sCat62))
            q(27) = q(27) + Val(CDbl(sCat70))
            q(28) = q(28) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
        sCat10 = 0
        sCat20 = 0
        sCat30 = 0
        sCat40 = 0
        sCat50 = 0
        sCat51 = 0
        sCat52 = 0
        sCat60 = 0
        sCat61 = 0
        sCat62 = 0
        sCat70 = 0
    Next nd

    ReDim dataArray(1 To Dict.count, 1 To 29)
    On Error Resume Next
    C = 0
    For Each k In Dict.Keys
        C = C + 1
        dataArray(C, 1) = Dict.Item(k)(0) 'Lvl1Code
        dataArray(C, 2) = Dict.Item(k)(1) 'Lvl1Desc
        dataArray(C, 3) = Dict.Item(k)(2) 'Lvl2Code
        dataArray(C, 4) = Dict.Item(k)(3) 'Lvl2Desc
        dataArray(C, 5) = Dict.Item(k)(4) 'Lvl3Code
        dataArray(C, 6) = Dict.Item(k)(5) 'Lvl3Desc
        dataArray(C, 7) = Dict.Item(k)(6) 'Lvl4Code
        dataArray(C, 8) = Dict.Item(k)(7) 'Lvl4Desc
        dataArray(C, 9) = Dict.Item(k)(8) 'Lvl5Code
        dataArray(C, 10) = Dict.Item(k)(9) 'Lvl5Desc
        dataArray(C, 11) = Dict.Item(k)(10) 'SortOrder
        dataArray(C, 12) = Dict.Item(k)(11) 'Index
        dataArray(C, 13) = Dict.Item(k)(12) 'Desc
        dataArray(C, 14) = Dict.Item(k)(13) 'Note
        dataArray(C, 15) = Dict.Item(k)(14) 'TakeoffQty
        dataArray(C, 16) = Dict.Item(k)(15) 'TakeoffUnit
        If Dict.Item(k)(17) > 0 Then dataArray(C, 17) = Dict.Item(k)(16) 'Labor Hours
        dataArray(C, 18) = Dict.Item(k)(17) 'Labor10
        dataArray(C, 19) = Dict.Item(k)(18) 'Material20
        dataArray(C, 20) = Dict.Item(k)(19) 'Equip30
        dataArray(C, 21) = Dict.Item(k)(20) 'Other40
        dataArray(C, 22) = Dict.Item(k)(21) 'Sub50
        dataArray(C, 23) = Dict.Item(k)(22) 'DPR Est 51
        dataArray(C, 24) = Dict.Item(k)(23) 'DPR Cont 52
        dataArray(C, 25) = Dict.Item(k)(24) 'Owner Allow 60
        dataArray(C, 26) = Dict.Item(k)(25) 'ConstCont 61
        dataArray(C, 27) = Dict.Item(k)(26) 'OwnerCont 62
        dataArray(C, 28) = Dict.Item(k)(27) 'OH&P 70
        For i = 18 To 28
            dataArray(C, 29) = dataArray(C, 29) + dataArray(C, i) 'Total
        Next i
    Next k
    On Error GoTo 0
    Set rsNew = ADOCopyArrayIntoRecordsetCEst(argArray:=dataArray)
End Sub

Private Function ADOCopyArrayIntoRecordsetCEst(argArray As Variant) As ADODB.RecordSet 'Create data recordset for pivot cache
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
    rsADO.Fields.Append "Manhours", adVariant
    rsADO.Fields.Append "Labor10", adVariant
    rsADO.Fields.Append "Material20", adVariant
    rsADO.Fields.Append "Equipment30", adVariant
    rsADO.Fields.Append "Other40", adVariant
    rsADO.Fields.Append "Sub50", adVariant
    rsADO.Fields.Append "DPREst51", adVariant
    rsADO.Fields.Append "DPRCont52", adVariant
    rsADO.Fields.Append "OwnerAllow60", adVariant
    rsADO.Fields.Append "ConstCont61", adVariant
    rsADO.Fields.Append "OwnerCont62", adVariant
    rsADO.Fields.Append "OHP70", adVariant
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
    Set ADOCopyArrayIntoRecordsetCEst = rsADO
    Set rsADO = Nothing
End Function

'*********************************************************
'Special FB code added for WBS 14 3-level sort. 05-19-2020
'*********************************************************

Sub xmlCtrlEst3FB() 'Build 3 level array
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare

    sCat10 = 0
    sCat20 = 0
    sCat30 = 0
    sCat40 = 0
    sCat50 = 0
    sCat51 = 0
    sCat52 = 0
    sCat60 = 0
    sCat61 = 0
    sCat62 = 0
    sCat70 = 0
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
    'Set Level One thru 3
        If nd.SelectSingleNode(sLvl1xNd).text = 0 Then GoTo errNextNd
        Level_FB (nd.SelectSingleNode(sLvl1xNd).text)
        If bCkbAll = True Then
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl1xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl1xNd).text)
        End If
    'Check if UserText6 code applied
        If nd.SelectSingleNode("UserTotal").text > 0 And nd.SelectSingleNode("JobCostCategoryUser").text = "" Then
            CodeChk = "*~*" & nd.SelectSingleNode("Description").text
            bCode = True
        Else
            CodeChk = nd.SelectSingleNode("Description").text
        End If
        For i = 1 To 6
            If i = 1 Then
                dval = nd.SelectSingleNode("LaborTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryLabor").text)
            ElseIf i = 2 Then
                dval = nd.SelectSingleNode("MatTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryMat").text)
            ElseIf i = 3 Then
                dval = nd.SelectSingleNode("SubconTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategorySubcon").text)
            ElseIf i = 4 Then
                dval = nd.SelectSingleNode("EquipTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryEquip").text)
            ElseIf i = 5 Then
                dval = nd.SelectSingleNode("OtherTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryOther").text)
            ElseIf i = 6 Then
                dval = nd.SelectSingleNode("UserTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryUser").text)
            End If
            Select Case nVal
                Case 10: sCat10 = sCat10 + dval
                Case 20: sCat20 = sCat20 + dval
                Case 30: sCat30 = sCat30 + dval
                Case 40: sCat40 = sCat40 + dval
                Case 50: sCat50 = sCat50 + dval
                Case 51: sCat51 = sCat51 + dval
                Case 52: sCat52 = sCat52 + dval
                Case 60: sCat60 = sCat60 + dval
                Case 61: sCat61 = sCat61 + dval
                Case 62: sCat62 = sCat62 + dval
                Case 70: sCat70 = sCat70 + dval
                Case Else
            End Select
            dval = 0
        Next i
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, sCode2, sLvl2, sCode3, sLvl3, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                CodeChk, _
                                nd.SelectSingleNode("ItemNote").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                nd.SelectSingleNode("TakeoffUnit").text, _
                                Val(CDbl(nd.SelectSingleNode("LaborHours").text)), _
                                Val(CDbl(sCat10)), _
                                Val(CDbl(sCat20)), _
                                Val(CDbl(sCat30)), _
                                Val(CDbl(sCat40)), _
                                Val(CDbl(sCat50)), _
                                Val(CDbl(sCat51)), _
                                Val(CDbl(sCat52)), _
                                Val(CDbl(sCat60)), _
                                Val(CDbl(sCat61)), _
                                Val(CDbl(sCat62)), _
                                Val(CDbl(sCat70)), _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)))
        Else
            q = Dict.Item(oTxt)
            q(10) = q(10) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(12) = q(12) + Val(CDbl(nd.SelectSingleNode("LaborHours").text))
            q(13) = q(13) + Val(CDbl(sCat10))
            q(14) = q(14) + Val(CDbl(sCat20))
            q(15) = q(15) + Val(CDbl(sCat30))
            q(16) = q(16) + Val(CDbl(sCat40))
            q(17) = q(17) + Val(CDbl(sCat50))
            q(18) = q(18) + Val(CDbl(sCat51))
            q(19) = q(19) + Val(CDbl(sCat52))
            q(20) = q(20) + Val(CDbl(sCat60))
            q(21) = q(21) + Val(CDbl(sCat61))
            q(22) = q(22) + Val(CDbl(sCat62))
            q(23) = q(23) + Val(CDbl(sCat70))
            q(24) = q(24) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
        sCat10 = 0
        sCat20 = 0
        sCat30 = 0
        sCat40 = 0
        sCat50 = 0
        sCat51 = 0
        sCat52 = 0
        sCat60 = 0
        sCat61 = 0
        sCat62 = 0
        sCat70 = 0
errNextNd:
    Next nd

    ReDim dataArray(1 To Dict.count, 1 To 25)
    On Error Resume Next
    C = 0
    For Each k In Dict.Keys
        C = C + 1
        dataArray(C, 1) = Dict.Item(k)(0) 'Lvl1Code
        dataArray(C, 2) = Dict.Item(k)(1) 'Lvl1Desc
        dataArray(C, 3) = Dict.Item(k)(2) 'Lvl2Code
        dataArray(C, 4) = Dict.Item(k)(3) 'Lvl2Desc
        dataArray(C, 5) = Dict.Item(k)(4) 'Lvl3Code
        dataArray(C, 6) = Dict.Item(k)(5) 'Lvl3Desc
        dataArray(C, 7) = Dict.Item(k)(6) 'SortOrder
        dataArray(C, 8) = Dict.Item(k)(7) 'Index
        dataArray(C, 9) = Dict.Item(k)(8) 'Desc
        dataArray(C, 10) = Dict.Item(k)(9) 'Note
        dataArray(C, 11) = Dict.Item(k)(10) 'TakeoffQty
        dataArray(C, 12) = Dict.Item(k)(11) 'TakeoffUnit
        If Dict.Item(k)(13) > 0 Then dataArray(C, 13) = Dict.Item(k)(12) 'Labor Hours
        dataArray(C, 14) = Dict.Item(k)(13) 'Labor10
        dataArray(C, 15) = Dict.Item(k)(14) 'Material20
        dataArray(C, 16) = Dict.Item(k)(15) 'Equip30
        dataArray(C, 17) = Dict.Item(k)(16) 'Other40
        dataArray(C, 18) = Dict.Item(k)(17) 'Sub50
        dataArray(C, 19) = Dict.Item(k)(18) 'DPR Est 51
        dataArray(C, 20) = Dict.Item(k)(19) 'DPR Cont 52
        dataArray(C, 21) = Dict.Item(k)(20) 'Owner Allow 60
        dataArray(C, 22) = Dict.Item(k)(21) 'ConstCont 61
        dataArray(C, 23) = Dict.Item(k)(22) 'OwnerCont 62
        dataArray(C, 24) = Dict.Item(k)(23) 'OH&P 70
        For i = 14 To 24
            dataArray(C, 25) = dataArray(C, 25) + dataArray(C, i) 'Total
        Next i
    Next k
    On Error GoTo 0
    Set rsNew = ADOCopyArrayIntoRecordsetCEst(argArray:=dataArray)
End Sub

Sub xmlCtrlEst4FB() 'Build 4 level array
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare

    sCat10 = 0
    sCat20 = 0
    sCat30 = 0
    sCat40 = 0
    sCat50 = 0
    sCat51 = 0
    sCat52 = 0
    sCat60 = 0
    sCat61 = 0
    sCat62 = 0
    sCat70 = 0
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
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
        If bCkbAll = True Then
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl2xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) & CInt(nd.SelectSingleNode(sLvl2xNd).text)
        End If
    'Check if UserText6 code applied
        If nd.SelectSingleNode("UserTotal").text > 0 And nd.SelectSingleNode("JobCostCategoryUser").text = "" Then
            CodeChk = "*~*" & nd.SelectSingleNode("Description").text
            bCode = True
        Else
            CodeChk = nd.SelectSingleNode("Description").text
        End If
        For i = 1 To 6
            If i = 1 Then
                dval = nd.SelectSingleNode("LaborTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryLabor").text)
            ElseIf i = 2 Then
                dval = nd.SelectSingleNode("MatTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryMat").text)
            ElseIf i = 3 Then
                dval = nd.SelectSingleNode("SubconTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategorySubcon").text)
            ElseIf i = 4 Then
                dval = nd.SelectSingleNode("EquipTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryEquip").text)
            ElseIf i = 5 Then
                dval = nd.SelectSingleNode("OtherTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryOther").text)
            ElseIf i = 6 Then
                dval = nd.SelectSingleNode("UserTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryUser").text)
            End If
            Select Case nVal
                Case 10: sCat10 = sCat10 + dval
                Case 20: sCat20 = sCat20 + dval
                Case 30: sCat30 = sCat30 + dval
                Case 40: sCat40 = sCat40 + dval
                Case 50: sCat50 = sCat50 + dval
                Case 51: sCat51 = sCat51 + dval
                Case 52: sCat52 = sCat52 + dval
                Case 60: sCat60 = sCat60 + dval
                Case 61: sCat61 = sCat61 + dval
                Case 62: sCat62 = sCat62 + dval
                Case 70: sCat70 = sCat70 + dval
                Case Else
            End Select
            dval = 0
        Next i
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, sCode2, sLvl2, sCode3, sLvl3, sCode4, sLvl4, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                CodeChk, _
                                nd.SelectSingleNode("ItemNote").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                nd.SelectSingleNode("TakeoffUnit").text, _
                                Val(CDbl(nd.SelectSingleNode("LaborHours").text)), _
                                Val(CDbl(sCat10)), _
                                Val(CDbl(sCat20)), _
                                Val(CDbl(sCat30)), _
                                Val(CDbl(sCat40)), _
                                Val(CDbl(sCat50)), _
                                Val(CDbl(sCat51)), _
                                Val(CDbl(sCat52)), _
                                Val(CDbl(sCat60)), _
                                Val(CDbl(sCat61)), _
                                Val(CDbl(sCat62)), _
                                Val(CDbl(sCat70)), _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)))
        Else
            q = Dict.Item(oTxt)
            q(12) = q(12) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(14) = q(14) + Val(CDbl(nd.SelectSingleNode("LaborHours").text))
            q(15) = q(15) + Val(CDbl(sCat10))
            q(16) = q(16) + Val(CDbl(sCat20))
            q(17) = q(17) + Val(CDbl(sCat30))
            q(18) = q(18) + Val(CDbl(sCat40))
            q(19) = q(19) + Val(CDbl(sCat50))
            q(20) = q(20) + Val(CDbl(sCat51))
            q(21) = q(21) + Val(CDbl(sCat52))
            q(22) = q(22) + Val(CDbl(sCat60))
            q(23) = q(23) + Val(CDbl(sCat61))
            q(24) = q(24) + Val(CDbl(sCat62))
            q(25) = q(25) + Val(CDbl(sCat70))
            q(26) = q(26) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
        sCat10 = 0
        sCat20 = 0
        sCat30 = 0
        sCat40 = 0
        sCat50 = 0
        sCat51 = 0
        sCat52 = 0
        sCat60 = 0
        sCat61 = 0
        sCat62 = 0
        sCat70 = 0
errNextNd:
    Next nd

    ReDim dataArray(1 To Dict.count, 1 To 27)
    On Error Resume Next
    C = 0
    For Each k In Dict.Keys
        C = C + 1
        dataArray(C, 1) = Dict.Item(k)(0) 'Lvl1Code
        dataArray(C, 2) = Dict.Item(k)(1) 'Lvl1Desc
        dataArray(C, 3) = Dict.Item(k)(2) 'Lvl2Code
        dataArray(C, 4) = Dict.Item(k)(3) 'Lvl2Desc
        dataArray(C, 5) = Dict.Item(k)(4) 'Lvl3Code
        dataArray(C, 6) = Dict.Item(k)(5) 'Lvl3Desc
        dataArray(C, 7) = Dict.Item(k)(6) 'Lvl4Code
        dataArray(C, 8) = Dict.Item(k)(7) 'Lvl4Desc
        dataArray(C, 9) = Dict.Item(k)(8) 'SortOrder
        dataArray(C, 10) = Dict.Item(k)(9) 'Index
        dataArray(C, 11) = Dict.Item(k)(10) 'Desc
        dataArray(C, 12) = Dict.Item(k)(11) 'Note
        dataArray(C, 13) = Dict.Item(k)(12) 'TakeoffQty
        dataArray(C, 14) = Dict.Item(k)(13) 'TakeoffUnit
        If Dict.Item(k)(15) > 0 Then dataArray(C, 15) = Dict.Item(k)(14) 'Labor Hours
        dataArray(C, 16) = Dict.Item(k)(15) 'Labor10
        dataArray(C, 17) = Dict.Item(k)(16) 'Material20
        dataArray(C, 18) = Dict.Item(k)(17) 'Equip30
        dataArray(C, 19) = Dict.Item(k)(18) 'Other40
        dataArray(C, 20) = Dict.Item(k)(19) 'Sub50
        dataArray(C, 21) = Dict.Item(k)(20) 'DPR Est 51
        dataArray(C, 22) = Dict.Item(k)(21) 'DPR Cont 52
        dataArray(C, 23) = Dict.Item(k)(22) 'Owner Allow 60
        dataArray(C, 24) = Dict.Item(k)(23) 'ConstCont 61
        dataArray(C, 25) = Dict.Item(k)(24) 'OwnerCont 62
        dataArray(C, 26) = Dict.Item(k)(25) 'OH&P 70
        For i = 16 To 26
            dataArray(C, 27) = dataArray(C, 27) + dataArray(C, i) 'Total
        Next i
    Next k
    On Error GoTo 0
    Set rsNew = ADOCopyArrayIntoRecordsetCEst(argArray:=dataArray)
End Sub

Sub xmlCtrlEst5FB() 'Build 5 level array
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare

    sCat10 = 0
    sCat20 = 0
    sCat30 = 0
    sCat40 = 0
    sCat50 = 0
    sCat51 = 0
    sCat52 = 0
    sCat60 = 0
    sCat61 = 0
    sCat62 = 0
    sCat70 = 0
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
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
        If nd.SelectSingleNode(sLvl1xNd).text = 0 Then GoTo errNextNd
        Level_FB3 (nd.SelectSingleNode(sLvl3xNd).text)
        If bCkbAll = True Then
            oTxt = nd.SelectSingleNode("Description").text & nd.SelectSingleNode("ItemNote").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl2xNd).text) & CInt(nd.SelectSingleNode(sLvl3xNd).text)
        Else
            oTxt = nd.SelectSingleNode("Description").text & CInt(nd.SelectSingleNode(sLvl1xNd).text) & CInt(nd.SelectSingleNode(sLvl2xNd).text) _
                   & CInt(nd.SelectSingleNode(sLvl3xNd).text)
        End If
    'Check if UserText6 code applied
        If nd.SelectSingleNode("UserTotal").text > 0 And nd.SelectSingleNode("JobCostCategoryUser").text = "" Then
            CodeChk = "*~*" & nd.SelectSingleNode("Description").text
            bCode = True
        Else
            CodeChk = nd.SelectSingleNode("Description").text
        End If
        For i = 1 To 6
            If i = 1 Then
                dval = nd.SelectSingleNode("LaborTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryLabor").text)
            ElseIf i = 2 Then
                dval = nd.SelectSingleNode("MatTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryMat").text)
            ElseIf i = 3 Then
                dval = nd.SelectSingleNode("SubconTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategorySubcon").text)
            ElseIf i = 4 Then
                dval = nd.SelectSingleNode("EquipTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryEquip").text)
            ElseIf i = 5 Then
                dval = nd.SelectSingleNode("OtherTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryOther").text)
            ElseIf i = 6 Then
                dval = nd.SelectSingleNode("UserTotal").text
                nVal = xNodeVal(nd.SelectSingleNode("JobCostCategoryUser").text)
            End If
            Select Case nVal
                Case 10: sCat10 = sCat10 + dval
                Case 20: sCat20 = sCat20 + dval
                Case 30: sCat30 = sCat30 + dval
                Case 40: sCat40 = sCat40 + dval
                Case 50: sCat50 = sCat50 + dval
                Case 51: sCat51 = sCat51 + dval
                Case 52: sCat52 = sCat52 + dval
                Case 60: sCat60 = sCat60 + dval
                Case 61: sCat61 = sCat61 + dval
                Case 62: sCat62 = sCat62 + dval
                Case 70: sCat70 = sCat70 + dval
                Case Else
            End Select
            dval = 0
        Next i
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, sCode2, sLvl2, _
                                sCode3, sLvl3, sCode4, sLvl4, sCode5, sLvl5, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                CodeChk, _
                                nd.SelectSingleNode("ItemNote").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                nd.SelectSingleNode("TakeoffUnit").text, _
                                Val(CDbl(nd.SelectSingleNode("LaborHours").text)), _
                                Val(CDbl(sCat10)), _
                                Val(CDbl(sCat20)), _
                                Val(CDbl(sCat30)), _
                                Val(CDbl(sCat40)), _
                                Val(CDbl(sCat50)), _
                                Val(CDbl(sCat51)), _
                                Val(CDbl(sCat52)), _
                                Val(CDbl(sCat60)), _
                                Val(CDbl(sCat61)), _
                                Val(CDbl(sCat62)), _
                                Val(CDbl(sCat70)), _
                                Val(CDbl(nd.SelectSingleNode("GrandTotal").text)))
        Else
            q = Dict.Item(oTxt)
            q(14) = q(14) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(16) = q(16) + Val(CDbl(nd.SelectSingleNode("LaborHours").text))
            q(17) = q(17) + Val(CDbl(sCat10))
            q(18) = q(18) + Val(CDbl(sCat20))
            q(19) = q(19) + Val(CDbl(sCat30))
            q(20) = q(20) + Val(CDbl(sCat40))
            q(21) = q(21) + Val(CDbl(sCat50))
            q(22) = q(22) + Val(CDbl(sCat51))
            q(23) = q(23) + Val(CDbl(sCat52))
            q(24) = q(24) + Val(CDbl(sCat60))
            q(25) = q(25) + Val(CDbl(sCat61))
            q(26) = q(26) + Val(CDbl(sCat62))
            q(27) = q(27) + Val(CDbl(sCat70))
            q(28) = q(28) + Val(CDbl(nd.SelectSingleNode("GrandTotal").text))
            Dict.Item(oTxt) = q
        End If
        sCat10 = 0
        sCat20 = 0
        sCat30 = 0
        sCat40 = 0
        sCat50 = 0
        sCat51 = 0
        sCat52 = 0
        sCat60 = 0
        sCat61 = 0
        sCat62 = 0
        sCat70 = 0
errNextNd:
    Next nd

    ReDim dataArray(1 To Dict.count, 1 To 29)
    On Error Resume Next
    C = 0
    For Each k In Dict.Keys
        C = C + 1
        dataArray(C, 1) = Dict.Item(k)(0) 'Lvl1Code
        dataArray(C, 2) = Dict.Item(k)(1) 'Lvl1Desc
        dataArray(C, 3) = Dict.Item(k)(2) 'Lvl2Code
        dataArray(C, 4) = Dict.Item(k)(3) 'Lvl2Desc
        dataArray(C, 5) = Dict.Item(k)(4) 'Lvl3Code
        dataArray(C, 6) = Dict.Item(k)(5) 'Lvl3Desc
        dataArray(C, 7) = Dict.Item(k)(6) 'Lvl4Code
        dataArray(C, 8) = Dict.Item(k)(7) 'Lvl4Desc
        dataArray(C, 9) = Dict.Item(k)(8) 'Lvl5Code
        dataArray(C, 10) = Dict.Item(k)(9) 'Lvl5Desc
        dataArray(C, 11) = Dict.Item(k)(10) 'SortOrder
        dataArray(C, 12) = Dict.Item(k)(11) 'Index
        dataArray(C, 13) = Dict.Item(k)(12) 'Desc
        dataArray(C, 14) = Dict.Item(k)(13) 'Note
        dataArray(C, 15) = Dict.Item(k)(14) 'TakeoffQty
        dataArray(C, 16) = Dict.Item(k)(15) 'TakeoffUnit
        If Dict.Item(k)(17) > 0 Then dataArray(C, 17) = Dict.Item(k)(16) 'Labor Hours
        dataArray(C, 18) = Dict.Item(k)(17) 'Labor10
        dataArray(C, 19) = Dict.Item(k)(18) 'Material20
        dataArray(C, 20) = Dict.Item(k)(19) 'Equip30
        dataArray(C, 21) = Dict.Item(k)(20) 'Other40
        dataArray(C, 22) = Dict.Item(k)(21) 'Sub50
        dataArray(C, 23) = Dict.Item(k)(22) 'DPR Est 51
        dataArray(C, 24) = Dict.Item(k)(23) 'DPR Cont 52
        dataArray(C, 25) = Dict.Item(k)(24) 'Owner Allow 60
        dataArray(C, 26) = Dict.Item(k)(25) 'ConstCont 61
        dataArray(C, 27) = Dict.Item(k)(26) 'OwnerCont 62
        dataArray(C, 28) = Dict.Item(k)(27) 'OH&P 70
        For i = 18 To 28
            dataArray(C, 29) = dataArray(C, 29) + dataArray(C, i) 'Total
        Next i
    Next k
    On Error GoTo 0
    Set rsNew = ADOCopyArrayIntoRecordsetCEst(argArray:=dataArray)
End Sub


