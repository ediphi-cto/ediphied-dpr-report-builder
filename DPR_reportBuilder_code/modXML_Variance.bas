Attribute VB_Name = "modXML_Variance"
Dim dVar As Double
Dim dQVar As Double
Public sTotal As String
Public bMarkups As Boolean
Public xLvlVar As Variant
Public sDoc As String


Sub xml_VAR_Level1() 'Build 1 level Variance
    Set owb = ActiveWorkbook
    wbName = owb.Name

'Load Estimate 1 XML file
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    

    On Error Resume Next
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    sDoc = 1
    For Each nd In xNode
        Level_One (nd.SelectSingleNode(sLvl1xNd).text)
'        oTxt = nd.SelectSingleNode("Description").Text & CInt(nd.SelectSingleNode(sLvl1xNd).Text)
        oTxt = nd.SelectSingleNode("Description").text & "-" & sCode1 & "-" & sLvl1
'        sVarID = nd.SelectSingleNode("Identity").Text
        
    'Add to data collection
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                nd.SelectSingleNode("Description").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                Val(CDbl(nd.SelectSingleNode("" & sTotal & "").text)), _
                                dQVar, dVar, nd.SelectSingleNode("TakeoffUnit").text, "")
        Else
            q = Dict.Item(oTxt)
            q(5) = q(5) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(6) = q(6) + Val(CDbl(nd.SelectSingleNode("" & sTotal & "").text))
            q(7) = q(7) + dQVar
            q(8) = q(8) + dVar
            Dict.Item(oTxt) = q
        End If
    Next nd
'Scan through Estimate 2 data and add if missing from Estimate 1

'Load Estimate 2 XML file
    XDocV.async = False
    XDocV.validateOnParse = False
    XDocV.Load xmlVarPath

    Set xLvlVar = XDocV.SelectNodes(sXpath)
    sDoc = 2
    dQVar = 0
    dVar = 0
    For Each ndvar In xLvlVar
        Level_One (ndvar.SelectSingleNode(sLvl1xNd).text) 'Check WBS Table for code/Name
        oTxt = ndvar.SelectSingleNode("Description").text & "-" & sCode1 & "-" & sLvl1
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                ndvar.SelectSingleNode("SortOrder").text, _
                                Format(ndvar.SelectSingleNode("Index").text, "0000") & ndvar.SelectSingleNode("ItemCode").text, _
                                ndvar.SelectSingleNode("Description").text, _
                                dQVar, dVar, _
                                Val(CDbl(ndvar.SelectSingleNode("TakeoffQty").text)), _
                                Val(CDbl(ndvar.SelectSingleNode("" & sTotal & "").text)), "", ndvar.SelectSingleNode("TakeoffUnit").text)
        Else
            q = Dict.Item(oTxt)
            q(5) = q(5) + dQVar
            q(6) = q(6) + dVar
            q(7) = q(7) + Val(CDbl(ndvar.SelectSingleNode("TakeoffQty").text))
            q(8) = q(8) + Val(CDbl(ndvar.SelectSingleNode("" & sTotal & "").text))
            Dict.Item(oTxt) = q
        End If
    Next ndvar
    On Error GoTo 0
    If Dict.count > 0 Then
        ReDim dataArray(1 To Dict.count, 1 To 16)
        On Error Resume Next
        C = 0
        For Each k In Dict.Keys
            C = C + 1
            dataArray(C, 1) = Dict.Item(k)(0)                       'Level 1 Code
            dataArray(C, 2) = Dict.Item(k)(1)                       'Level 1 Desc
            dataArray(C, 3) = Dict.Item(k)(2)                       'Sort Order
            dataArray(C, 4) = Dict.Item(k)(3)                       'Item Code
            dataArray(C, 5) = Dict.Item(k)(4)                       'Description
            dataArray(C, 6) = Dict.Item(k)(5)                       'Est 1 Qty
            dataArray(C, 7) = Dict.Item(k)(6) / Dict.Item(k)(5)     'Est 1 Unit
            dataArray(C, 8) = Dict.Item(k)(6)                       'Est 1 Total
            dataArray(C, 9) = Dict.Item(k)(7)                       'Est 2 Qty
            dataArray(C, 10) = Dict.Item(k)(8) / Dict.Item(k)(7)    'Est 2 Unit
            dataArray(C, 11) = Dict.Item(k)(8)                      'Est 2 Total
            dataArray(C, 12) = Dict.Item(k)(7) - Dict.Item(k)(5)    'Qty Var
            dataArray(C, 13) = dataArray(C, 10) - dataArray(C, 7)   'Unit Var
            dataArray(C, 14) = Dict.Item(k)(8) - Dict.Item(k)(6)    'Value Var
            dataArray(C, 15) = Dict.Item(k)(9)                      'EST1 U/M
            If Dict.Item(k)(10) = "" Then
                dataArray(C, 16) = Dict.Item(k)(9)                  'EST2 U/M
            Else
                dataArray(C, 16) = Dict.Item(k)(10)                 'EST2 U/M
            End If
        Next k
        On Error GoTo 0
        
        Set rsNew = ADOCopyArrayIntoRecordset_VAR(argArray:=dataArray)
        If bRefresh = False Then
            Call Create_PivotTable_ODBC_MO_VAR
            MsgBox "Report Complete", vbOKOnly, "DPR Report Builder"
        End If
    Else
        MsgBox "Unable to create this report. Make sure the WBS codes you have selected have values associated.", vbCritical, "Unable To Create Report"
    End If
End Sub

Sub xml_VAR_Level2() 'Build 2 level variance
    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    
    On Error Resume Next
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    sDoc = 1
    For Each nd In xNode
        Level_One (nd.SelectSingleNode(sLvl1xNd).text)
        Level_Two (nd.SelectSingleNode(sLvl2xNd).text)
        oTxt = nd.SelectSingleNode("Description").text & sCode1 & sLvl1 & sCode2 & sLvl2
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                nd.SelectSingleNode("Description").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                Val(CDbl(nd.SelectSingleNode("" & sTotal & "").text)), _
                                dQVar, dVar, nd.SelectSingleNode("TakeoffUnit").text, "")
        Else
            q = Dict.Item(oTxt)
            q(7) = q(7) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(8) = q(8) + Val(CDbl(nd.SelectSingleNode("" & sTotal & "").text))
            q(9) = q(9) + dQVar
            q(10) = q(10) + dVar
            Dict.Item(oTxt) = q
        End If
    Next nd
'Scan through Estimate 2 data and add if missing from Estimate 1
    'Load Estimate 2 XML file
    XDocV.async = False
    XDocV.validateOnParse = False
    XDocV.Load xmlVarPath
    Set xLvlVar = XDocV.SelectNodes(sXpath)
    sDoc = 2
    dVar = 0
    For Each ndvar In xLvlVar
        Level_One (ndvar.SelectSingleNode(sLvl1xNd).text)
        Level_Two (ndvar.SelectSingleNode(sLvl2xNd).text)
        oTxt = ndvar.SelectSingleNode("Description").text & sCode1 & sLvl1 & sCode2 & sLvl2
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                ndvar.SelectSingleNode("SortOrder").text, _
                                Format(ndvar.SelectSingleNode("Index").text, "0000") & ndvar.SelectSingleNode("ItemCode").text, _
                                ndvar.SelectSingleNode("Description").text, _
                                dQVar, dVar, _
                                Val(CDbl(ndvar.SelectSingleNode("TakeoffQty").text)), _
                                Val(CDbl(ndvar.SelectSingleNode("" & sTotal & "").text)), "", ndvar.SelectSingleNode("TakeoffUnit").text)
        Else
            q = Dict.Item(oTxt)
            q(7) = q(7) + dQVar
            q(8) = q(8) + dVar
            q(9) = q(9) + Val(CDbl(ndvar.SelectSingleNode("TakeoffQty").text))
            q(10) = q(10) + Val(CDbl(ndvar.SelectSingleNode("" & sTotal & "").text))
            Dict.Item(oTxt) = q
        End If
    Next ndvar
    On Error GoTo 0
    If Dict.count > 0 Then
        ReDim dataArray(1 To Dict.count, 1 To 18)
        On Error Resume Next
        C = 0
        For Each k In Dict.Keys
            C = C + 1
            dataArray(C, 1) = Dict.Item(k)(0)                       'Lvl 1 Code
            dataArray(C, 2) = Dict.Item(k)(1)                       'Lvl 1 Desc
            dataArray(C, 3) = Dict.Item(k)(2)                       'Lvl 2 Code
            dataArray(C, 4) = Dict.Item(k)(3)                       'lvl 2 Desc
            dataArray(C, 5) = Dict.Item(k)(4)                       'Sort Order
            dataArray(C, 6) = Dict.Item(k)(5)                       'Item Code
            dataArray(C, 7) = Dict.Item(k)(6)                       'Description
            dataArray(C, 8) = Dict.Item(k)(7)                       'Est 1 Qty
            dataArray(C, 9) = Dict.Item(k)(8) / Dict.Item(k)(7)     'Est 1 Unit
            dataArray(C, 10) = Dict.Item(k)(8)                      'Est 1 Total
            dataArray(C, 11) = Dict.Item(k)(9)                      'Est 2 Qty
            dataArray(C, 12) = Dict.Item(k)(10) / Dict.Item(k)(9)   'Est 2 Unit
            dataArray(C, 13) = Dict.Item(k)(10)                     'Est 2 Total
            dataArray(C, 14) = Dict.Item(k)(9) - Dict.Item(k)(7)    'Qty Var
            dataArray(C, 15) = dataArray(C, 12) - dataArray(C, 9)   'Unit Var
            dataArray(C, 16) = Dict.Item(k)(10) - Dict.Item(k)(8)   'Value Var
            dataArray(C, 17) = Dict.Item(k)(11)                     'EST1 U/M
            If Dict.Item(k)(12) = "" Then
                dataArray(C, 18) = Dict.Item(k)(11)                 'EST2 U/M
            Else
                dataArray(C, 18) = Dict.Item(k)(12)                 'EST2 U/M
            End If
        Next k
        On Error GoTo 0
        Set rsNew = ADOCopyArrayIntoRecordset_VAR(argArray:=dataArray)
        If bRefresh = False Then
            Call Create_PivotTable_ODBC_MO_VAR
            MsgBox "Report Complete", vbOKOnly, "DPR Report Builder"
        End If
    Else
        MsgBox "Unable to create this report. Make sure the WBS codes you have selected have values associated.", vbCritical, "Unable To Create Report"
    End If
End Sub

Sub xml_VAR_Level3() 'Build 3 level variance
    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare

    On Error Resume Next
    sDoc = 1
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
        Level_One (nd.SelectSingleNode(sLvl1xNd).text)
        Level_Two (nd.SelectSingleNode(sLvl2xNd).text)
        Level_Three (nd.SelectSingleNode(sLvl3xNd).text)
        oTxt = nd.SelectSingleNode("Description").text & sCode1 & sLvl1 & sCode2 & sLvl2 & sCode3 & sLvl3
 
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                sCode3, sLvl3, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                nd.SelectSingleNode("Description").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                Val(CDbl(nd.SelectSingleNode("" & sTotal & "").text)), _
                                dQVar, dVar, nd.SelectSingleNode("TakeoffUnit").text, "")
        Else
            q = Dict.Item(oTxt)
            q(9) = q(9) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(10) = q(10) + Val(CDbl(nd.SelectSingleNode("" & sTotal & "").text))
            q(11) = q(11) + dQVar
            q(12) = q(12) + dVar
            Dict.Item(oTxt) = q
        End If
    Next nd
'Scan through Estimate 2 data and add if missing from Estimate 1
'Load Estimate 2 XML file
    XDocV.async = False
    XDocV.validateOnParse = False
    XDocV.Load xmlVarPath
    Set xLvlVar = XDocV.SelectNodes(sXpath)
    dVar = 0
    sDoc = 2
    For Each ndvar In xLvlVar
        Level_One (ndvar.SelectSingleNode(sLvl1xNd).text)
        Level_Two (ndvar.SelectSingleNode(sLvl2xNd).text)
        Level_Three (ndvar.SelectSingleNode(sLvl3xNd).text)
        oTxt = ndvar.SelectSingleNode("Description").text & sCode1 & sLvl1 & sCode2 & sLvl2 & sCode3 & sLvl3
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                sCode3, sLvl3, _
                                ndvar.SelectSingleNode("SortOrder").text, _
                                Format(ndvar.SelectSingleNode("Index").text, "0000") & ndvar.SelectSingleNode("ItemCode").text, _
                                ndvar.SelectSingleNode("Description").text, _
                                dQVar, dVar, _
                                Val(CDbl(ndvar.SelectSingleNode("TakeoffQty").text)), _
                                Val(CDbl(ndvar.SelectSingleNode("" & sTotal & "").text)), "", ndvar.SelectSingleNode("TakeoffUnit").text)
        Else
            q = Dict.Item(oTxt)
            q(9) = q(9) + dQVar
            q(10) = q(10) + dVar
            q(11) = q(11) + Val(CDbl(ndvar.SelectSingleNode("TakeoffQty").text))
            q(12) = q(12) + Val(CDbl(ndvar.SelectSingleNode("" & sTotal & "").text))
            Dict.Item(oTxt) = q
        End If
    Next ndvar
    On Error GoTo 0
    If Dict.count > 0 Then
        ReDim dataArray(1 To Dict.count, 1 To 20)
        On Error Resume Next
        C = 0
        For Each k In Dict.Keys
            C = C + 1
            dataArray(C, 1) = Dict.Item(k)(0)                       'Lvl 1 Code
            dataArray(C, 2) = Dict.Item(k)(1)                       'Lvl 1 Desc
            dataArray(C, 3) = Dict.Item(k)(2)                       'Lvl 2 Code
            dataArray(C, 4) = Dict.Item(k)(3)                       'lvl 2 Desc
            dataArray(C, 5) = Dict.Item(k)(4)                       'Lvl 3 Code
            dataArray(C, 6) = Dict.Item(k)(5)                       'lvl 3 Desc
            dataArray(C, 7) = Dict.Item(k)(6)                       'Sort Order
            dataArray(C, 8) = Dict.Item(k)(7)                       'Item Code
            dataArray(C, 9) = Dict.Item(k)(8)                       'Description
            dataArray(C, 10) = Dict.Item(k)(9)                      'Est 1 Qty
            dataArray(C, 11) = Dict.Item(k)(10) / Dict.Item(k)(9)   'Est 1 Unit
            dataArray(C, 12) = Dict.Item(k)(10)                     'Est 1 Total
            dataArray(C, 13) = Dict.Item(k)(11)                     'Est 2 Qty
            dataArray(C, 14) = Dict.Item(k)(12) / Dict.Item(k)(11)  'Est 2 Unit
            dataArray(C, 15) = Dict.Item(k)(12)                     'Est 2 Total
            dataArray(C, 16) = Dict.Item(k)(11) - Dict.Item(k)(9)   'Qty Var
            dataArray(C, 17) = dataArray(C, 14) - dataArray(C, 11)  'Unit Var
            dataArray(C, 18) = Dict.Item(k)(12) - Dict.Item(k)(10)  'Value Var
            dataArray(C, 19) = Dict.Item(k)(13)                     'EST1 U/M
            If Dict.Item(k)(14) = "" Then
                dataArray(C, 20) = Dict.Item(k)(13)                 'EST2 U/M
            Else
                dataArray(C, 20) = Dict.Item(k)(14)                 'EST2 U/M
            End If
        Next k
        On Error GoTo 0
        Set rsNew = ADOCopyArrayIntoRecordset_VAR(argArray:=dataArray)
        If bRefresh = False Then
            Call Create_PivotTable_ODBC_MO_VAR
            MsgBox "Report Complete", vbOKOnly, "DPR Report Builder"
        End If
    Else
        MsgBox "Unable to create this report. Make sure the WBS codes you have selected have values associated.", vbCritical, "Unable To Create Report"
    End If
End Sub

Sub xml_VAR_Level4() 'Build 4 level variance
    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    

    On Error Resume Next
    sDoc = 1
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
        Level_One (nd.SelectSingleNode(sLvl1xNd).text)
        Level_Two (nd.SelectSingleNode(sLvl2xNd).text)
        Level_Three (nd.SelectSingleNode(sLvl3xNd).text)
        Level_Four (nd.SelectSingleNode(sLvl4xNd).text)
        oTxt = nd.SelectSingleNode("Description").text & sCode1 & sLvl1 & sCode2 & sLvl2 & sCode3 & sLvl3 & sCode4 & sLvl4
 
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                sCode3, sLvl3, _
                                sCode4, sLvl4, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                nd.SelectSingleNode("Description").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                Val(CDbl(nd.SelectSingleNode("" & sTotal & "").text)), _
                                dQVar, dVar, nd.SelectSingleNode("TakeoffUnit").text, "")
        Else
            q = Dict.Item(oTxt)
            q(11) = q(11) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(12) = q(12) + Val(CDbl(nd.SelectSingleNode("" & sTotal & "").text))
            q(13) = q(13) + dQVar
            q(14) = q(14) + dVar
            Dict.Item(oTxt) = q
        End If
    Next nd
'Scan through Estimate 2 data and add if missing from Estimate 1
    'Load Estimate 2 XML file
    XDocV.async = False
    XDocV.validateOnParse = False
    XDocV.Load xmlVarPath
    Set xLvlVar = XDocV.SelectNodes(sXpath)
    sDoc = 2
    dVar = 0
    For Each ndvar In xLvlVar
        Level_One (ndvar.SelectSingleNode(sLvl1xNd).text)
        Level_Two (ndvar.SelectSingleNode(sLvl2xNd).text)
        Level_Three (ndvar.SelectSingleNode(sLvl3xNd).text)
        Level_Four (ndvar.SelectSingleNode(sLvl4xNd).text)
        oTxt = ndvar.SelectSingleNode("Description").text & sCode1 & sLvl1 & sCode2 & sLvl2 & sCode3 & sLvl3 & sCode4 & sLvl4
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                sCode3, sLvl3, _
                                sCode4, sLvl4, _
                                ndvar.SelectSingleNode("SortOrder").text, _
                                Format(ndvar.SelectSingleNode("Index").text, "0000") & ndvar.SelectSingleNode("ItemCode").text, _
                                ndvar.SelectSingleNode("Description").text, _
                                dQVar, dVar, _
                                Val(CDbl(ndvar.SelectSingleNode("TakeoffQty").text)), _
                                Val(CDbl(ndvar.SelectSingleNode("" & sTotal & "").text)), "", ndvar.SelectSingleNode("TakeoffUnit").text)
        Else
            q = Dict.Item(oTxt)
            q(11) = q(11) + dQVar
            q(12) = q(12) + dVar
            q(13) = q(13) + Val(CDbl(ndvar.SelectSingleNode("TakeoffQty").text))
            q(14) = q(14) + Val(CDbl(ndvar.SelectSingleNode("" & sTotal & "").text))
            Dict.Item(oTxt) = q
        End If
    Next ndvar
    On Error GoTo 0
    If Dict.count > 0 Then
        ReDim dataArray(1 To Dict.count, 1 To 22)
        On Error Resume Next
        C = 0
        For Each k In Dict.Keys
            C = C + 1
            dataArray(C, 1) = Dict.Item(k)(0)                           'Lvl 1 Code
            dataArray(C, 2) = Dict.Item(k)(1)                           'Lvl 1 Desc
            dataArray(C, 3) = Dict.Item(k)(2)                           'Lvl 2 Code
            dataArray(C, 4) = Dict.Item(k)(3)                           'lvl 2 Desc
            dataArray(C, 5) = Dict.Item(k)(4)                           'Lvl 3 Code
            dataArray(C, 6) = Dict.Item(k)(5)                           'lvl 3 Desc
            dataArray(C, 7) = Dict.Item(k)(6)                           'Lvl 4 Code
            dataArray(C, 8) = Dict.Item(k)(7)                           'lvl 4 Desc
            dataArray(C, 9) = Dict.Item(k)(8)                           'Sort Order
            dataArray(C, 10) = Dict.Item(k)(9)                          'Item Code
            dataArray(C, 11) = Dict.Item(k)(10)                         'Description
            dataArray(C, 12) = Dict.Item(k)(11)                         'Est 1 Qty
            dataArray(C, 13) = Dict.Item(k)(12) / Dict.Item(k)(11)      'Est 1 Unit
            dataArray(C, 14) = Dict.Item(k)(12)                         'Est 1 Total
            dataArray(C, 15) = Dict.Item(k)(13)                         'Est 2 Qty
            dataArray(C, 16) = Dict.Item(k)(14) / Dict.Item(k)(13)      'Est 2 Unit
            dataArray(C, 17) = Dict.Item(k)(14)                         'Est 2 Total
            dataArray(C, 18) = Dict.Item(k)(13) - Dict.Item(k)(11)      'Qty Var
            dataArray(C, 19) = dataArray(C, 16) - dataArray(C, 13)      'Unit Var
            dataArray(C, 20) = Dict.Item(k)(14) - Dict.Item(k)(12)      'Value Var
            dataArray(C, 21) = Dict.Item(k)(15)                         'EST1 U/M
            If Dict.Item(k)(16) = "" Then
                dataArray(C, 22) = Dict.Item(k)(15)                     'EST2 U/M
            Else
                dataArray(C, 22) = Dict.Item(k)(16)                     'EST2 U/M
            End If
        Next k
        On Error GoTo 0
        Set rsNew = ADOCopyArrayIntoRecordset_VAR(argArray:=dataArray)
        If bRefresh = False Then
            Call Create_PivotTable_ODBC_MO_VAR
            MsgBox "Report Complete", vbOKOnly, "DPR Report Builder"
        End If
    Else
        MsgBox "Unable to create this report. Make sure the WBS codes you have selected have values associated.", vbCritical, "Unable To Create Report"
    End If
End Sub

Sub xml_VAR_Level5() 'Build 5 level variance
    Set owb = ActiveWorkbook
    wbName = owb.Name
    XDoc.async = False
    XDoc.validateOnParse = False
    XDoc.Load xmlPath
    pth = owb.Path & ".xlsm"
    Set Dict = CreateObject("scripting.dictionary")
    Dict.CompareMode = vbTextCompare
    

    On Error Resume Next
    sDoc = 1
    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
    For Each nd In xNode
        Level_One (nd.SelectSingleNode(sLvl1xNd).text)
        Level_Two (nd.SelectSingleNode(sLvl2xNd).text)
        Level_Three (nd.SelectSingleNode(sLvl3xNd).text)
        Level_Four (nd.SelectSingleNode(sLvl4xNd).text)
        Level_Five (nd.SelectSingleNode(sLvl5xNd).text)
        oTxt = nd.SelectSingleNode("Description").text & sCode1 & sLvl1 & sCode2 & sLvl2 & sCode3 & sLvl3 & sCode4 & sLvl4 & sCode5 & sLvl5
 
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                sCode3, sLvl3, _
                                sCode4, sLvl4, _
                                sCode5, sLvl5, _
                                nd.SelectSingleNode("SortOrder").text, _
                                Format(nd.SelectSingleNode("Index").text, "0000") & nd.SelectSingleNode("ItemCode").text, _
                                nd.SelectSingleNode("Description").text, _
                                Val(CDbl(nd.SelectSingleNode("TakeoffQty").text)), _
                                Val(CDbl(nd.SelectSingleNode("" & sTotal & "").text)), _
                                dQVar, dVar, nd.SelectSingleNode("TakeoffUnit").text, "")
        Else
            q = Dict.Item(oTxt)
            q(13) = q(13) + Val(CDbl(nd.SelectSingleNode("TakeoffQty").text))
            q(14) = q(14) + Val(CDbl(nd.SelectSingleNode("" & sTotal & "").text))
            q(15) = q(15) + dQVar
            q(16) = q(16) + dVar
            Dict.Item(oTxt) = q
        End If
    Next nd
'Scan through Estimate 2 data and add if missing from Estimate 1
    'Load Estimate 2 XML file
    XDocV.async = False
    XDocV.validateOnParse = False
    XDocV.Load xmlVarPath
    Set xLvlVar = XDocV.SelectNodes(sXpath)
    dVar = 0
    sDoc = 2
    For Each ndvar In xLvlVar
        Level_One (ndvar.SelectSingleNode(sLvl1xNd).text)
        Level_Two (ndvar.SelectSingleNode(sLvl2xNd).text)
        Level_Three (ndvar.SelectSingleNode(sLvl3xNd).text)
        Level_Four (ndvar.SelectSingleNode(sLvl4xNd).text)
        Level_Five (ndvar.SelectSingleNode(sLvl5xNd).text)
        oTxt = ndvar.SelectSingleNode("Description").text & sCode1 & sLvl1 & sCode2 & sLvl2 & sCode3 & sLvl3 & sCode4 & sLvl4 & sCode5 & sLvl5
        If Not Dict.Exists(oTxt) Then
            Dict.Add oTxt, Array(sCode1, sLvl1, _
                                sCode2, sLvl2, _
                                sCode3, sLvl3, _
                                sCode4, sLvl4, _
                                sCode5, sLvl5, _
                                ndvar.SelectSingleNode("SortOrder").text, _
                                Format(ndvar.SelectSingleNode("Index").text, "0000") & ndvar.SelectSingleNode("ItemCode").text, _
                                ndvar.SelectSingleNode("Description").text, _
                                dQVar, dVar, _
                                Val(CDbl(ndvar.SelectSingleNode("TakeoffQty").text)), _
                                Val(CDbl(ndvar.SelectSingleNode("" & sTotal & "").text)), "", ndvar.SelectSingleNode("TakeoffUnit").text)
        Else
            q = Dict.Item(oTxt)
            q(13) = q(13) + dQVar
            q(14) = q(14) + dVar
            q(15) = q(15) + Val(CDbl(ndvar.SelectSingleNode("TakeoffQty").text))
            q(16) = q(16) + Val(CDbl(ndvar.SelectSingleNode("" & sTotal & "").text))
            Dict.Item(oTxt) = q
        End If
    Next ndvar
    On Error GoTo 0
    If Dict.count > 0 Then
        ReDim dataArray(1 To Dict.count, 1 To 24)
        On Error Resume Next
        C = 0
        For Each k In Dict.Keys
            C = C + 1
            dataArray(C, 1) = Dict.Item(k)(0)                           'Lvl 1 Code
            dataArray(C, 2) = Dict.Item(k)(1)                           'Lvl 1 Desc
            dataArray(C, 3) = Dict.Item(k)(2)                           'Lvl 2 Code
            dataArray(C, 4) = Dict.Item(k)(3)                           'lvl 2 Desc
            dataArray(C, 5) = Dict.Item(k)(4)                           'Lvl 3 Code
            dataArray(C, 6) = Dict.Item(k)(5)                           'lvl 3 Desc
            dataArray(C, 7) = Dict.Item(k)(6)                           'Lvl 4 Code
            dataArray(C, 8) = Dict.Item(k)(7)                           'lvl 4 Desc
            dataArray(C, 9) = Dict.Item(k)(8)                           'Lvl 5 Code
            dataArray(C, 10) = Dict.Item(k)(9)                          'lvl 5 Desc
            dataArray(C, 11) = Dict.Item(k)(10)                         'Sort Order
            dataArray(C, 12) = Dict.Item(k)(11)                         'Item Code
            dataArray(C, 13) = Dict.Item(k)(12)                         'Description
            dataArray(C, 14) = Dict.Item(k)(13)                         'Est 1 Qty
            dataArray(C, 15) = Dict.Item(k)(14) / Dict.Item(k)(13)      'Est 1 Unit
            dataArray(C, 16) = Dict.Item(k)(14)                         'Est 1 Total
            dataArray(C, 17) = Dict.Item(k)(15)                         'Est 2 Qty
            dataArray(C, 18) = Dict.Item(k)(16) / Dict.Item(k)(15)      'Est 2 Unit
            dataArray(C, 19) = Dict.Item(k)(16)                         'Est 2 Total
            dataArray(C, 20) = Dict.Item(k)(15) - Dict.Item(k)(13)      'Qty Var
            dataArray(C, 21) = dataArray(C, 18) - dataArray(C, 15)      'Unit Var
            dataArray(C, 22) = Dict.Item(k)(16) - Dict.Item(k)(14)      'Value Var
            dataArray(C, 23) = Dict.Item(k)(17)                         'EST1 U/M
            If Dict.Item(k)(18) = "" Then
                dataArray(C, 24) = Dict.Item(k)(17)                     'EST2 U/M
            Else
                dataArray(C, 24) = Dict.Item(k)(18)                     'EST2 U/M
            End If
        Next k
        On Error GoTo 0
        Set rsNew = ADOCopyArrayIntoRecordset_VAR(argArray:=dataArray)
        If bRefresh = False Then
            Call Create_PivotTable_ODBC_MO_VAR
            MsgBox "Report Complete", vbOKOnly, "DPR Report Builder"
        End If
    Else
        MsgBox "Unable to create this report. Make sure the WBS codes you have selected have values associated.", vbCritical, "Unable To Create Report"
    End If
End Sub

Private Function ADOCopyArrayIntoRecordset_VAR(argArray As Variant) As ADODB.Recordset
'Create data recordset for pivot cache
'MN: Replace this with our generation of "rsnew" to feed into a pivotcache
Dim cnnConn As ADODB.Connection
Dim rsADO As ADODB.Recordset
Dim cmdCommand As ADODB.Command
Dim lngR As Long
Dim lngC As Long

    Set rsADO = New ADODB.Recordset
    For i = 1 To iLvl
        Select Case i
            Case 1
                sLvl1Code = sLvl1Code & i
                rsADO.Fields.Append sLvl1Code, adVariant '
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
    rsADO.Fields.Append "Est1_Qty", adVariant
    rsADO.Fields.Append "Est1_Unit", adVariant
    rsADO.Fields.Append "Est1_Total", adVariant
    rsADO.Fields.Append "Est2_Qty", adVariant
    rsADO.Fields.Append "Est2_Unit", adVariant
    rsADO.Fields.Append "Est2_Total", adVariant
    rsADO.Fields.Append "VarQty", adVariant
    rsADO.Fields.Append "VarUnit", adVariant
    rsADO.Fields.Append "VarTotal", adVariant
    rsADO.Fields.Append "Est1_UM", adVariant
    rsADO.Fields.Append "Est2_UM", adVariant
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
    Set ADOCopyArrayIntoRecordset_VAR = rsADO
End Function

Function Level_One(sIndex As String)
    If sDoc = 1 Then
        Set xLvl1 = XDoc.SelectNodes(sXpath1 & "[Index=" & sIndex & "]")
    Else
        Set xLvl1 = XDocV.SelectNodes(sXpath1 & "[Index=" & sIndex & "]")
    End If
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
End Function

Function Level_Two(sIndex As String)
    If sDoc = 1 Then
        Set xLvl2 = XDoc.SelectNodes(sXpath2 & "[Index=" & sIndex & "]")
    Else
        Set xLvl2 = XDocV.SelectNodes(sXpath2 & "[Index=" & sIndex & "]")
    End If
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
End Function

Function Level_Three(sIndex As String)
    If sDoc = 1 Then
        Set xLvl3 = XDoc.SelectNodes(sXpath3 & "[Index=" & sIndex & "]")
    Else
        Set xLvl3 = XDocV.SelectNodes(sXpath3 & "[Index=" & sIndex & "]")
    End If
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
End Function

Function Level_Four(sIndex As String)
    If sDoc = 1 Then
        Set xLvl4 = XDoc.SelectNodes(sXpath4 & "[Index=" & sIndex & "]")
    Else
        Set xLvl4 = XDocV.SelectNodes(sXpath4 & "[Index=" & sIndex & "]")
    End If
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
End Function

Function Level_Five(sIndex As String)
    If sDoc = 1 Then
        Set xLvl5 = XDoc.SelectNodes(sXpath5 & "[Index=" & sIndex & "]")
    Else
        Set xLvl5 = XDocV.SelectNodes(sXpath5 & "[Index=" & sIndex & "]")
    End If
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
End Function

