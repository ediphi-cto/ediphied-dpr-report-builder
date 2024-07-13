Attribute VB_Name = "modLaborHr"
Public dLRate As Double
Public SCode As String
'Sub LaborHrProjection() 'Load Labor Hour Projection sheet
'Dim oTxt        As String
'Dim q           As Variant
'Dim dict         As Object
'Dim k           As Variant
'Dim C           As Long
'Dim pth         As String
'Dim X
'Dim XDoc As MSXML2.DOMDocument60
'
'
'    Set owb = ActiveWorkbook
'    wbName = owb.Name
''    Set XDoc = New MSXML2.DOMDocument60
''    XDoc.async = False
''    XDoc.validateOnParse = False
''    XDoc.Load xmlPath
''    pth = owb.Path & ".xlsm"
''
''    sXpath1 = "/Estimate/CategoryTable/Category"
'    sLvl1Code = "Code"
'    sLvl1xNd = "Category"
'
'
'    Set dict = CreateObject("scripting.dictionary")
'    dict.CompareMode = vbTextCompare
'
'    Set xNode = XDoc.SelectNodes(sXpath)    'Item Level
'    Set xLvl1 = XDoc.SelectNodes(sXpath1)   'Bid Package Level
'
'
''Set Labor Hr Projection Tables
'    Dim TradeArr() As Variant
'    Dim rng As Range, cel As Range
'    Dim dPerc As Double
'
''****Load Percentages to Array
'    Set rng = Sheet0.ListObjects("tblTradeCat").DataBodyRange
'    TradeArr = rng
''*****
'
'    On Error Resume Next
'    For Each nd1 In xLvl1
'        sIndex = nd1.SelectSingleNode("Index").text
'        sLvl1 = nd1.SelectSingleNode(sLvl1Code).text & "  " & nd1.SelectSingleNode("Name").text
'    'Find Labor Percentage from TradeCat table
'        For X = LBound(TradeArr) To UBound(TradeArr)
'            If TradeArr(X, 1) = nd1.SelectSingleNode(sLvl1Code).text Then
'                dPerc = TradeArr(X, 3)
'                sCode = TradeArr(X, 5)
'                'Search for rate in Labor Table
'                Set xLvl2 = XDoc.SelectNodes("/Estimate/LaborTable/Labor[Trade= '" & sCode & "']")
'                    For Each nd2 In xLvl2
'                        dLRate = nd2.SelectSingleNode("Rate").text
'                        Exit For
'                    Next nd2
'                Exit For
'            Else
'                dPerc = 0.5
'                sCode = ""
'            End If
'        Next X
'
'        If sLvl1 <> "(none)" Then
'            For Each nd In xNode
'                If CInt(nd.SelectSingleNode(sLvl1xNd).text) = CInt(nd1.SelectSingleNode("Index").text) Then
'                    oTxt = sLvl1
'                    If Not dict.Exists(oTxt) Then
'                        dict.Add oTxt, Array(sIndex, sLvl1, val(CDbl(nd.SelectSingleNode("GrandTotal").text)), dLRate, dPerc)
'                    Else
'                        q = dict.Item(oTxt)
'                        q(2) = q(2) + val(CDbl(nd.SelectSingleNode("GrandTotal").text))
'                        dict.Item(oTxt) = q
'                    End If
'                End If
'            Next nd
'        End If
'        'reset rate
'        dLRate = 0
'    Next nd1
'    ReDim dataArray(1 To dict.count, 1 To 4)
'        For Each k In dict.Keys
'            C = C + 1
'            dataArray(C, 1) = dict.Item(k)(1)
'            dataArray(C, 2) = dict.Item(k)(2)
'            dataArray(C, 3) = dict.Item(k)(3)
'            dataArray(C, 4) = dict.Item(k)(4)
'        Next k
'    On Error GoTo 0
'
'    Set ows = Sheet4
'    Set lObj = ows.ListObjects(1)
''Clear existing data
'    lObj.ShowTotals = False
'    If lObj.ListRows.count > 0 Then
'        lObj.DataBodyRange.Rows.Delete
'    End If
'    lObj.ListRows.Add 1
'    Range(lObj.DataBodyRange.Cells(1, 1), lObj.DataBodyRange.Cells(UBound(dataArray, 1), UBound(dataArray, 2))) = dataArray
'    lObj.ListRows(1).Range.Interior.Pattern = xlNone
'    lObj.ShowTotals = True
''Set conditional formatting
'    With lObj
'        .DataBodyRange.FormatConditions.Delete
'        .DataBodyRange.FormatConditions.Add Type:=xlExpression, Formula1:="=MOD(ROW(),2)"
'        .DataBodyRange.FormatConditions(.DataBodyRange.FormatConditions.count).SetFirstPriority
'        With .DataBodyRange.FormatConditions(1).Interior
'            .PatternColorIndex = xlAutomatic
'            .ThemeColor = xlThemeColorDark1
'            .TintAndShade = -4.99893185216834E-02
'        End With
'        .DataBodyRange.FormatConditions(1).StopIfTrue = False
'    End With
''Format Print Layout
'    Call SheetFormatting
'End Sub


