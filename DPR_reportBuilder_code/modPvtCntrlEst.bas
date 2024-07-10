Attribute VB_Name = "modPvtCntrlEst"

Dim sCat10 As String
Dim sCat20 As String
Dim sCat21 As String
Dim sCat22 As String
Dim sCat30 As String
Dim sCat35 As String
Dim sCat40 As String
Dim sCat45 As String
Dim sCat49 As String
Dim sCat50 As String
Dim sCat51 As String
Dim sCat52 As String
Dim sCat53 As String
Dim sCat60 As String
Dim sCat61 As String
Dim sCat62 As String
Dim sCat70 As String


Dim dval As Double
Dim nVal As String

Function GetColumnNumber(colName As String) As Integer
    Dim col As ListColumn
    For Each col In lObj.ListColumns
        If col.name = colName Then
            GetColumnNumber = col.Index
            Exit For
        End If
    Next col
End Function

Sub CollectDataIntoDictionary()
    Dim dict As New Scripting.Dictionary
    Dim key As Variant

    Set ows = ThisWorkbook.Sheets("pivot data")
    Set lObj = ows.ListObjects("tblEdiphiPivotDataUseSplit")
    
    For i = 1 To iLvl
        Select Case i
            Case 1
                iLvl1Code = GetColumnNumber(sLvl1Code)
                iLvl1Item = GetColumnNumber(sLvl1Item)
            Case 2
                iLvl2Code = GetColumnNumber(sLvl2Code)
                iLvl2Item = GetColumnNumber(sLvl2Item)
            Case 3
                iLvl3Code = GetColumnNumber(sLvl3Code)
                iLvl3Item = GetColumnNumber(sLvl3Item)
            Case 4
                iLvl4Code = GetColumnNumber(sLvl4Code)
                iLvl4Item = GetColumnNumber(sLvl4Item)
            Case 5
                iLvl5Code = GetColumnNumber(sLvl5Code)
                iLvl5Item = GetColumnNumber(sLvl5Item)
        End Select
    Next i
    
    Call clear_sCat
    Dim row As ListRow
    For Each row In lObj.ListRows
        For i = 1 To 5
            If i = 1 Then
                dval = row.Range.Cells(1, 8).Value      'Labor
                nVal = row.Range.Cells(1, 44).Value
            ElseIf i = 2 Then
                dval = row.Range.Cells(1, 10).Value     'Material
                nVal = row.Range.Cells(1, 56).Value
            ElseIf i = 3 Then
                dval = row.Range.Cells(1, 12).Value     'Subcontractor
                nVal = row.Range.Cells(1, 53).Value
            ElseIf i = 4 Then
                dval = row.Range.Cells(1, 11).Value     'Equipment
                nVal = row.Range.Cells(1, 50).Value
            ElseIf i = 5 Then
                dval = row.Range.Cells(1, 13).Value     'Other
                nVal = row.Range.Cells(1, 47).Value
            End If
            Select Case nVal
                Case 10: sCat10 = sCat10 + dval
                Case 20: sCat20 = sCat20 + dval
                Case 21: sCat21 = sCat21 + dval
                Case 22: sCat22 = sCat22 + dval
                Case 30: sCat30 = sCat30 + dval
                Case 35: sCat35 = sCat35 + dval
                Case 40: sCat40 = sCat40 + dval
                Case 45: sCat45 = sCat45 + dval
                Case 49: sCat49 = sCat49 + dval
                Case 50: sCat50 = sCat50 + dval
                Case 51: sCat51 = sCat51 + dval
                Case 52: sCat52 = sCat52 + dval
                Case 53: sCat53 = sCat53 + dval
                Case 60: sCat60 = sCat60 + dval
                Case 61: sCat61 = sCat61 + dval
                Case 62: sCat62 = sCat62 + dval
                Case 70: sCat70 = sCat70 + dval
                Case Else
            End Select
            dval = 0
        Next i
    
        key = row.Range.Cells(1, 1).Value & "-" & row.Range.Cells(1, 5).Value
        If Not dict.Exists(key) Then
            dict.Add key, Array(row.Range(1, iLvl1Code).Value, row.Range(1, iLvl1Item).Value, _
                                row.Range(1, iLvl2Code).Value, row.Range(1, iLvl2Item).Value, _
                                row.Range(1, iLvl3Code).Value, row.Range(1, iLvl3Item).Value, _
                                row.Range(1, iLvl4Code).Value, row.Range(1, iLvl4Item).Value, _
                                row.Range(1, iLvl5Code).Value, row.Range(1, iLvl5Item).Value, _
                                row.Range.Cells(1, 2).Value, _
                                row.Range.Cells(1, 4).Value, _
                                row.Range.Cells(1, 68).Value, _
                                val(CDbl(row.Range.Cells(1, 5).Value)), _
                                row.Range.Cells(1, 6).Value, _
                                val(CDbl(row.Range.Cells(1, 67).Value)), _
                                val(CDbl(sCat10)), val(CDbl(sCat20)), val(CDbl(sCat21)), val(CDbl(sCat22)), _
                                val(CDbl(sCat30)), val(CDbl(sCat35)), val(CDbl(sCat40)), val(CDbl(sCat45)), _
                                val(CDbl(sCat49)), val(CDbl(sCat50)), val(CDbl(sCat51)), val(CDbl(sCat52)), _
                                val(CDbl(sCat53)), val(CDbl(sCat60)), val(CDbl(sCat61)), val(CDbl(sCat62)), _
                                val(CDbl(sCat70)), val(CDbl(row.Range.Cells(1, 66).Value)))
        Else
            q = dict.Item(oTxt)
            q(14) = q(14) + val(CDbl(row.Range.Cells(1, 5).Value))        'Takeoff Qty
            q(16) = q(16) + val(CDbl(row.Range.Cells(1, 67).Value))       'LaborHours
            q(17) = q(17) + val(CDbl(sCat10))
            q(18) = q(18) + val(CDbl(sCat20))
            q(19) = q(19) + val(CDbl(sCat21))
            q(20) = q(20) + val(CDbl(sCat22))
            q(21) = q(21) + val(CDbl(sCat30))
            q(22) = q(22) + val(CDbl(sCat35))
            q(23) = q(23) + val(CDbl(sCat40))
            q(24) = q(24) + val(CDbl(sCat45))
            q(25) = q(25) + val(CDbl(sCat49))
            q(26) = q(26) + val(CDbl(sCat50))
            q(27) = q(27) + val(CDbl(sCat51))
            q(28) = q(28) + val(CDbl(sCat52))
            q(29) = q(29) + val(CDbl(sCat53))
            q(30) = q(30) + val(CDbl(sCat60))
            q(31) = q(31) + val(CDbl(sCat61))
            q(32) = q(32) + val(CDbl(sCat62))
            q(33) = q(33) + val(CDbl(sCat70))
            q(34) = q(34) + val(CDbl(row.Range.Cells(1, 66).Value))     'Grand Total
            dict.Item(oTxt) = q
        End If
        Call clear_sCat
    Next row

    ReDim dataArray(1 To dict.Count, 1 To 34)
'    On Error Resume Next
    C = 0
    Dim k As Variant
    For Each k In dict.Keys
        C = C + 1
        dataArray(C, 1) = dict.Item(k)(0) 'Lvl1Code
        dataArray(C, 2) = dict.Item(k)(1) 'Lvl1Desc
        dataArray(C, 3) = dict.Item(k)(2) 'Lvl2Code
        dataArray(C, 4) = dict.Item(k)(3) 'Lvl2Desc
        dataArray(C, 5) = dict.Item(k)(4) 'Lvl3Code
        dataArray(C, 6) = dict.Item(k)(5) 'Lvl3Desc
        dataArray(C, 7) = dict.Item(k)(6) 'Lvl4Code
        dataArray(C, 8) = dict.Item(k)(7) 'Lvl4Desc
        dataArray(C, 9) = dict.Item(k)(8) 'Lvl5Code
        dataArray(C, 10) = dict.Item(k)(9) 'Lvl5Desc
        dataArray(C, 11) = dict.Item(k)(10) 'ItemCode
        dataArray(C, 12) = dict.Item(k)(11) 'Description
        dataArray(C, 13) = dict.Item(k)(12) 'ItemNote
        dataArray(C, 14) = dict.Item(k)(13) 'TakeoffQty
        dataArray(C, 15) = dict.Item(k)(14) 'TakeoffUnit
        If dict.Item(k)(15) > 0 Then dataArray(C, 16) = dict.Item(k)(15) 'Laborhours
        dataArray(C, 17) = dict.Item(k)(16) 'Labor10
        dataArray(C, 18) = dict.Item(k)(17) 'Material20
        dataArray(C, 19) = dict.Item(k)(18) 'MaterialLumpSum21
        dataArray(C, 20) = dict.Item(k)(19) 'MaterialUnitRate22
        dataArray(C, 21) = dict.Item(k)(20) 'Equip30
        dataArray(C, 22) = dict.Item(k)(21) 'Equip35
        dataArray(C, 23) = dict.Item(k)(22) 'Other40
        dataArray(C, 24) = dict.Item(k)(23) 'Other45
        dataArray(C, 25) = dict.Item(k)(24) 'Other49
        dataArray(C, 26) = dict.Item(k)(25) 'Sub50
        dataArray(C, 27) = dict.Item(k)(26) 'DPR Est 51
        dataArray(C, 28) = dict.Item(k)(27) 'DPR Cont 52
        dataArray(C, 29) = dict.Item(k)(28) 'DPR Cont 53
        dataArray(C, 30) = dict.Item(k)(29) 'Owner Allow 60
        dataArray(C, 31) = dict.Item(k)(30) 'ConstCont 61
        dataArray(C, 32) = dict.Item(k)(31) 'OwnerCont 62
        dataArray(C, 33) = dict.Item(k)(32) 'OH&P 70
        For i = 17 To 33
            dataArray(C, 34) = dataArray(C, 34) + dataArray(C, i) 'Total
        Next i
    Next k
    On Error GoTo 0
    Set rsNew = ADOCopyArrayIntoRecordsetCEst(argArray:=dataArray)
End Sub

Private Function ADOCopyArrayIntoRecordsetCEst(argArray As Variant) As ADODB.Recordset 'Create data recordset for pivot cache
Dim rsADO As ADODB.Recordset
Dim lngR As Long
Dim lngC As Long

    Set rsADO = New ADODB.Recordset
    For i = 1 To 5 'iLvl
        Select Case i
            Case 1
                sLvl1Code = sLvl1Code & i
                rsADO.Fields.Append sLvl1Code, adVariant
                rsADO.Fields.Append sLvl1Item, adVariant
            Case 2
                If sLvl2Code = "" Then sLvl2Code = "NA_Code" & i Else sLvl2Code = sLvl2Code & i
                If sLvl2Item = "" Then sLvl2Item = "NA_Item" & i
                rsADO.Fields.Append sLvl2Code, adVariant
                rsADO.Fields.Append sLvl2Item, adVariant
            Case 3
                If sLvl3Code = "" Then sLvl3Code = "NA_Code" & i Else sLvl3Code = sLvl3Code & i
                If sLvl3Item = "" Then sLvl3Item = "NA_Item" & i
                rsADO.Fields.Append sLvl3Code, adVariant
                rsADO.Fields.Append sLvl3Item, adVariant
            Case 4
                If sLvl4Code = "" Then sLvl4Code = "NA_Code" & i Else sLvl4Code = sLvl4Code & i
                If sLvl4Item = "" Then sLvl4Item = "NA_Item" & i
                rsADO.Fields.Append sLvl4Code, adVariant
                rsADO.Fields.Append sLvl4Item, adVariant
            Case 5
                If sLvl5Code = "" Then sLvl5Code = "NA_Code" & i Else sLvl5Code = sLvl5Code & i
                If sLvl5Item = "" Then sLvl5Item = "NA_Item" & i
                rsADO.Fields.Append sLvl5Code, adVariant
                rsADO.Fields.Append sLvl5Item, adVariant
        End Select
    Next i
    rsADO.Fields.Append "ItemCode", adVariant
    rsADO.Fields.Append "Description", adVariant
    rsADO.Fields.Append "ItemNote", adVariant
    rsADO.Fields.Append "TakeoffQty", adVariant
    rsADO.Fields.Append "TakeoffUnit", adVariant
    rsADO.Fields.Append "LaborHours", adVariant
    rsADO.Fields.Append "Labor10", adVariant
    rsADO.Fields.Append "Material20", adVariant
    rsADO.Fields.Append "Material21", adVariant
    rsADO.Fields.Append "Material22", adVariant
    rsADO.Fields.Append "Equipment30", adVariant
    rsADO.Fields.Append "Equipment35", adVariant
    rsADO.Fields.Append "Other40", adVariant
    rsADO.Fields.Append "Other45", adVariant
    rsADO.Fields.Append "Other49", adVariant
    rsADO.Fields.Append "Sub50", adVariant
    rsADO.Fields.Append "DPREst51", adVariant
    rsADO.Fields.Append "DPRCont52", adVariant
    rsADO.Fields.Append "SKUnitRates53", adVariant
    rsADO.Fields.Append "OwnerAllow60", adVariant
    rsADO.Fields.Append "ConstCont61", adVariant
    rsADO.Fields.Append "OwnerCont62", adVariant
    rsADO.Fields.Append "OHP70", adVariant
    rsADO.Fields.Append "GrandTotal", adVariant
    rsADO.Open
    On Error Resume Next
    For lngR = 1 To UBound(argArray, 1)
        rsADO.AddNew
           For lngC = 1 To UBound(argArray, 2) - 1
                rsADO.Fields(lngC - 1).Value = argArray(lngR, lngC)
           Next lngC
           rsADO.MoveNext
           lngC = 0
    Next lngR
    On Error GoTo 0
    rsADO.MoveFirst
    Set ADOCopyArrayIntoRecordsetCEst = rsADO
    Set rsADO = Nothing
End Function


Public Sub Create_PivotTable_ODBC_CntrlEst()
    bPvt = True
    Set ptCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, Version:=xlPivotTableVersion15)
    Set ptCache.Recordset = rsNew
'    Set ptCache = ActiveWorkbook.PivotCaches.Create( _
'    SourceType:=xlDatabase, SourceData:="tblEdiphiPivotDataUseSplit", Version:=xlPivotTableVersion15)
    
    ActiveWorkbook.Sheets.Add(Before:=Sheet4).name = sSht
    Set ows = ActiveSheet
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    Set pt = ptCache.CreatePivotTable(TableDestination:=ows.Range("B13"), tableName:=sSht)
    With pt
        .TableStyle2 = "DPR_CntrlEst"
        .HasAutoFormat = False
        .DisplayErrorString = True
        .ErrorString = "-"
        .NullString = "-"
        .ShowDrillIndicators = False
        .TableRange1.Font.Size = 12
        .TableRange1.Font.name = "Franklin Gothic Book"
        .TableRange1.VerticalAlignment = xlTop
        .RepeatItemsOnEachPrintedPage = False
        .ManualUpdate = True
        .ShowTableStyleColumnStripes = True
    End With
    X = 1

    On Error Resume Next
    For i = 1 To iLvl
    Select Case i
        Case 1 'Group Level 1
            With pt.PivotFields(sLvl1Code)
                .Orientation = xlRowField
                .Position = X
            End With
            With pt.PivotFields(sLvl1Code)
                .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                .LayoutForm = xlTabular
            End With
            X = X + 1
            With pt.PivotFields(sLvl1Item)
                .Orientation = xlRowField
                .Position = X
            End With
            With pt.PivotFields(sLvl1Item)
                '.Caption = sLvl1Name
                .LayoutBlankLine = True
                .LayoutSubtotalLocation = xlAtBottom
                .LayoutCompactRow = False
                .SubtotalName = "Subtotal: ?"
            End With
            X = X + 1
        Case 2
            With pt.PivotFields(sLvl2Code)
                .Orientation = xlRowField
                .Position = X
            End With
            With pt.PivotFields(sLvl2Code)
                .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                .LayoutForm = xlTabular
            End With
            X = X + 1
            With pt.PivotFields(sLvl2Item)
                .Orientation = xlRowField
                .Position = X
            End With
            With pt.PivotFields(sLvl2Item)
                '.Caption = sLvl2Name
                .LayoutBlankLine = True
                .LayoutSubtotalLocation = xlAtBottom
                .LayoutCompactRow = False
                .SubtotalName = "Subtotal: ?"
            End With
            X = X + 1
'Group Level 3
        Case 3
            With pt.PivotFields(sLvl3Code)
                .Orientation = xlRowField
                .Position = X
            End With
            With pt.PivotFields(sLvl3Code)
                .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                .LayoutForm = xlTabular
            End With
            X = X + 1
            With pt.PivotFields(sLvl3Item)
                .Orientation = xlRowField
                .Position = X
            End With
            With pt.PivotFields(sLvl3Item)
                '.Caption = sLvl3Name
                .LayoutBlankLine = True
                .LayoutSubtotalLocation = xlAtBottom
                .LayoutCompactRow = False
                .SubtotalName = "Subtotal: ?"
            End With
            X = X + 1
'Group Level 4
        Case 4
            With pt.PivotFields(sLvl4Code)
                .Orientation = xlRowField
                .Position = X
            End With
            With pt.PivotFields(sLvl4Code)
                .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                .LayoutForm = xlTabular
            End With
            X = X + 1
            With pt.PivotFields(sLvl4Item)
                .Orientation = xlRowField
                .Position = X
            End With
            With pt.PivotFields(sLvl4Item)
                '.Caption = sLvl4Name
                .LayoutBlankLine = True
                .LayoutSubtotalLocation = xlAtBottom
                .LayoutCompactRow = False
                .SubtotalName = "Subtotal: ?"
            End With
            X = X + 1
'Group Level 5
        Case 5
            With pt.PivotFields(sLvl5Code)
                .Orientation = xlRowField
                .Position = X
            End With
            With pt.PivotFields(sLvl5Code)
                .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                .LayoutForm = xlTabular
            End With
            X = X + 1
            With pt.PivotFields(sLvl5Item)
                .Orientation = xlRowField
                .Position = X
            End With
            With pt.PivotFields(sLvl5Item)
                '.Caption = sLvl5Name
                .LayoutBlankLine = True
                .LayoutSubtotalLocation = xlAtBottom
                .LayoutCompactRow = False
                .SubtotalName = "Subtotal: ?"
            End With
            X = X + 1
        End Select
    Next i
    On Error GoTo 0
    
    On Error Resume Next
'Field ItemCode
    With pt.PivotFields("ItemCode")
        .Orientation = xlRowField
        .Position = X
    End With
'Field ItemDesc
    X = X + 1
    With pt.PivotFields("Description")
        .Orientation = xlRowField
        .Position = X
    End With
'Field Comments
    X = X + 1
    With pt.PivotFields("ItemNote")
        .Orientation = xlRowField
        .Position = X
    End With
'Field TOQty
    X = X + 1
    With pt.PivotFields("TakeoffQty")
        .Orientation = xlRowField
        .Position = X
    End With
'Field TOUnit
    X = X + 1
    With pt.PivotFields("TakeoffUnit")
        .Orientation = xlRowField
        .Position = X
    End With
    pt.ManualUpdate = False
    
'Field Item Code
    Set pf = pt.PivotFields("ItemCode")
    With pf
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .LayoutForm = xlTabular
        .LabelRange.EntireColumn.Hidden = True
    End With
'Field ItemDesc
    Set pf = pt.PivotFields("Description")
    With pf
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
'Field Comments
    Set pf = pt.PivotFields("ItemNote")
        With pf
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
'Field TOQty
    Set pf = pt.PivotFields("TakeoffQty")
    With pf
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
'Field TOUnit
    Set pf = pt.PivotFields("TakeoffUnit")
    With pf
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
    
    pt.PivotSelect "'Description'[All]", xlLabelOnly + xlFirstRow, True
    With Selection
        .ColumnWidth = 40
        .WrapText = True
        .Orientation = 0
        .AddIndent = True
        .IndentLevel = 1
    End With
    
    pt.PivotSelect "'ItemNote'[All]", xlLabelOnly + xlFirstRow, True
    With Selection
        .ColumnWidth = 30
        .WrapText = True
    End With
    pt.PivotFields("ItemNote").PivotItems("(blank)").Caption = "-"
    pt.PivotSelect "ItemNote['(blank)']", xlDataAndLabel, True
    pt.PivotSelect "ItemNote", xlDataAndLabel, True
    
    pt.PivotSelect "'TakeoffQty'[All]", xlLabelOnly + xlFirstRow, True
    With Selection
        .NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
        .ColumnWidth = 11
        .HorizontalAlignment = xlRight
    End With
    pt.PivotFields("TakeoffQty").PivotItems("(blank)").Caption = "-"
    pt.PivotSelect "TakeoffQty['(blank)']", xlDataAndLabel, True
    pt.PivotSelect "TakeoffQty", xlDataAndLabel, True
    
    pt.PivotSelect "'TakeoffUnit'[All]", xlLabelOnly + xlFirstRow, True
    With Selection
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
    End With
    On Error GoTo 0
'Manhours
    pt.AddDataField pt.PivotFields("LaborHours"), "Sum of LaborHours", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of LaborHours")
        .NumberFormat = "#,##0_);[Red](#,##0);_(""-""??_)"
    End With
'Labor10
    pt.AddDataField pt.PivotFields("Labor10"), "Sum of Labor10", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of Labor10")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'Material20
    pt.AddDataField pt.PivotFields("Material20"), "Sum of Material20", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of Material20")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'Material21
    pt.AddDataField pt.PivotFields("Material21"), "Sum of Material21", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of Material21")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'Material22
    pt.AddDataField pt.PivotFields("Material22"), "Sum of Material22", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of Material22")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'Equipment30
    pt.AddDataField pt.PivotFields("Equipment30"), "Sum of Equipment30", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of Equipment30")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'Equipment35
    pt.AddDataField pt.PivotFields("Equipment35"), "Sum of Equipment35", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of Equipment35")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'Other40
    pt.AddDataField pt.PivotFields("Other40"), "Sum of Other40", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of Other40")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'Other45
    pt.AddDataField pt.PivotFields("Other45"), "Sum of Other45", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of Other45")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'Other40
    pt.AddDataField pt.PivotFields("Other49"), "Sum of Other49", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of Other49")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'Sub50
    pt.AddDataField pt.PivotFields("Sub50"), "Sum of Sub50", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of Sub50")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'DPREst51
    pt.AddDataField pt.PivotFields("DPREst51"), "Sum of DPREst51", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of DPREst51")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'DPRCont52
    pt.AddDataField pt.PivotFields("DPRCont52"), "Sum of DPRCont52", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of DPRCont52")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'DPRCont53
    pt.AddDataField pt.PivotFields("SKUnitRates53"), "Sum of SKUnitRates53", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of SKUnitRates53")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'OwnerAllow60
    pt.AddDataField pt.PivotFields("OwnerAllow60"), "Sum of OwnerAllow60", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of OwnerAllow60")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'ConstCont61
    pt.AddDataField pt.PivotFields("ConstCont61"), "Sum of ConstCont61", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of ConstCont61")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'OwnerCont62
    pt.AddDataField pt.PivotFields("OwnerCont62"), "Sum of OwnerCont62", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of OwnerCont62")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'OHP70
    pt.AddDataField pt.PivotFields("OHP70"), "Sum of OHP70", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of OHP70")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'GrandTotal
    pt.AddDataField pt.PivotFields("GrandTotal"), "Sum of GrandTotal", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of GrandTotal")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'Format Report Levels
    For Y = 1 To iLvl
        Select Case Y
            Case 1
                Call FrmtCLvl1
            Case 2
                Call FrmtCLvl2
            Case 3
                Call FrmtCLvl3
            Case 4
                Call FrmtCLvl4
            Case 5
                Call FrmtCLvl5
       End Select
    Next Y
    Call FrmtGTRow
    Call CEstAddons
    Call SetSheetCHeadings
    Call ReadGrandTotalRow
    
    If bCode = True Then Call FormatCntrlEst
    ows.Range("A1").Select
    iLvl = 0
    pic = "DPRLogo.25.png"
    Call PageSetup
    Call ResetSheetScroll
    Application.ScreenUpdating = True
    bPvt = False
    On Error GoTo 0
    If bCode = True Then MsgBox "Highlighted rows indicate uncategorized items", vbInformation, "Category Codes"
    Set ptCache = Nothing
    Set pt = Nothing

End Sub

Private Sub FrmtCLvl1()
'Format Level 1
    Columns("B:B").ColumnWidth = 0.05
    Columns("C:C").ColumnWidth = colW
    Application.PivotTableSelection = True
    pt.PivotSelect "'" & sLvl1Item & "'[All]", xlDataAndLabel + xlFirstRow, True
    With Selection
        .WrapText = False
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
        .Font.Color = -16777216
        .Font.TintAndShade = 0
        .Font.Bold = True
        .Font.Size = 12
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).ThemeColor = 1
        .Borders(xlEdgeTop).TintAndShade = -0.14996795556505
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).ThemeColor = 1
        .Borders(xlEdgeBottom).TintAndShade = -0.14996795556505
        .Borders(xlEdgeBottom).Weight = xlThin
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorDark2
        .Interior.TintAndShade = 0.799981688894314
        .Interior.PatternTintAndShade = 0
    End With
'Format Level - 1 Subtotals
    pt.PivotSelect "'" & sLvl1Item & "'[All;Total]", xlDataAndLabel + xlFirstRow, True
    With Selection
        .WrapText = False
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .Font.Size = 12
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleNone
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
        .Font.ThemeFont = xlThemeFontMinor
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).ThemeColor = 4
        .Borders(xlEdgeTop).TintAndShade = 0
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).ThemeColor = 4
        .Borders(xlEdgeBottom).TintAndShade = 0
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    pt.PivotSelect "'" & sLvl1Item & "'[All;Total]", xlDataOnly + xlFirstRow, True
    With Selection
        .WrapText = False
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub FrmtCLvl2()
''Format Level 2
    Columns("D:D").EntireColumn.Hidden = True
    Columns("E:E").ColumnWidth = colW
    pt.PivotSelect "'" & sLvl2Item & "'[All]", xlDataAndLabel + xlFirstRow, True
    With Selection
        .WrapText = False
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = -16777216
        .Font.TintAndShade = 0
    End With
''Format Level - 2 Subtotals
    pt.PivotSelect "'" & sLvl2Item & "'[All;Total]", xlDataAndLabel + xlFirstRow, True
    With Selection
        .WrapText = False
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = -16777216
        .Font.TintAndShade = 0
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    pt.PivotSelect "'" & sLvl2Item & "'[All;Total]", xlDataOnly + xlFirstRow, True
    Selection.HorizontalAlignment = xlRight
End Sub

Private Sub FrmtCLvl3()
''Format Level 3
    Columns("F:F").EntireColumn.Hidden = True
    Columns("G:G").ColumnWidth = colW
    pt.PivotSelect "'" & sLvl3Item & "'[All]", xlLabelOnly + xlFirstRow, True
    With Selection
        .WrapText = False
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = -16777216
        .Font.TintAndShade = 0
    End With
''Format Level - 3 Subtotals
    pt.PivotSelect "'" & sLvl3Item & "'[All;Total]", xlDataAndLabel + xlFirstRow, True
    With Selection
        .WrapText = False
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = -16777216
        .Font.TintAndShade = 0
    End With
    pt.PivotSelect "'" & sLvl3Item & "'[All;Total]", xlDataOnly + xlFirstRow, True
    Selection.HorizontalAlignment = xlRight
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
End Sub

Private Sub FrmtCLvl4()
''Format Level 4
    Columns("H:H").EntireColumn.Hidden = True
    Columns("I:I").ColumnWidth = colW
    pt.PivotSelect "'" & sLvl4Item & "'[All]", xlLabelOnly + xlFirstRow, True
    With Selection
        .WrapText = False
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = -16777216
        .Font.TintAndShade = 0
    End With
''Format Level - 4 Subtotals
    pt.PivotSelect "'" & sLvl4Item & "'[All;Total]", xlDataAndLabel + xlFirstRow, True
    With Selection
        .WrapText = False
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = -16777216
        .Font.TintAndShade = 0
    End With
    pt.PivotSelect "'" & sLvl4Item & "'[All;Total]", xlDataOnly + xlFirstRow, True
    Selection.HorizontalAlignment = xlRight
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
End Sub

Private Sub FrmtCLvl5()
''Format Level 5
    Columns("J:J").EntireColumn.Hidden = True
    Columns("K:K").ColumnWidth = colW
    pt.PivotSelect "'" & sLvl5Item & "'[All]", xlLabelOnly + xlFirstRow, True
    With Selection
        .WrapText = False
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = -16777216
        .Font.TintAndShade = 0
    End With
''Format Level - 5 Subtotals
    pt.PivotSelect "'" & sLvl5Item & "'[All;Total]", xlDataAndLabel + xlFirstRow, True
    With Selection
        .WrapText = False
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = -16777216
        .Font.TintAndShade = 0
    End With
    pt.PivotSelect "'" & sLvl5Item & "'[All;Total]", xlDataOnly + xlFirstRow, True
    Selection.HorizontalAlignment = xlRight
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
End Sub

Sub FrmtCGTRow()
    pt.GrandTotalName = StrConv(sGTLvl1 & " Subtotal", vbUpperCase)
    pt.PivotSelect "'Column Grand Total'", xlDataAndLabel + xlFirstRow, True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        '.NumberFormat = "$#,##0_);($#,##0)"
        .Font.Size = 12
        .Font.Bold = True
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

Sub SetSheetCHeadings()
Dim sLeft As Single
    
    Set ows = ActiveSheet
    Set pt = ows.PivotTables(1)
    r = 7
    With ows.Range("B1")
        .FormulaR1C1 = StrConv(sRprt, vbUpperCase)
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
        .Font.name = "FrnkGothITC Bk BT"
        .Font.Size = 18
        .RowHeight = 35.25
    End With
    With ows.PivotTables(1).TableRange1
        iCol = .Cells(.Cells.Count).Column
    End With
    With ows.Range(Cells(1, 2), Cells(1, iCol))
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    End With
        
    For i = 1 To iLvl
        Select Case i
            Case 1
                With ows.Cells(r, 3)
                    .FormulaR1C1 = StrConv(sLvl1Name, vbUpperCase)
                    .HorizontalAlignment = xlLeft
                End With
            Case 2
                With ows.Cells(r + 1, 5)
                    .FormulaR1C1 = StrConv(sLvl2Name, vbUpperCase)
                    .HorizontalAlignment = xlLeft
                End With
            Case 3
                With ows.Cells(r + 2, 7)
                    .FormulaR1C1 = StrConv(sLvl3Name, vbUpperCase)
                    .HorizontalAlignment = xlLeft
                End With
            Case 4
                With ows.Cells(r + 3, 9)
                    .FormulaR1C1 = StrConv(sLvl4Name, vbUpperCase)
                    .HorizontalAlignment = xlLeft
                End With
            Case 5
                With ows.Cells(r + 4, 11)
                    .FormulaR1C1 = StrConv(sLvl5Name, vbUpperCase)
                    .HorizontalAlignment = xlLeft
                End With
        End Select
    Next
    ows.Cells(r + iLvl, iCol - 22).FormulaR1C1 = "DESCRIPTION"
    ows.Cells(r + iLvl, iCol - 21).FormulaR1C1 = "COMMENTS"
    ows.Cells(r + iLvl, iCol - 20).FormulaR1C1 = "QUANTITY"
    ows.Cells(r + iLvl, iCol - 19).FormulaR1C1 = "UNIT"
    ows.Cells(r + iLvl, iCol - 18).FormulaR1C1 = "MANHOURS"
    ows.Cells(r + iLvl, iCol - 17).FormulaR1C1 = "LABOR " & Chr(10) & "10"
    ows.Cells(r + iLvl, iCol - 16).FormulaR1C1 = "MATERIAL " & Chr(10) & "20"
    ows.Cells(r + iLvl, iCol - 15).FormulaR1C1 = "MAT LUMP SUM " & Chr(10) & "21"
    ows.Cells(r + iLvl, iCol - 14).FormulaR1C1 = "MAT UNIT RATE " & Chr(10) & "22"
    ows.Cells(r + iLvl, iCol - 13).FormulaR1C1 = "EQUIPMENT " & Chr(10) & "30"
    ows.Cells(r + iLvl, iCol - 12).FormulaR1C1 = "EQUIP OPERATED " & Chr(10) & "35"
    ows.Cells(r + iLvl, iCol - 11).FormulaR1C1 = "OTHER GCs/GRs " & Chr(10) & "40"
    ows.Cells(r + iLvl, iCol - 10).FormulaR1C1 = "OTHER MARKUPS " & Chr(10) & "45"
    ows.Cells(r + iLvl, iCol - 9).FormulaR1C1 = "SK DPR " & Chr(10) & "49"
    ows.Cells(r + iLvl, iCol - 8).FormulaR1C1 = "SK TRADES " & Chr(10) & "50"
    ows.Cells(r + iLvl, iCol - 7).FormulaR1C1 = "DPR EST " & Chr(10) & "51"
    ows.Cells(r + iLvl, iCol - 6).FormulaR1C1 = "DPR CONT " & Chr(10) & "52"
    ows.Cells(r + iLvl, iCol - 5).FormulaR1C1 = "SK UNIT RATES " & Chr(10) & "53"
    ows.Cells(r + iLvl, iCol - 4).FormulaR1C1 = "OWNER ALLOW " & Chr(10) & "60"
    ows.Cells(r + iLvl, iCol - 3).FormulaR1C1 = "CONST CONT " & Chr(10) & "61"
    ows.Cells(r + iLvl, iCol - 2).FormulaR1C1 = "OWNER CONT " & Chr(10) & "62"
    ows.Cells(r + iLvl, iCol - 1).FormulaR1C1 = "OH&P " & Chr(10) & "70"
    ows.Cells(r + iLvl, iCol).FormulaR1C1 = "TOTAL"
    ows.Range(Cells(r + iLvl, iCol - 18), Cells(r + iLvl, iCol)).ColumnWidth = 16.33
    ows.Range(Cells(r + iLvl, iCol - 22), Cells(r + iLvl, iCol)).HorizontalAlignment = xlCenter
    ows.Range(Cells(r + iLvl, iCol - 22), Cells(r + iLvl, iCol)).VerticalAlignment = xlCenter
    ows.Rows(r + iLvl + 1 & ":13").EntireRow.Hidden = True
    
    With ows.Range(Cells(r, 2), Cells(r + iLvl, iCol))
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
        .Font.name = "FrnkGothITC Bk BT"
        .Font.Size = 12
        .Font.Bold = False
    End With
    
    ows.Rows("2:6").RowHeight = 17.55
    ows.Rows(r + iLvl).RowHeight = 35
    ows.Columns(iCol).ColumnWidth = 20
    
    Sheets("EstData").Shapes("grpHeading").Copy
    Application.GoTo Sheets(sSht).Range("B1")
    ActiveSheet.Paste
    Set myShape = ows.Shapes("grpHeading")
    Set cl = Range(Cells(1, 2), Cells(6, iCol))
    With myShape
        .Left = cl.Left
        .Top = cl.Top
        .Height = cl.Height
        .Width = cl.Width
    End With
    With myShape.GroupItems
        '.Item("txtSubHeading3").TextFrame.Characters.Text = _
            "CONSTRUCTION AREA: " & FormatNumber(dJobSz, 0) & " " & StrConv(sJobUM, vbUpperCase)
        '.Item("txtSubHeading6").TextFrame.Characters.Text = _
            "DURATION: " & StrConv(sDuration, vbUpperCase)
    End With
    ows.Range("A1").Select
End Sub

Sub ReadGrandTotalRow()
    Dim dataRange As Range
    Dim grandTotalRow As Range
    Dim cell As Range
    Dim colIndex As Integer
    Dim sColName As String
    
    ' Set the worksheet and pivot table
    Set ows = ActiveSheet
    Set pt = ows.PivotTables(1)
    
    ' Get the data body range of the pivot table
    Set dataRange = pt.DataBodyRange
    
    ' Check if the pivot table has a grand total row
    If pt.TableRange2.Rows.Count > dataRange.Rows.Count Then
        ' The grand total row is the last row in the data body range
        Set grandTotalRow = dataRange.Rows(dataRange.Rows.Count)
         
        ' Loop through each cell in the grand total row
        For colIndex = 1 To grandTotalRow.Columns.Count
            Set cell = grandTotalRow.Cells(1, colIndex)
            On Error Resume Next
            sColName = pt.PivotFields(cell.Column).name
            On Error GoTo 0
            Select Case colIndex
            Case 4, 5, 7, 9, 10, 14
                If cell.Value = 0 Then
                    cell.EntireColumn.Hidden = True
                End If
            Case Else
            End Select
        Next colIndex
    Else
        MsgBox "The pivot table does not have a grand total row."
    End If
End Sub

Sub clear_sCat()
    sCat10 = 0
    sCat20 = 0
    sCat21 = 0
    sCat22 = 0
    sCat30 = 0
    sCat35 = 0
    sCat40 = 0
    sCat45 = 0
    sCat49 = 0
    sCat50 = 0
    sCat51 = 0
    sCat52 = 0
    sCat53 = 0
    sCat60 = 0
    sCat61 = 0
    sCat62 = 0
    sCat70 = 0
End Sub
