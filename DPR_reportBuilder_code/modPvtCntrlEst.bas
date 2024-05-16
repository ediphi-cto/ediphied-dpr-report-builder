Attribute VB_Name = "modPvtCntrlEst"
Public Sub Create_PivotTable_ODBC_CntrlEst()
    bPvt = True
    Set ptCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, Version:=xlPivotTableVersion15)
    Set ptCache.Recordset = rsNew
    ActiveWorkbook.Sheets.Add(Before:=Sheet4).Name = sSht
    Set ows = ActiveSheet
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    Set pt = ptCache.CreatePivotTable(TableDestination:=ows.Range("B13"), TableName:=sSht)
    With pt
        .TableStyle2 = "DPR_CntrlEst"
        .HasAutoFormat = False
        .DisplayErrorString = True
        .ErrorString = "-"
        .NullString = "-"
        .ShowDrillIndicators = False
        .TableRange1.Font.Size = 12
        .TableRange1.Font.Name = "Franklin Gothic Book"
        .TableRange1.VerticalAlignment = xlTop
        .RepeatItemsOnEachPrintedPage = False
        .ManualUpdate = True
        .ShowTableStyleColumnStripes = True
    End With
    x = 1

    On Error Resume Next
    For i = 1 To iLvl
    Select Case i
        Case 1 'Group Level 1
            With pt.PivotFields(sLvl1Code)
                .Orientation = xlRowField
                .Position = x
            End With
            With pt.PivotFields(sLvl1Code)
                .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                .LayoutForm = xlTabular
            End With
            x = x + 1
            With pt.PivotFields(sLvl1Item)
                .Orientation = xlRowField
                .Position = x
            End With
            With pt.PivotFields(sLvl1Item)
                '.Caption = sLvl1Name
                .LayoutBlankLine = True
                .LayoutSubtotalLocation = xlAtBottom
                .LayoutCompactRow = False
                .SubtotalName = "Subtotal: ?"
            End With
            x = x + 1
        Case 2
            With pt.PivotFields(sLvl2Code)
                .Orientation = xlRowField
                .Position = x
            End With
            With pt.PivotFields(sLvl2Code)
                .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                .LayoutForm = xlTabular
            End With
            x = x + 1
            With pt.PivotFields(sLvl2Item)
                .Orientation = xlRowField
                .Position = x
            End With
            With pt.PivotFields(sLvl2Item)
                '.Caption = sLvl2Name
                .LayoutBlankLine = True
                .LayoutSubtotalLocation = xlAtBottom
                .LayoutCompactRow = False
                .SubtotalName = "Subtotal: ?"
            End With
            x = x + 1
'Group Level 3
        Case 3
            With pt.PivotFields(sLvl3Code)
                .Orientation = xlRowField
                .Position = x
            End With
            With pt.PivotFields(sLvl3Code)
                .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                .LayoutForm = xlTabular
            End With
            x = x + 1
            With pt.PivotFields(sLvl3Item)
                .Orientation = xlRowField
                .Position = x
            End With
            With pt.PivotFields(sLvl3Item)
                '.Caption = sLvl3Name
                .LayoutBlankLine = True
                .LayoutSubtotalLocation = xlAtBottom
                .LayoutCompactRow = False
                .SubtotalName = "Subtotal: ?"
            End With
            x = x + 1
'Group Level 4
        Case 4
            With pt.PivotFields(sLvl4Code)
                .Orientation = xlRowField
                .Position = x
            End With
            With pt.PivotFields(sLvl4Code)
                .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                .LayoutForm = xlTabular
            End With
            x = x + 1
            With pt.PivotFields(sLvl4Item)
                .Orientation = xlRowField
                .Position = x
            End With
            With pt.PivotFields(sLvl4Item)
                '.Caption = sLvl4Name
                .LayoutBlankLine = True
                .LayoutSubtotalLocation = xlAtBottom
                .LayoutCompactRow = False
                .SubtotalName = "Subtotal: ?"
            End With
            x = x + 1
'Group Level 5
        Case 5
            With pt.PivotFields(sLvl5Code)
                .Orientation = xlRowField
                .Position = x
            End With
            With pt.PivotFields(sLvl5Code)
                .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                .LayoutForm = xlTabular
            End With
            x = x + 1
            With pt.PivotFields(sLvl5Item)
                .Orientation = xlRowField
                .Position = x
            End With
            With pt.PivotFields(sLvl5Item)
                '.Caption = sLvl5Name
                .LayoutBlankLine = True
                .LayoutSubtotalLocation = xlAtBottom
                .LayoutCompactRow = False
                .SubtotalName = "Subtotal: ?"
            End With
            x = x + 1
        End Select
    Next i
    On Error GoTo 0
    
    On Error Resume Next
'Field ItemCode
    With pt.PivotFields("ItemCode")
        .Orientation = xlRowField
        .Position = x
    End With
'Field ItemDesc
    x = x + 1
    With pt.PivotFields("Description")
        .Orientation = xlRowField
        .Position = x
    End With
'Field Comments
    x = x + 1
    With pt.PivotFields("ItemNote")
        .Orientation = xlRowField
        .Position = x
    End With
'Field TOQty
    x = x + 1
    With pt.PivotFields("TakeoffQty")
        .Orientation = xlRowField
        .Position = x
    End With
'Field TOUnit
    x = x + 1
    With pt.PivotFields("TakeoffUnit")
        .Orientation = xlRowField
        .Position = x
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
    pt.AddDataField pt.PivotFields("Manhours"), "Sum of Manhours", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of Manhours")
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
'Equipment30
    pt.AddDataField pt.PivotFields("Equipment30"), "Sum of Equipment30", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of Equipment30")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'Other40
    pt.AddDataField pt.PivotFields("Other40"), "Sum of Other40", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of Other40")
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
    CEstAddons
    Call SetSheetCHeadings
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
    Set ptCache.Recordset = Nothing
    Set ptCache = Nothing
    Set pt = Nothing
    Set rsNew = Nothing
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
        .Font.Name = "FrnkGothITC Bk BT"
        .Font.Size = 18
        .RowHeight = 35.25
    End With
    With ows.PivotTables(1).TableRange1
        iCol = .Cells(.Cells.count).Column
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
    ows.Cells(r + iLvl, iCol - 16).FormulaR1C1 = "DESCRIPTION"
    ows.Cells(r + iLvl, iCol - 15).FormulaR1C1 = "COMMENTS"
    ows.Cells(r + iLvl, iCol - 14).FormulaR1C1 = "QUANTITY"
    ows.Cells(r + iLvl, iCol - 13).FormulaR1C1 = "UNIT"
    ows.Cells(r + iLvl, iCol - 12).FormulaR1C1 = "MANHOURS"
    ows.Cells(r + iLvl, iCol - 11).FormulaR1C1 = "LABOR " & Chr(10) & "10"
    ows.Cells(r + iLvl, iCol - 10).FormulaR1C1 = "MATERIAL " & Chr(10) & "20"
    ows.Cells(r + iLvl, iCol - 9).FormulaR1C1 = "EQUIPMENT " & Chr(10) & "30"
    ows.Cells(r + iLvl, iCol - 8).FormulaR1C1 = "OTHER " & Chr(10) & "40"
    ows.Cells(r + iLvl, iCol - 7).FormulaR1C1 = "SUB " & Chr(10) & "50"
    ows.Cells(r + iLvl, iCol - 6).FormulaR1C1 = "DPR EST " & Chr(10) & "51"
    ows.Cells(r + iLvl, iCol - 5).FormulaR1C1 = "DPR CONT " & Chr(10) & "52"
    ows.Cells(r + iLvl, iCol - 4).FormulaR1C1 = "OWNER ALLOW " & Chr(10) & "60"
    ows.Cells(r + iLvl, iCol - 3).FormulaR1C1 = "CONST CONT " & Chr(10) & "61"
    ows.Cells(r + iLvl, iCol - 2).FormulaR1C1 = "OWNER CONT " & Chr(10) & "62"
    ows.Cells(r + iLvl, iCol - 1).FormulaR1C1 = "OH&P " & Chr(10) & "70"
    ows.Cells(r + iLvl, iCol).FormulaR1C1 = "TOTAL"
    ows.Range(Cells(r + iLvl, iCol - 12), Cells(r + iLvl, iCol)).ColumnWidth = 16.33
    ows.Range(Cells(r + iLvl, iCol - 16), Cells(r + iLvl, iCol)).HorizontalAlignment = xlCenter
    ows.Range(Cells(r + iLvl, iCol - 16), Cells(r + iLvl, iCol)).VerticalAlignment = xlCenter
    ows.Rows(r + iLvl + 1 & ":13").EntireRow.Hidden = True
    
    With ows.Range(Cells(r, 2), Cells(r + iLvl, iCol))
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
        .Font.Name = "FrnkGothITC Bk BT"
        .Font.Size = 12
        .Font.Bold = False
    End With
    
    ows.Rows("2:6").RowHeight = 17.55
    ows.Rows(r + iLvl).RowHeight = 35
    ows.Columns(iCol).ColumnWidth = 20
    
    Sheets("EstData").Shapes("grpHeading").Copy
    Application.Goto Sheets(sSht).Range("B1")
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



