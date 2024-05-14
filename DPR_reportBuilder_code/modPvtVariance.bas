Attribute VB_Name = "modPvtVariance"
Option Explicit
Public Sub Create_PivotTable_ODBC_MO_VAR()
    Application.ScreenUpdating = False
    bPvt = True
    Set ptCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, Version:=xlPivotTableVersion15)
    Set ptCache.Recordset = rsNew
    ActiveWorkbook.Sheets.Add(Before:=Sheet4).Name = sSht
    Set ows = ActiveSheet
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    With ptCache
        .CreatePivotTable TableDestination:=ows.Range("B13"), TableName:=sSht
    End With
    Set pt = ows.PivotTables(1)
    With pt
        .TableStyle2 = "DPR_Estimating_Style_01"
        .HasAutoFormat = False
        .DisplayErrorString = True
        .ErrorString = "0"
        .NullString = "0"
        .ShowDrillIndicators = False
        .TableRange1.Font.Size = 12
        .TableRange1.Font.Name = "Franklin Gothic Book"
        .TableRange1.VerticalAlignment = xlTop
        .RepeatItemsOnEachPrintedPage = False
        .ManualUpdate = True
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
    
'Field Est1_Qty
    X = X + 1
    With pt.PivotFields("Est1_Qty")
        .Orientation = xlRowField
        .Position = X
    End With
'Field EST1_UM
    X = X + 1
    With pt.PivotFields("Est1_UM")
        .Orientation = xlRowField
        .Position = X
    End With
'Field Est1_UnitPrice
    X = X + 1
    With pt.PivotFields("Est1_Unit")
        .Orientation = xlRowField
        .Position = X
    End With
'Field Est2_Qty
    X = X + 1
    With pt.PivotFields("Est2_Qty")
        .Orientation = xlRowField
        .Position = X
    End With
'Field Est2_UM
    X = X + 1
    With pt.PivotFields("Est2_UM")
        .Orientation = xlRowField
        .Position = X
    End With
'Field Est2_UnitPrice
    X = X + 1
    With pt.PivotFields("Est2_Unit")
        .Orientation = xlRowField
        .Position = X
    End With
'Field VarQty
    X = X + 1
    With pt.PivotFields("VarQty")
        .Orientation = xlRowField
        .Position = X
    End With
'Field VarUnit
    X = X + 1
    With pt.PivotFields("VarUnit")
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
'Field Est1_Qty
    Set pf = pt.PivotFields("Est1_Qty")
    With pf
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
'Field Est1_UM
    Set pf = pt.PivotFields("Est1_UM")
    With pf
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
'Field Est1_Unit
    Set pf = pt.PivotFields("Est1_Unit")
    With pf
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
'Field Est2_Qty
    Set pf = pt.PivotFields("Est2_Qty")
    With pf
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
'Field Est2_UM
    Set pf = pt.PivotFields("Est2_UM")
    With pf
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
'Field Est2_Unit
    Set pf = pt.PivotFields("Est2_Unit")
    With pf
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
'Field VarQty
    Set pf = pt.PivotFields("VarQty")
    With pf
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With
'Field VarUnit
    Set pf = pt.PivotFields("VarUnit")
    With pf
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .LayoutForm = xlTabular
        .RepeatLabels = True
    End With

'Format Fields
    pt.PivotSelect "'Description'[All]", xlLabelOnly + xlFirstRow, True
    With Selection
        .ColumnWidth = 60
        .WrapText = True
        .Orientation = 0
        .AddIndent = True
        .IndentLevel = 1
    End With
    pt.PivotSelect "'Est1_Qty'[All]", xlLabelOnly + xlFirstRow, True
    With Selection
        .NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
        .ColumnWidth = 11
        .HorizontalAlignment = xlRight
    End With
    pt.PivotFields("Est1_Qty").PivotItems("(blank)").Caption = "-"
    pt.PivotSelect "Est1_Qty['(blank)']", xlDataAndLabel, True
    pt.PivotSelect "Est1_Qty", xlDataAndLabel, True
    
    pt.PivotSelect "'Est1_UM'[All]", xlLabelOnly + xlFirstRow, True
    With Selection
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
    End With
    
    pt.PivotSelect "'Est1_Unit'[All]", xlLabelOnly + xlFirstRow, True
    With Selection
        .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
        .ColumnWidth = 16
        .HorizontalAlignment = xlRight
    End With
    pt.PivotFields("Est1_Unit").PivotItems("(blank)").Caption = "-"
    pt.PivotSelect "Est1_Unit['(blank)']", xlDataAndLabel, True
    pt.PivotSelect "Est1_Unit", xlDataAndLabel, True
    
    pt.PivotSelect "'Est2_Qty'[All]", xlLabelOnly + xlFirstRow, True
    With Selection
        .NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
        .ColumnWidth = 11
        .HorizontalAlignment = xlRight
    End With
    pt.PivotFields("Est2_Qty").PivotItems("(blank)").Caption = "-"
    pt.PivotSelect "Est2_Qty['(blank)']", xlDataAndLabel, True
    pt.PivotSelect "Est2_Qty", xlDataAndLabel, True
    
    pt.PivotSelect "'Est2_UM'[All]", xlLabelOnly + xlFirstRow, True
    With Selection
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
    End With
    
    pt.PivotSelect "'Est2_Unit'[All]", xlLabelOnly + xlFirstRow, True
    With Selection
        .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
        .ColumnWidth = 16
        .HorizontalAlignment = xlRight
    End With
    pt.PivotFields("Est2_Unit").PivotItems("(blank)").Caption = "-"
    pt.PivotSelect "Est2_Unit['(blank)']", xlDataAndLabel, True
    pt.PivotSelect "Est2_Unit", xlDataAndLabel, True

    pt.PivotSelect "'VarQty'[All]", xlLabelOnly + xlFirstRow, True
    With Selection
        .NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
        .ColumnWidth = 11
        .HorizontalAlignment = xlRight
    End With
    pt.PivotFields("VarQty").PivotItems("(blank)").Caption = "-"
    pt.PivotSelect "VarQty['(blank)']", xlDataAndLabel, True
    pt.PivotSelect "VarQty", xlDataAndLabel, True
    
    pt.PivotSelect "'VarUnit'[All]", xlLabelOnly + xlFirstRow, True
    With Selection
        .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
        .ColumnWidth = 16
        .HorizontalAlignment = xlRight
    End With
    pt.PivotFields("VarUnit").PivotItems("(blank)").Caption = "-"
    pt.PivotSelect "VarUnit['(blank)']", xlDataAndLabel, True
    pt.PivotSelect "VarUnit", xlDataAndLabel, True

    On Error GoTo 0
'ESTIMATE 1
'Estimate 1 Total
    pt.AddDataField pt.PivotFields("Est1_Total"), "Sum of Est1_Total", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of Est1_Total")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
    
'ESTIMATE 2
'Estimate 2 Total
    pt.AddDataField pt.PivotFields("Est2_Total"), "Sum of Est2_Total", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of Est2_Total")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
    
'VARIANCE
'Total Variance
    pt.AddDataField pt.PivotFields("VarTotal"), "Sum of VarTotal", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of VarTotal")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
    
'Format Report Levels
    For Y = 1 To iLvl
        Select Case Y
            Case 1
                Call FrmtLvl1VAR
            Case 2
                Call FrmtLvl2VAR
            Case 3
                Call FrmtLvl3VAR
            Case 4
                Call FrmtLvl4VAR
            Case 5
                Call FrmtLvl5VAR
       End Select
    Next Y
    Call FrmtGTRowVAR
    
    If bMarkups = True Then Call Addons_VAR
    Call SetSheetHeadingsVAR
    Call SetVarBorders
    ows.Range("A1").Select
    Application.ScreenUpdating = True
    iLvl = 0
    pic = "DPRLogo.25.png"
    Call SheetFormatting
    Call PageSetup
    Call ResetSheetScroll
    bPvt = False
    On Error GoTo 0
    Set ptCache.Recordset = Nothing
    Set ptCache = Nothing
    Set pt = Nothing
    Set rsNew = Nothing
End Sub

Private Sub FrmtLvl1VAR()
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

Private Sub FrmtLvl2VAR()
''Format Level 2
    Columns("D:D").EntireColumn.Hidden = True
    Columns("E:E").ColumnWidth = colW
    pt.PivotSelect "'" & sLvl2Item & "'[All]", xlLabelOnly + xlFirstRow, True
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
    pt.PivotSelect "'" & sLvl2Item & "'[All;Total]", xlDataOnly + xlFirstRow, True
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

Private Sub FrmtLvl3VAR()
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

Private Sub FrmtLvl4VAR()
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

Private Sub FrmtLvl5VAR()
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

Sub SetVarBorders()

    pt.PivotSelect "'Sum of Est1_Total'", xlDataAndLabel, True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    
    pt.PivotSelect "'Sum of VarTotal'", xlDataAndLabel, True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    
    r = 0
    With ows.UsedRange
        r = .Cells(.Cells.count).row
        lngCol = .Cells(.Cells.count).Column
    End With
    
    With ows.Range(Cells(13, lngCol), Cells(r, lngCol))
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
    With .Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With .Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With .Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
End Sub



Sub FrmtGTRowVAR()
    If bMarkups = True Then
        pt.GrandTotalName = StrConv(sGTLvl1 & " Subtotal", vbUpperCase)
    Else
        pt.GrandTotalName = StrConv("Total", vbUpperCase)
    End If
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

Sub SetSheetHeadingsVAR()
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
    With ows.Range(Cells(1, 2), Cells(1, iCol + 1))
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
    ows.Cells(r + iLvl, iCol - 11).FormulaR1C1 = "DESCRIPTION"
    ows.Cells(r, iCol - 10).FormulaR1C1 = "ESTIMATE 1"
    ows.Cells(r + iLvl, iCol - 10).FormulaR1C1 = "QTY"
    ows.Cells(r + iLvl, iCol - 9).FormulaR1C1 = "UNIT"
    ows.Cells(r + iLvl, iCol - 8).FormulaR1C1 = "UNIT COST"
    ows.Cells(r, iCol - 7).FormulaR1C1 = "ESTIMATE 2"
    ows.Cells(r + iLvl, iCol - 7).FormulaR1C1 = "QTY"
    ows.Cells(r + iLvl, iCol - 6).FormulaR1C1 = "UNIT"
    ows.Cells(r + iLvl, iCol - 5).FormulaR1C1 = "UNIT COST"
    ows.Cells(r, iCol - 4).FormulaR1C1 = "VARIANCE"
    ows.Cells(r + iLvl, iCol - 4).FormulaR1C1 = "QTY VAR"
    ows.Cells(r + iLvl, iCol - 3).FormulaR1C1 = "COST VAR"
    ows.Cells(r + iLvl, iCol - 2).FormulaR1C1 = "EST-1 TOTAL"
    ows.Cells(r + iLvl, iCol - 1).FormulaR1C1 = "EST-2 TOTAL"
    ows.Cells(r + iLvl, iCol).FormulaR1C1 = "TOTAL VAR"
    ows.Cells(r, iCol + 1).FormulaR1C1 = "COMMENTS"
    ows.Range(Cells(r + iLvl, iCol - 11), Cells(r + iLvl, iCol)).HorizontalAlignment = xlCenter
    ows.Rows(r + iLvl + 1 & ":13").EntireRow.Hidden = True
    
    With ows.Range(Cells(r, 2), Cells(r + iLvl, iCol + 1))
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
        .Font.Name = "FrnkGothITC Bk BT"
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    With ows.Range(Cells(r, iCol - 10), Cells(r, iCol - 8))
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
        .Font.Name = "FrnkGothITC Bk BT"
        .Font.Size = 12
        .Font.Bold = True
        .HorizontalAlignment = xlCenterAcrossSelection
    End With
    
    With ows.Range(Cells(r, iCol - 7), Cells(r, iCol - 5))
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
        .Font.Name = "FrnkGothITC Bk BT"
        .Font.Size = 12
        .Font.Bold = True
        .HorizontalAlignment = xlCenterAcrossSelection
    End With
    
    With ows.Range(Cells(r, iCol - 10), Cells(r + iLvl, iCol - 8))
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.349986266670736
            .Weight = xlThin
        End With
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.349986266670736
            .Weight = xlThin
        End With
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    
    With ows.Range(Cells(r, iCol - 7), Cells(r + iLvl, iCol - 5))
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.349986266670736
            .Weight = xlThin
        End With
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.349986266670736
            .Weight = xlThin
        End With
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    
    With ows.Range(Cells(r, iCol - 4), Cells(r + iLvl, iCol))
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.349986266670736
            .Weight = xlThin
        End With
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.349986266670736
            .Weight = xlThin
        End With
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    
    With ows.Range(Cells(r, iCol - 4), Cells(r, iCol))
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorAccent1
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
        .Font.Name = "FrnkGothITC Bk BT"
        .Font.Size = 12
        .Font.Bold = True
        .HorizontalAlignment = xlCenterAcrossSelection
    End With
    
    ows.Range(Cells(r, iCol + 1), Cells(r, iCol + 1)).HorizontalAlignment = xlCenter
    
    ows.Rows("2:6").RowHeight = 17.55
    ows.Rows(r + iLvl).RowHeight = 18
    ows.Columns(iCol - 10).ColumnWidth = 11     'Est1 Qty
    ows.Columns(iCol - 9).ColumnWidth = 7       'Est1 UM
    ows.Columns(iCol - 8).ColumnWidth = 18      'Est1 Unit
    ows.Columns(iCol - 7).ColumnWidth = 11      'Est2 Qty
    ows.Columns(iCol - 6).ColumnWidth = 7       'Est2 UM
    ows.Columns(iCol - 5).ColumnWidth = 18      'Est2 Unit
    ows.Columns(iCol - 4).ColumnWidth = 18      'Qty Var
    ows.Columns(iCol - 3).ColumnWidth = 18      'Unit Var
    ows.Columns(iCol - 2).ColumnWidth = 18      'Est 1 Total
    ows.Columns(iCol - 1).ColumnWidth = 18      'Est 2 Total
    ows.Columns(iCol).ColumnWidth = 18          'Var Total
    ows.Columns(iCol + 1).ColumnWidth = 40      'Comments
    
    ows.Columns(iCol - 10).Hidden = True        'Est1 Qty
    ows.Columns(iCol - 9).Hidden = True         'Est1 UM
    ows.Columns(iCol - 8).Hidden = True         'Est1 Unit
    ows.Columns(iCol - 7).Hidden = True         'Est2 Qty
    ows.Columns(iCol - 6).Hidden = True         'Est2 UM
    ows.Columns(iCol - 5).Hidden = True         'Est2 Unit
    ows.Columns(iCol - 4).Hidden = True         'Qty Var
    ows.Columns(iCol - 3).Hidden = True         'Unit Var
    ows.Columns(iCol + 1).Hidden = True         'Comments
    
    

    Sheets("EstData").Shapes("grpHeadingVar").Copy
    Application.Goto Sheets(sSht).Range("B1")
    ActiveSheet.Paste
    Set myShape = ows.Shapes("grpHeadingVar")
    Set cl = Range(Cells(1, 2), Cells(6, iCol))
    With myShape
        .Left = cl.Left
        .Top = cl.Top
        .Height = cl.Height
        .Width = cl.Width
    End With
    With myShape.GroupItems
        .Item("txtSubHeading5").TextFrame.Characters.text = Range("=varRptEst2").Value
        '.Item("txtSubHeading6").TextFrame.Characters.Text = _
            "DURATION: " & StrConv(sDuration, vbUpperCase)
    End With
    ows.Range("A1").Select
End Sub


