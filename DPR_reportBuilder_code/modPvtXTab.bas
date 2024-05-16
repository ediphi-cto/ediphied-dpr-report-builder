Attribute VB_Name = "modPvtXTab"
Option Explicit
Public Sub Create_PivotTable_ODBC_XT()
Application.ScreenUpdating = False
    bPvt = True
    sVal3 = "Cost/" & Range("rngJobUnitName").Value & " "
    sJobUM = Range("rngJobUnitName").Value

    Set ptCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, Version:=xlPivotTableVersion15)
    Set ptCache.Recordset = rsNew
    ActiveWorkbook.Sheets.Add(Before:=Sheet4).Name = sSht
    Set ows = ActiveSheet
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    Set pt = ptCache.CreatePivotTable(TableDestination:=ows.Range("B9"), TableName:=sSht)
    With pt
        .TableStyle2 = "CrossTabReport_1"
        .HasAutoFormat = False
        .DisplayErrorString = True
        .ErrorString = "0"
        .NullString = "0"
        .ShowDrillIndicators = False
        .TableRange1.Font.Size = 12
        .TableRange1.Font.Name = "Franklin Gothic Book"
        .RepeatItemsOnEachPrintedPage = False
    End With
    x = 1
    'Build Levels
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
    'Set Values Area
    'On Error Resume Next
    pt.AddDataField pt.PivotFields("Amount"), "Sum of Amount", xlSum
    With pt.PivotFields("Sum of Amount")
        .Caption = "Amount "
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
    
    pt.AddDataField pt.PivotFields("Cost/Unit"), "Sum of Cost/Unit", xlSum
    With pt.PivotFields("Sum of Cost/Unit")
        .Caption = "Cost/Unit "
        .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
    End With

    pt.AddDataField pt.PivotFields("Cost/SF"), "Sum of Cost/SF", xlSum
    With pt.PivotFields("Sum of Cost/SF")
        .Caption = sVal3
        .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
    End With

'Set Column Data
    With pt.PivotFields(sLvl0Item)
        .Orientation = xlColumnField
        .Position = 1
    End With
    With pt.PivotFields(sLvl0Item)
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .LayoutForm = xlTabular
    End With
    With pt.PivotFields("LevelQuantity")
        .Orientation = xlColumnField
        .Position = 2
    End With
    With pt.PivotFields("LevelQuantity")
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .LayoutForm = xlTabular
    End With
    
    Call FormatXTab
    Call FormatXGrandTotal
    Call XTabHeadings
    
    For Y = 1 To iLvl
        Select Case Y
            Case 1
                Call FormatXLevel1
            Case 2
                Call FormatXLevel2
            Case 3
                Call FormatXLevel3
            Case 4
                Call FormatXLevel4
            Case 5
                Call FormatXLevel5
       End Select
    Next Y
    
    Call SetSheetHeadings
'Page Formatting for Printing
    With pt.TableRange1
        iGTRow = .Cells(.Cells.count).row
        iCol = .Cells(.Cells.count).Column
    End With
    ActiveSheet.PageSetup.PrintArea = Range(Cells(1, 2), Cells(iGTRow, iCol)).Address
    ActiveSheet.PageSetup.PrintTitleRows = "$1:$10"

    ows.Range("B12").HorizontalAlignment = xlLeft
    ows.Range(Cells(10, 3), Cells(12, z - 1)).Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Franklin Gothic Book"
        .Size = 12
        .Bold = True
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    iLvl = 0
    Call XTabAddons
    pic = "DPRLogo.25.png"
    Call PageSetup
    Call ResetSheetScroll
    sLvl1Item = ""
    sLvl2Item = ""
    sLvl3Item = ""
    sLvl4Item = ""
    sLvl5Item = ""
    bPvt = False
    Set ptCache.Recordset = Nothing
    Set ptCache = Nothing
    Set pt = Nothing
    Set rsNew = Nothing
End Sub

Sub FormatXTab()
'Format Data Columns
    'On Error Resume Next
    pt.PivotSelect "'" & sVal3 & "'", xlDataOnly, True
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    Selection.HorizontalAlignment = xlRight
    With Selection.Font
        .Name = "Franklin Gothic Book"
        .Size = 12
        .Bold = False
        .Underline = xlUnderlineStyleNone
        .Color = -16777216
        .TintAndShade = 0
    End With
    
    pt.PivotSelect "'Amount '", xlDataOnly, True
    z = Selection.Column
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Font
        .Name = "Franklin Gothic Book"
        .Size = 12
        .Bold = False
        .Underline = xlUnderlineStyleNone
        .Color = -16777216
        .TintAndShade = 0
    End With
    Selection.HorizontalAlignment = xlRight
    
    pt.PivotSelect "'Cost/Unit '", xlDataOnly, True
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Font
        .Name = "Franklin Gothic Book"
        .Size = 12
        .Bold = False
        .Underline = xlUnderlineStyleNone
        .Color = -16777216
        .TintAndShade = 0
    End With
    Selection.HorizontalAlignment = xlRight
    On Error GoTo 0
End Sub

Sub XTabHeadings()
'Format Column Header 1
    'On Error Resume Next
    pt.PivotSelect "'" & sLvl0Item & "'", xlLabelOnly, True
    With Selection.Font
        .Name = "Franklin Gothic Book"
        .Size = 12
        .Bold = True
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.HorizontalAlignment = xlCenterAcrossSelection
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With
'Format Column Header 2
    pt.PivotSelect "'LevelQuantity'", xlLabelOnly, True
    With Selection.Font
        .Name = "Franklin Gothic Book"
        .Size = 12
        .Bold = True
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.HorizontalAlignment = xlCenterAcrossSelection
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With
'Format Column Header 3
    pt.PivotSelect "'Amount '", xlLabelOnly, True
    With Selection.Font
        .Name = "Franklin Gothic Book"
        .Size = 12
        .Bold = True
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.HorizontalAlignment = xlCenter
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.149937437055574
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With
'Format Column Header 4
    pt.PivotSelect "'Cost/Unit '", xlLabelOnly, True
    With Selection.Font
        .Name = "Franklin Gothic Book"
        .Size = 12
        .Bold = True
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.HorizontalAlignment = xlCenter
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.149937437055574
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With

'Format Column Header 5
    pt.PivotSelect "'" & sVal3 & "'", xlLabelOnly, True
    With Selection.Font
        .Name = "Franklin Gothic Book"
        .Size = 12
        .Bold = True
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.HorizontalAlignment = xlCenter
'Format Total Column Header 1
    pt.PivotSelect "'Amount ' 'Row Grand Total'", xlLabelOnly, True
    With Selection.Font
        .Name = "Franklin Gothic Book"
        .Size = 12
        .Bold = True
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With
'Format Total Column Header 2
    pt.PivotSelect "'" & sVal3 & "' 'Row Grand Total'", xlLabelOnly, True
    With Selection.Font
        .Name = "Franklin Gothic Book"
        .Size = 12
        .Bold = True
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    pt.PivotSelect "'" & sVal3 & "'", xlDataOnly, True
    Selection.EntireColumn.Hidden = True
    pt.PivotSelect "'Amount '", xlDataAndLabel, True
    Selection.ColumnWidth = 16.25
    pt.PivotSelect "'Cost/Unit '", xlDataOnly, True
    Selection.ColumnWidth = 16.25
    pt.PivotSelect "'Cost/Unit ' 'Row Grand Total'", xlDataAndLabel, True
    Selection.EntireColumn.Hidden = True
    pt.PivotSelect "'" & sVal3 & "' 'Row Grand Total'", xlDataAndLabel, True
    Selection.ColumnWidth = 16
    
End Sub

Sub FormatXLevel1()
''Format Level 1
    pt.PivotSelect "'" & sLvl1Item & "'", xlLabelOnly + xlFirstRow, True
    Columns("B:B").ColumnWidth = 0.05
    If iLvl > 1 Then
        Selection.ColumnWidth = 0.05
    Else
        Selection.ColumnWidth = 45
    End If
    With Selection.Font
        .Name = "Franklin Gothic Book"
        .Size = 12
        .Bold = True
        .Underline = xlUnderlineStyleNone
        .Color = -16777216
        .TintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With
    If iLvl <> 1 Then
        pt.PivotSelect "'" & sLvl1Item & "'", xlDataAndLabel + xlFirstRow, True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight2
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        Selection.Borders(xlInsideVertical).LineStyle = xlNone
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    End If
''Format Level - 1 Subtotals
    If iLvl = 1 Then Exit Sub
    pt.PivotSelect "'" & sLvl1Item & "'[All;Total]", xlDataAndLabel + xlFirstRow, True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Name = "Franklin Gothic Book"
        .Size = 12
        .Bold = True
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
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
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With
    pt.PivotSelect "'" & sLvl1Item & "'[All;'Blank']", xlDataAndLabel + xlFirstRow, True
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

Sub FormatXLevel2()
''Format Level - 2

    pt.PivotSelect "'" & sLvl2Item & "'[All]", xlLabelOnly + xlFirstRow, True
    Columns("D:D").EntireColumn.Hidden = True
    If iLvl > 2 Then
        Selection.ColumnWidth = 0.05
    Else
        Selection.ColumnWidth = 45
    End If
    Selection.InsertIndent 1
    If iLvl = 2 Then
        With Selection.Font
            .Name = "Franklin Gothic Book"
            .Size = 12
            .Bold = False
            .Color = -16777216
            .TintAndShade = 0
        End With
    Else
        With Selection.Font
            .Name = "Franklin Gothic Book"
            .Size = 12
            .Bold = True
            .Color = -16777216
            .TintAndShade = 0
        End With
    End If
    pt.PivotSelect "'" & sLvl2Item & "'[All]", xlDataAndLabel + xlFirstRow, True
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.14996795556505
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.14996795556505
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    pt.PivotSelect "'" & sLvl2Item & "'[All]", xlLabelOnly + xlFirstRow, True
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 9.99481185338908E-02
        .Weight = xlThin
    End With
''Format Level - 2 Subtotals
    If iLvl = 2 Then Exit Sub
    pt.PivotSelect "'" & sLvl2Item & "'[All;Total]", xlLabelOnly + xlFirstRow, True
    Selection.InsertIndent 1
    pt.PivotSelect "'" & sLvl2Item & "'[All;Total]", xlDataAndLabel + xlFirstRow, True
    With Selection.Font
        .Name = "Franklin Gothic Book"
        .Size = 12
        .Bold = True
        .Color = -16777216
        .TintAndShade = 0
    End With
    pt.PivotSelect "'" & sLvl2Item & "'[All;Total]", xlDataOnly + xlFirstRow, True
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

Sub FormatXLevel3()
''Format Level 3
    Columns("F:F").EntireColumn.Hidden = True
    If iLvl > 3 Then
        Columns("G:G").ColumnWidth = 0.05
    Else
        Columns("G:G").ColumnWidth = 55
    End If
    pt.PivotSelect "'" & sLvl3Item & "'[All]", xlLabelOnly + xlFirstRow, True
    Columns("F:F").EntireColumn.Hidden = True
    If iLvl > 3 Then
        Selection.ColumnWidth = 0.05
    Else
        Selection.ColumnWidth = 55
    End If
    Selection.InsertIndent 2
    If iLvl = 3 Then
        With Selection.Font
            .Name = "Franklin Gothic Book"
            .Size = 12
            .Bold = False
            .Color = -16777216
            .TintAndShade = 0
        End With
    Else
        With Selection.Font
            .Name = "Franklin Gothic Book"
            .Size = 12
            .Bold = True
            .Color = -16777216
            .TintAndShade = 0
        End With
    End If
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
''Format Level - 3 Subtotals
    If iLvl = 3 Then Exit Sub
    pt.PivotSelect "'" & sLvl3Item & "'[All;Total]", xlLabelOnly + xlFirstRow, True
    Selection.InsertIndent 2
    pt.PivotSelect "'" & sLvl3Item & "'[All;Total]", xlDataAndLabel + xlFirstRow, True
    With Selection.Font
        .Name = "Franklin Gothic Book"
        .Size = 12
        .Bold = True
        .Color = -16777216
        .TintAndShade = 0
    End With
    pt.PivotSelect "'" & sLvl3Item & "'[All;Total]", xlDataOnly + xlFirstRow, True
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

Sub FormatXLevel4()
''Format Level 4
    pt.PivotSelect "'" & sLvl4Item & "'[All]", xlLabelOnly + xlFirstRow, True
    Columns("H:H").EntireColumn.Hidden = True
    If iLvl > 4 Then
        Selection.ColumnWidth = 0.05
    Else
        Selection.ColumnWidth = 55
    End If
    Selection.InsertIndent 3
    If iLvl = 4 Then
        With Selection.Font
            .Name = "Franklin Gothic Book"
            .Size = 12
            .Bold = False
            .Color = -16777216
            .TintAndShade = 0
        End With
    Else
        With Selection.Font
            .Name = "Franklin Gothic Book"
            .Size = 12
            .Bold = True
            .Color = -16777216
            .TintAndShade = 0
        End With
    End If
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With

''Format Level - 4 Subtotals
    If iLvl = 4 Then Exit Sub
    pt.PivotSelect "'" & sLvl4Item & "'[All;Total]", xlLabelOnly + xlFirstRow, True
    Selection.InsertIndent 3
    pt.PivotSelect "'" & sLvl4Item & "'[All;Total]", xlDataAndLabel + xlFirstRow, True
    With Selection.Font
        .Name = "Franklin Gothic Book"
        .Size = 12
        .Bold = True
        .Color = -16777216
        .TintAndShade = 0
    End With
    pt.PivotSelect "'" & sLvl4Item & "'[All;Total]", xlDataOnly + xlFirstRow, True
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

Sub FormatXLevel5()
''Format Level 5
    pt.PivotSelect "'" & sLvl5Item & "'[All]", xlLabelOnly + xlFirstRow, True
    Columns("J:J").EntireColumn.Hidden = True
    Selection.ColumnWidth = 55
    Selection.InsertIndent 4
    If iLvl = 5 Then
        With Selection.Font
            .Name = "Franklin Gothic Book"
            .Size = 12
            .Bold = False
            .Color = -16777216
            .TintAndShade = 0
        End With
    Else
        With Selection.Font
            .Name = "Franklin Gothic Book"
            .Size = 10
            .Bold = True
            .Color = -16777216
            .TintAndShade = 0
        End With
    End If
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
''Format Level - 5 Subtotals
    If iLvl = 5 Then Exit Sub
    pt.PivotSelect "'" & sLvl5Item & "'[All;Total]", xlLabelOnly + xlFirstRow, True
    Selection.InsertIndent 4
    pt.PivotSelect "'" & sLvl5Item & "'[All;Total]", xlDataAndLabel + xlFirstRow, True
    With Selection.Font
        .Name = "Franklin Gothic Book"
        .Size = 12
        .Bold = True
        .Color = -16777216
        .TintAndShade = 0
    End With
    pt.PivotSelect "'" & sLvl5Item & "'[All;Total]", xlDataOnly + xlFirstRow, True
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
Sub FormatXGrandTotal()
''Format Grand Total
    'On Error Resume Next
    pt.GrandTotalName = "SUB TOTAL"
    pt.PivotSelect "'Column Grand Total'", xlDataAndLabel + xlFirstRow, True
        With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Name = "Franklin Gothic Book"
        .Size = 12
        .Bold = True
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
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
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -4.99893185216834E-02
        .Weight = xlThin
    End With
    On Error GoTo 0
End Sub

Sub SetSheetHeadings()
Dim sLeft As Single
Dim shpName As String

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
    
    ows.Range("A9").EntireRow.Hidden = True
    Sheets("EstData").Shapes("grpHeading").Copy
    Application.Goto Sheets(sSht).Range("B1")
    ActiveSheet.Paste
    Set myShape = ows.Shapes("grpHeading")
    Set cl = Range(Cells(1, 2), Cells(8, iCol))
    With myShape
        .Left = cl.Left
        .Top = cl.Top
        .Height = cl.Height
        .Width = cl.Width
    End With
End Sub






