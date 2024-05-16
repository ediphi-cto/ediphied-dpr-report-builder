Attribute VB_Name = "modPvtLevel"
Option Explicit

Public Sub Create_PivotTable_ODBC_MO()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    bPvt = True
    
    Dim shtName As String
    If isFirstReport() Then
        shtName = "Level Report"
    Else
        shtName = "Detailed Backup"
    End If
    
    sSht = GetUniqueSheetName(shtName)
    Set ptCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, SourceData:="tblEdiphiPivotData", Version:=xlPivotTableVersion15)
    ActiveWorkbook.Sheets.Add(Before:=Sheet4).Name = sSht
    Set ows = ActiveSheet
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    Set pt = ptCache.CreatePivotTable(TableDestination:=ows.Range("B13"), TableName:=sSht)
    With pt
        .TableStyle2 = "DPR_Estimating_Style_01"
        .HasAutoFormat = False
        .DisplayErrorString = True
        .ErrorString = "0"
        .NullString = "~"
        .ShowDrillIndicators = False
        .TableRange1.Font.Size = 12
        .TableRange1.Font.Name = "Franklin Gothic Book"
        .TableRange1.VerticalAlignment = xlTop
        .RepeatItemsOnEachPrintedPage = False
        .ManualUpdate = True
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
                'MN TODO: this assumes a fixed column width, but we have use groups of varying count
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
                'MN:  len of Array(False...) needs to be dynamically set bc of use group columns
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
'Field Unit Price
    x = x + 1
    With pt.PivotFields("UnitPrice")
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
'Field TOCost
    Set pf = pt.PivotFields("UnitPrice")
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
    'pt.PivotFields("TakeoffQty").PivotItems("(blank)").Caption = "-"
    'pt.PivotSelect "TakeoffQty['(blank)']", xlDataAndLabel, True
    pt.PivotSelect "TakeoffQty", xlDataAndLabel, True
    
    pt.PivotSelect "'TakeoffUnit'[All]", xlLabelOnly + xlFirstRow, True
    With Selection
        .ColumnWidth = 11
        .HorizontalAlignment = xlCenter
    End With
    
    
    pt.PivotSelect "'UnitPrice'[All]", xlLabelOnly + xlFirstRow, True
    With Selection
        .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
        .ColumnWidth = 16
        .HorizontalAlignment = xlRight
    End With
    'pt.PivotFields("UnitPrice").PivotItems("(blank)").Caption = "-"
    'pt.PivotSelect "UnitPrice['(blank)']", xlDataAndLabel, True
    pt.PivotSelect "UnitPrice", xlDataAndLabel, True
    On Error GoTo 0
'Gross Total
    pt.AddDataField pt.PivotFields("GrandTotal"), "Sum of GrandTotal", xlSum
    On Error Resume Next
    With pt.PivotFields("Sum of GrandTotal")
        .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    End With
'Format Report Levels
    For Y = 1 To iLvl
        Select Case Y
            Case 1
                Call FrmtLvl1
            Case 2
                Call FrmtLvl2
            Case 3
                Call FrmtLvl3
            Case 4
                Call FrmtLvl4
            Case 5
                Call FrmtLvl5
       End Select
    Next Y
    Call FrmtGTRow
    Call Addons
    Call SetSheetHeadings
    ows.Range("A1").Select
    Application.ScreenUpdating = True
  
    pic = "DPRLogo.25.png"
    Call PageSetup
    Call ResetSheetScroll
    bPvt = False
    Set ptCache.Recordset = Nothing
    Set ptCache = Nothing
    Set pt = Nothing
    Set rsNew = Nothing
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    
    On Error GoTo 0
End Sub

Private Sub FrmtLvl1()
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

Private Sub FrmtLvl2()
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

Private Sub FrmtLvl3()
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

Private Sub FrmtLvl4()
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

Private Sub FrmtLvl5()
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

Sub FrmtGTRow()
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

Sub SetSheetHeadings()
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
    ows.Cells(r + iLvl, iCol - 5).FormulaR1C1 = "DESCRIPTION"
    ows.Cells(r + iLvl, iCol - 4).FormulaR1C1 = "COMMENTS"
    ows.Cells(r + iLvl, iCol - 3).FormulaR1C1 = "QUANTITY"
    ows.Cells(r + iLvl, iCol - 2).FormulaR1C1 = "UNIT"
    ows.Cells(r + iLvl, iCol - 1).FormulaR1C1 = "UNIT COST"
    ows.Cells(r + iLvl, iCol).FormulaR1C1 = "TOTAL"
    ows.Range(Cells(r + iLvl, iCol - 5), Cells(r + iLvl, iCol)).HorizontalAlignment = xlCenter
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
        .Font.Bold = True
    End With
    
    ows.Rows("2:6").RowHeight = 17.55
    ows.Rows(r + iLvl).RowHeight = 18
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

