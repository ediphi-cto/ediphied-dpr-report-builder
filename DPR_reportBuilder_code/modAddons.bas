Attribute VB_Name = "modAddons"
Option Explicit

Sub Addons()
    i = 0
    sJobUM = Range("rngJobUnitName").Value
    If ows.PivotTables.Count > 0 Then
        If pt = "" Then
            Set pt = ows.PivotTables(1)
        End If
        With pt.TableRange1
            iGTRow = .Cells(.Cells.Count).row
            iGTCol = .Cells(.Cells.Count).Column
        End With
    Else
        iGTRow = ows.Range("SysEnd").row
        iGTCol = ows.Range("H1").Column
    End If
    Application.ScreenUpdating = False
    iAddRow = ActualUsedRange(ows).Rows.Count
    If iAddRow > iGTRow Then
        Exit Sub
    End If
    
    Set lObj = Sheet0.ListObjects("tblTotals")
        If lObj.ListRows.Count < 1 Then Exit Sub
        i = iGTRow + 2
        X = iGTCol
        For r = 1 To lObj.DataBodyRange.Rows.Count
            If UCase(lObj.DataBodyRange(r, 7)) = "UPPER" Then
                With ows.Cells(i, 3)
                    .Value = lObj.DataBodyRange(r, 4).Value
                    .InsertIndent 2
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                If lObj.DataBodyRange(r, 5) <> "" Then
                    With ows.Cells(i, X - 3)
                        .Style = "Percent"
                        .FormulaR1C1 = lObj.DataBodyRange(r, 5) / 100
                        .NumberFormat = "0.00%"
                        .Font.Size = 12
                        .Font.Color = -16777216
                    End With
                End If
                With ows.Cells(i, X)
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .FormulaR1C1 = lObj.DataBodyRange(r, 6)
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, X - 1)
                    .FormulaR1C1 = "=IFERROR(RC[1]/rngJobSize,0)"
                    .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                i = i + 1
            End If
        Next r
        
        With ows.Cells(i, 2)
            .Value = StrConv("Projected Construction Costs", vbUpperCase)
            .Font.Size = 12
            .Font.Bold = True
            .Font.ThemeColor = xlThemeColorDark1
            .Font.TintAndShade = 0
        End With
        With ows.Cells(i, X - 1)
            .FormulaR1C1 = "=IFERROR(RC[1]/rngJobSize,0)"
            .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
            .Font.Size = 12
            .Font.Bold = True
            .Font.ThemeColor = xlThemeColorDark1
            .Font.TintAndShade = 0
        End With
        With ows.Cells(i, X)
            Y = iGTRow - i
            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
            .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
            .Font.Size = 12
            .Font.Bold = True
            .Font.ThemeColor = xlThemeColorDark1
            .Font.TintAndShade = 0
        End With
        With ows.Range(ows.Cells(i, 2), ows.Cells(i, iGTCol)).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
'Set Allocate False
    r = 0
    Set lObj = Sheet0.ListObjects("tblTotals")
        Y = i
        i = i + 2
        X = iGTCol
        For r = 1 To lObj.DataBodyRange.Rows.Count
            If lObj.DataBodyRange(r, 7) = "Lower" Then
                With ows.Cells(i, 3)
                    .Value = lObj.DataBodyRange(r, 4).Value
                    .InsertIndent 2
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                If lObj.DataBodyRange(r, 5) <> "" Then
                    With ows.Cells(i, X - 3)
                        .Style = "Percent"
                        .FormulaR1C1 = lObj.DataBodyRange(r, 5) / 100
                        .NumberFormat = "0.00%"
                        .Font.Size = 12
                        .Font.Color = -16777216
                    End With
                End If
                With ows.Cells(i, X)
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    If lObj.DataBodyRange(r, 6).Value <> "" Then
                        .FormulaR1C1 = lObj.DataBodyRange(r, 6)
                    Else
                        .FormulaR1C1 = 0
                    End If
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, X - 1)
                    .FormulaR1C1 = "=IFERROR(RC[1]/rngJobSize,0)"
                    .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                i = i + 1
            End If
        Next r
        
        With ows.Cells(i, 2)
            .Value = StrConv("TOTAL", vbUpperCase)
            .Font.Size = 12
            .Font.Bold = True
            .Font.ThemeColor = xlThemeColorDark1
            .Font.TintAndShade = 0
        End With
        With ows.Cells(i, X - 1)
            .FormulaR1C1 = "=IFERROR(RC[1]/rngJobSize,0)"
            .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
            .Font.Size = 12
            .Font.Bold = True
            .Font.ThemeColor = xlThemeColorDark1
            .Font.TintAndShade = 0
        End With
        With ows.Cells(i, X)
            Y = Y - i
            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
            .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
            .Font.Size = 12
            .Font.Bold = True
            .Font.ThemeColor = xlThemeColorDark1
            .Font.TintAndShade = 0
        End With
        With ows.Range(ows.Cells(i, 2), ows.Cells(i, iGTCol)).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

    Application.ScreenUpdating = True
    Call SheetFormatting
End Sub

Sub XTabAddons()
Dim dTotal As Double
Dim C
    sJobUM = Sheet1.Range("rngJobUnitName").Value
    dTotal = Sheet3.Range("SysEnd").Offset(0, 6).Value
    i = 0
    If ows.name = "" Then
        Set ows = ActiveSheet
    End If
'Projected Construction Costs
    Application.ScreenUpdating = False
        If ows.PivotTables.Count > 0 Then
            If pt = "" Then
                Set pt = ActiveSheet.PivotTables(1)
            End If
            With pt.TableRange1
                iGTRow = .Cells(.Cells.Count).row
                iGTCol = .Cells(.Cells.Count).Column
            End With
        Else
            Exit Sub
        End If
        For Each pf In pt.ColumnFields
            If pf.Position = 1 Then
                sLvl1Item = pf.SourceName
                z = pf.dataRange.Column
                Exit For
            End If
        Next pf
        
        iAddRow = ActualUsedRange(ows).Rows.Count
        If iAddRow <> iGTRow Then
            ows.Range(ows.Cells(iGTRow, 2).Offset(1, 0), ows.Cells(iAddRow, iGTCol)).EntireRow.Delete
        End If
        Set lObj = Sheet0.ListObjects("tblTotals")
        If lObj.ListRows.Count < 1 Then Exit Sub
        i = iGTRow + 2
        For r = 1 To lObj.DataBodyRange.Rows.Count
            If lObj.DataBodyRange(r, 7) = "Upper" Then
                With ows.Cells(i, 2)
                    .Value = lObj.DataBodyRange(r, 4).Value
                    .Font.Size = 12
                    .Font.Bold = True
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 2)
                    .FormulaR1C1 = lObj.DataBodyRange(r, 6).Value
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol)
                    .FormulaR1C1 = "=IFERROR(RC[-2]/rngJobSize,0)"
                    .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                For C = z To iGTCol - 3 Step 3
                    Y = iGTRow - i
                    With ows.Cells(i, C)
                        If lObj.DataBodyRange(r, 6).Value <> "" Then
                            .FormulaR1C1 = "=IFERROR(SUM(" & lObj.DataBodyRange(r, 6).Value & " *(R[" & Y & "]C/" & dTotal & ")),0)"
                        Else
                            .FormulaR1C1 = "=IFERROR(SUM(0*(R[" & Y & "]C/" & dTotal & ")),0)"
                        End If
                       .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                       .Font.Size = 12
                       .Font.Color = -16777216
                    End With
                    With ows.Cells(i, C).Offset(0, 1)
                        .FormulaR1C1 = "=IFERROR(SUM(RC[-1]/StripChar(R11C[-1])),0)"
                        .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
                        .Font.Size = 12
                        .Font.Color = -16777216
                    End With
                Next C
            i = i + 1
            End If
        Next r
'Projected Construction Costs Totals
        Y = iGTRow - i
        ows.Cells(i, 2).Value = "PROJECTED CONSTRUCTION COSTS"
        With ows.Range(ows.Cells(i, 2), ows.Cells(i, z - 1))
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0
            End With
            With .Font
                .Size = 12
                .Bold = True
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End With
        With ows.Cells(i, iGTCol - 2)
            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
            .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With .Font
                .Size = 12
                .Bold = True
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End With
        With ows.Cells(i, iGTCol)
            .FormulaR1C1 = "=IFERROR(RC[-2]/rngJobSize,0)"
            .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With .Font
                .Size = 12
                .Bold = True
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End With
        For C = z To iGTCol - 3 Step 3
            With ows.Cells(i, C)
                .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
                .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                With .Font
                    .Size = 12
                    .Bold = True
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
            End With
            With ows.Cells(i, C).Offset(0, 1)
                .FormulaR1C1 = "=IFERROR(SUM(RC[-1]/StripChar(R11C[-1])),0)"
                .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0
                End With
                With .Font
                    .Size = 12
                    .Bold = True
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
            End With
        Next C
        i = i + 1

'Other Project Costs
        Set lObj = Sheet0.ListObjects("tblTotals")
        X = i - 1
        i = i + 1
        For r = 1 To lObj.DataBodyRange.Rows.Count
            If lObj.DataBodyRange(r, 7) = "Lower" Then
        With ows.Cells(i, 2)
            .Value = lObj.DataBodyRange(r, 4).Value
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = -16777216
        End With
        With ows.Cells(i, iGTCol - 2)
            .FormulaR1C1 = lObj.DataBodyRange(r, 6).Value
            .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
            .Font.Size = 12
            .Font.Color = -16777216
        End With
        With ows.Cells(i, iGTCol)
            .FormulaR1C1 = "=IFERROR(RC[-2]/rngJobSize,0)"
            .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
            .Font.Size = 12
            .Font.Color = -16777216
        End With
        For C = z To iGTCol - 3 Step 3
            Y = iGTRow - i
            With ows.Cells(i, C)
                If lObj.DataBodyRange(r, 6).Value <> "" Then
                    .FormulaR1C1 = "=IFERROR(SUM(" & lObj.DataBodyRange(r, 6).Value & " *(R[" & Y & "]C/" & dTotal & ")),0)"
                Else
                    .FormulaR1C1 = "=IFERROR(SUM(0*(R[" & Y & "]C/" & dTotal & ")),0)"
                End If
                .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                .Font.Size = 12
                .Font.Color = -16777216
            End With
            With ows.Cells(i, C).Offset(0, 1)
                .FormulaR1C1 = "=IFERROR(SUM(RC[-1]/StripChar(R11C[-1])),0)"
                .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
                .Font.Size = 12
                .Font.Color = -16777216
            End With
        Next C
        i = i + 1
        End If
    Next r
        
'Other Project Costs Totals
        Y = X - i
        ows.Cells(i, 2).Value = "TOTAL"
        With ows.Range(ows.Cells(i, 2), ows.Cells(i, z - 1))
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0
            End With
            With .Font
                .Size = 12
                .Bold = True
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End With
        With ows.Cells(i, iGTCol - 2)
            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
            .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With .Font
                .Size = 12
                .Bold = True
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End With
        With ows.Cells(i, iGTCol)
            .FormulaR1C1 = "=IFERROR(RC[-2]/rngJobSize,0)"
            .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0
            End With
            With .Font
                .Size = 12
                .Bold = True
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End With
        For C = z To iGTCol - 3 Step 3
            With ows.Cells(i, C)
                .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
                .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0
                End With
                With .Font
                    .Size = 12
                    .Bold = True
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
            End With
            With ows.Cells(i, C).Offset(0, 1)
                .FormulaR1C1 = "=IFERROR(SUM(RC[-1]/StripChar(R11C[-1])),0)"
                .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0
                End With
                With .Font
                    .Size = 12
                    .Bold = True
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
            End With
        Next C
    
'Format Borders
    If i > 0 Then
        For C = z To iGTCol - 3 Step 3
            With ows.Range(ows.Cells(iGTRow, iGTCol), ows.Cells(i, iGTCol))
                With .Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
            End With
            With ows.Range(ows.Cells(iGTRow, iGTCol - 2), ows.Cells(i, iGTCol - 2))
                With .Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With .Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ThemeColor = 1
                    .TintAndShade = -0.349986266670736
                    .Weight = xlThin
                End With
            End With
            With ows.Range(ows.Cells(iGTRow, C), ows.Cells(i, C).Offset(0, 1))
                With .Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With .Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .ThemeColor = 1
                    .TintAndShade = -0.349986266670736
                    .Weight = xlThin
                End With
                With .Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
            End With
        Next C
        ActiveSheet.PageSetup.PrintArea = ows.Range(ows.Cells(1, 2), ows.Cells(i, iGTCol)).Address
        ActiveSheet.PageSetup.PrintTitleRows = "$1:$10"
    End If
    Application.ScreenUpdating = True
    SheetFormatting
End Sub

Sub clearAddons()
    If Not IsNull(ows) = True Then
        Set ows = ActiveSheet
    End If
    If ows.PivotTables.Count > 0 Then
        Set pt = ows.PivotTables(1)
        With pt.TableRange1
            iGTRow = .Cells(.Cells.Count).row
            iGTCol = .Cells(.Cells.Count).Column
        End With
     Else
        iGTRow = ows.Range("SysEnd").row
        iGTCol = ows.Range("H1").Column
     End If
    iAddRow = ActualUsedRange(ows).Rows.Count
    If iAddRow > iGTRow Then
        ows.Range(ows.Cells(iGTRow, 2).Offset(1, 0), ows.Cells(iAddRow, iGTCol)).EntireRow.Delete
    End If
    SheetFormatting
End Sub

'Sub CEstAddons() 'Control Estimate Addons for new codes
'Dim dTotal As Double
'    sJobUM = Sheet1.Range("rngJobUnitName").value
'    dTotal = Sheet3.Range("SysEnd").Offset(0, 6).value
'    i = 0
'    If ows.name = "" Then
'        Set ows = ActiveSheet
'    End If
''Projected Construction Costs
'    Application.ScreenUpdating = False
'        If ows.PivotTables.count > 0 Then
'            If pt = "" Then
'                Set pt = ActiveSheet.PivotTables(1)
'            End If
'            With pt.TableRange1
'                iGTRow = .Cells(.Cells.count).row
'                iGTCol = .Cells(.Cells.count).Column
'            End With
'        Else
'            Exit Sub
'        End If
'        For Each pf In pt.ColumnFields
'            If pf.Position = 1 Then
'                'sLvl1Item = pf.SourceName
'                z = pf.DataRange.Column
'                Exit For
'            End If
'        Next pf
'        iAddRow = ActualUsedRange(ows).Rows.count + 12
'        If iAddRow <> iGTRow Then
'            ows.Range(ows.Cells(iGTRow, 2).Offset(1, 0), ows.Cells(iAddRow, iGTCol)).EntireRow.Delete
'        End If
'        Set lObj = Sheet0.ListObjects("tblTotals")
'        If lObj.ListRows.count < 1 Then Exit Sub
'        i = iGTRow + 2
'        'Add Sort if by Phase Code
'        With lObj
'            .Sort.SortFields.Clear
'            .Sort.SortFields.Add key:=Range("tblTotals[JobCost]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'            .Sort.SortFields.Add key:=Range("tblTotals[SortOrder]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'            With .Sort
'                .Header = xlYes
'                .MatchCase = False
'                .Orientation = xlTopToBottom
'                .SortMethod = xlPinYin
'                .Apply
'            End With
'        End With
'        For r = 1 To lObj.DataBodyRange.Rows.count
'            If lObj.DataBodyRange(r, 7) = "Upper" Then
'                If ows.Cells(13, 3).value = "Job Cost" And lObj.DataBodyRange(r, 10).value <> lObj.DataBodyRange(r - 1, 10).value Then
'                    With ows.Cells(i, 2)
'                        .value = lObj.DataBodyRange(r, 10).value & "-" & lObj.DataBodyRange(r, 11).value
'                        .Font.Size = 12
'                        .Font.Bold = True
'                        .Font.Color = -16777216
'                    End With
'                    With ows.Range(ows.Cells(i, 2), ows.Cells(i, iGTCol))
'                         .Interior.Pattern = xlSolid
'                         .Interior.PatternColorIndex = xlAutomatic
'                         .Interior.ThemeColor = xlThemeColorDark2
'                         .Interior.TintAndShade = 0.799981688894314
'                         .Interior.PatternTintAndShade = 0
'                    End With
'                    i = i + 1
'                End If
'                With ows.Cells(i, 3)
'                    .value = lObj.DataBodyRange(r, 4).value
'                    .Font.Size = 12
'                    .InsertIndent 2
'                    .Font.Color = -16777216
'                End With
'                With ows.Cells(i, iGTCol - 12)
'                    .FormulaR1C1 = 0
'                    .NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
'                End With
'                With ows.Cells(i, iGTCol - 11)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],10)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                End With
'                With ows.Cells(i, iGTCol - 10)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],20)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
'                End With
'                With ows.Cells(i, iGTCol - 9)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],30)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                End With
'                With ows.Cells(i, iGTCol - 8)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],40)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
'                End With
'                With ows.Cells(i, iGTCol - 7)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],50)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                End With
'                With ows.Cells(i, iGTCol - 6)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],51)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
'                End With
'                With ows.Cells(i, iGTCol - 5)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],52)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                End With
'                With ows.Cells(i, iGTCol - 4)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],60)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
'                End With
'                With ows.Cells(i, iGTCol - 3)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],61)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                End With
'                With ows.Cells(i, iGTCol - 2)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],62)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
'                End With
'                With ows.Cells(i, iGTCol - 1)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],70)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                End With
'                With ows.Cells(i, iGTCol)
'                    .FormulaR1C1 = "=SUM(RC[-11]:RC[-1])"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
'                End With
'                i = i + 1
'            End If
'        Next r
''Projected Construction Costs Totals
'        Y = iGTRow - i
'        ows.Cells(i, 2).value = "PROJECTED CONSTRUCTION COSTS"
'        With ows.Range(ows.Cells(i, 2), ows.Cells(i, z))
'            With .Interior
'                .Pattern = xlSolid
'                .PatternColorIndex = xlAutomatic
'                .ThemeColor = xlThemeColorAccent1
'                .TintAndShade = 0
'            End With
'            With .Font
'                .Size = 12
'                .Bold = True
'                .ThemeColor = xlThemeColorDark1
'                .TintAndShade = 0
'            End With
'        End With
'        For C = 1 To 11
'            With ows.Cells(i, iGTCol - C)
'                .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
'                .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                With .Interior
'                    .Pattern = xlSolid
'                    .PatternColorIndex = xlAutomatic
'                    .ThemeColor = xlThemeColorAccent1
'                    .TintAndShade = 0
'                    .PatternTintAndShade = 0
'                End With
'                With .Font
'                    .Size = 12
'                    .Bold = True
'                    .ThemeColor = xlThemeColorDark1
'                    .TintAndShade = 0
'                End With
'            End With
'        Next C
'        With ows.Cells(i, iGTCol - 12)
'            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
'            .NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
'            With .Interior
'                .Pattern = xlSolid
'                .PatternColorIndex = xlAutomatic
'                .ThemeColor = xlThemeColorAccent1
'                .TintAndShade = 0
'                .PatternTintAndShade = 0
'            End With
'            With .Font
'                .Size = 12
'                .Bold = True
'                .ThemeColor = xlThemeColorDark1
'                .TintAndShade = 0
'            End With
'        End With
'        With ows.Cells(i, iGTCol)
'            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
'            .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'            With .Interior
'                .Pattern = xlSolid
'                .PatternColorIndex = xlAutomatic
'                .ThemeColor = xlThemeColorAccent1
'                .TintAndShade = 0
'                .PatternTintAndShade = 0
'            End With
'            With .Font
'                .Size = 12
'                .Bold = True
'                .ThemeColor = xlThemeColorDark1
'                .TintAndShade = 0
'            End With
'        End With
'        i = i + 1
'
''Other Project Costs
'        Set lObj = Sheet0.ListObjects("tblTotals")
'        x = i - 1
'        i = i + 1
'        For r = 1 To lObj.DataBodyRange.Rows.count
'            If lObj.DataBodyRange(r, 7) = "Lower" Then
'                If ows.Cells(13, 3).value = "Job Cost" And lObj.DataBodyRange(r, 10).value <> lObj.DataBodyRange(r - 1, 10).value Then
'                    With ows.Cells(i, 2)
'                        .value = lObj.DataBodyRange(r, 10).value & "-" & lObj.DataBodyRange(r, 11).value
'                        .Font.Size = 12
'                        .Font.Bold = True
'                        .Font.Color = -16777216
'                    End With
'                    With ows.Range(ows.Cells(i, 2), ows.Cells(i, iGTCol))
'                         .Interior.Pattern = xlSolid
'                         .Interior.PatternColorIndex = xlAutomatic
'                         .Interior.ThemeColor = xlThemeColorDark2
'                         .Interior.TintAndShade = 0.799981688894314
'                         .Interior.PatternTintAndShade = 0
'                    End With
'                    i = i + 1
'                End If
'                With ows.Cells(i, 3)
'                    .value = lObj.DataBodyRange(r, 4).value
'                    .Font.Size = 12
'                    .InsertIndent 2
'                    .Font.Color = -16777216
'                End With
'                With ows.Cells(i, iGTCol - 12)
'                    .FormulaR1C1 = 0
'                    .NumberFormat = "_(#,##0_);_((#,##0);_(""-""_);_(@_)"
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
'                End With
'                With ows.Cells(i, iGTCol - 11)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],10)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                End With
'                With ows.Cells(i, iGTCol - 10)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],20)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
'                End With
'                With ows.Cells(i, iGTCol - 9)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],30)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                End With
'                With ows.Cells(i, iGTCol - 8)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],40)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
'                End With
'                With ows.Cells(i, iGTCol - 7)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],50)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                End With
'                With ows.Cells(i, iGTCol - 6)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],51)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
'                End With
'                With ows.Cells(i, iGTCol - 5)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],52)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                End With
'                With ows.Cells(i, iGTCol - 4)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],60)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
'                End With
'                With ows.Cells(i, iGTCol - 3)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],61)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                End With
'                With ows.Cells(i, iGTCol - 2)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],62)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
'                End With
'                With ows.Cells(i, iGTCol - 1)
'                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],70)"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                End With
'                With ows.Cells(i, iGTCol)
'                    .FormulaR1C1 = "=SUM(RC[-11]:RC[-1])"
'                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                    .Font.Size = 12
'                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
'                End With
'                i = i + 1
'            End If
'        Next r
'
''Other Project Costs Totals
'        Y = x - i
'        ows.Cells(i, 2).value = "TOTAL"
'        With ows.Range(ows.Cells(i, 2), ows.Cells(i, z))
'            With .Interior
'                .Pattern = xlSolid
'                .PatternColorIndex = xlAutomatic
'                .ThemeColor = xlThemeColorAccent1
'                .TintAndShade = 0
'            End With
'            With .Font
'                .Size = 12
'                .Bold = True
'                .ThemeColor = xlThemeColorDark1
'                .TintAndShade = 0
'            End With
'        End With
'        For C = 1 To 11
'            With ows.Cells(i, iGTCol - C)
'                .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
'                .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'                With .Interior
'                    .Pattern = xlSolid
'                    .PatternColorIndex = xlAutomatic
'                    .ThemeColor = xlThemeColorAccent1
'                    .TintAndShade = 0
'                    .PatternTintAndShade = 0
'                End With
'                With .Font
'                    .Size = 12
'                    .Bold = True
'                    .ThemeColor = xlThemeColorDark1
'                    .TintAndShade = 0
'                End With
'            End With
'        Next C
'        With ows.Cells(i, iGTCol - 12)
'            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
'            .NumberFormat = "_(#,##0_);_((#,##0);_(""-""_);_(@_)"
'            With .Interior
'                .Pattern = xlSolid
'                .PatternColorIndex = xlAutomatic
'                .ThemeColor = xlThemeColorAccent1
'                .TintAndShade = 0
'                .PatternTintAndShade = 0
'            End With
'            With .Font
'                .Size = 12
'                .Bold = True
'                .ThemeColor = xlThemeColorDark1
'                .TintAndShade = 0
'            End With
'        End With
'        With ows.Cells(i, iGTCol)
'            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
'            .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
'            With .Interior
'                .Pattern = xlSolid
'                .PatternColorIndex = xlAutomatic
'                .ThemeColor = xlThemeColorAccent1
'                .TintAndShade = 0
'                .PatternTintAndShade = 0
'            End With
'            With .Font
'                .Size = 12
'                .Bold = True
'                .ThemeColor = xlThemeColorDark1
'                .TintAndShade = 0
'            End With
'        End With
'        i = i + 1
'        ActiveSheet.PageSetup.PrintArea = ows.Range(ows.Cells(1, 2), ows.Cells(i, iGTCol)).Address
'        ActiveSheet.PageSetup.PrintTitleRows = "$1:$10"
'    Application.ScreenUpdating = True
'End Sub
Sub CEstAddons() 'Control Estimate Addons for new codes
Dim dTotal As Double
    sJobUM = Sheet1.Range("rngJobUnitName").Value
    dTotal = Sheet3.Range("SysEnd").Offset(0, 6).Value
    i = 0
    If ows.name = "" Then
        Set ows = ActiveSheet
    End If
'Projected Construction Costs
    Application.ScreenUpdating = False
        If ows.PivotTables.Count > 0 Then
            If pt = "" Then
                Set pt = ActiveSheet.PivotTables(1)
            End If
            With pt.TableRange1
                iGTRow = .Cells(.Cells.Count).row
                iGTCol = .Cells(.Cells.Count).Column
            End With
        Else
            Exit Sub
        End If
        For Each pf In pt.ColumnFields
            If pf.Position = 1 Then
                'sLvl1Item = pf.SourceName
                z = pf.dataRange.Column
                Exit For
            End If
        Next pf
        iAddRow = ActualUsedRange(ows).Rows.Count + 12
        If iAddRow <> iGTRow Then
            ows.Range(ows.Cells(iGTRow, 2).Offset(1, 0), ows.Cells(iAddRow, iGTCol)).EntireRow.Delete
        End If
        Set lObj = Sheet0.ListObjects("tblTotals")
        If lObj.ListRows.Count < 1 Then Exit Sub
        i = iGTRow + 2
        'Add Sort if by Phase Code
        With lObj
            .Sort.SortFields.Clear
            .Sort.SortFields.Add key:=Range("tblTotals[JobCost]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Sort.SortFields.Add key:=Range("tblTotals[SortOrder]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            With .Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End With
        For r = 1 To lObj.DataBodyRange.Rows.Count
            If lObj.DataBodyRange(r, 7) = "Upper" Then
                If ows.Cells(13, 3).Value = "Job Cost" And lObj.DataBodyRange(r, 10).Value <> lObj.DataBodyRange(r - 1, 10).Value Then
                    With ows.Cells(i, 2)
                        .Value = lObj.DataBodyRange(r, 10).Value & "-" & lObj.DataBodyRange(r, 11).Value
                        .Font.Size = 12
                        .Font.Bold = True
                        .Font.Color = -16777216
                    End With
                    With ows.Range(ows.Cells(i, 2), ows.Cells(i, iGTCol))
                         .Interior.Pattern = xlSolid
                         .Interior.PatternColorIndex = xlAutomatic
                         .Interior.ThemeColor = xlThemeColorDark2
                         .Interior.TintAndShade = 0.799981688894314
                         .Interior.PatternTintAndShade = 0
                    End With
                    i = i + 1
                End If
                With ows.Cells(i, 3)
                    .Value = lObj.DataBodyRange(r, 4).Value
                    .Font.Size = 12
                    .InsertIndent 2
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 18)
                    .FormulaR1C1 = 0
                    .NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                With ows.Cells(i, iGTCol - 17)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],10)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 16)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],20)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                With ows.Cells(i, iGTCol - 15)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],21)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 14)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],22)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                With ows.Cells(i, iGTCol - 13)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],30)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 12)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],35)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                With ows.Cells(i, iGTCol - 11)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],40)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 10)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],45)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                With ows.Cells(i, iGTCol - 9)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],49)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 8)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],50)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                With ows.Cells(i, iGTCol - 7)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],51)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 6)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],52)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                 With ows.Cells(i, iGTCol - 5)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],53)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 4)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],60)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                With ows.Cells(i, iGTCol - 3)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],61)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 2)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],62)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                With ows.Cells(i, iGTCol - 1)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],70)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol)
                    .FormulaR1C1 = "=SUM(RC[-11]:RC[-1])"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                i = i + 1
            End If
        Next r
'Projected Construction Costs Totals
        Y = iGTRow - i
        ows.Cells(i, 2).Value = "PROJECTED CONSTRUCTION COSTS"
        With ows.Range(ows.Cells(i, 2), ows.Cells(i, z))
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0
            End With
            With .Font
                .Size = 12
                .Bold = True
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End With
        For C = 1 To 17
            With ows.Cells(i, iGTCol - C)
                .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
                .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                With .Font
                    .Size = 12
                    .Bold = True
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
            End With
        Next C
        With ows.Cells(i, iGTCol - 18)
            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
            .NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With .Font
                .Size = 12
                .Bold = True
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End With
        With ows.Cells(i, iGTCol)
            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
            .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With .Font
                .Size = 12
                .Bold = True
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End With
        i = i + 1
    
'Other Project Costs
        Set lObj = Sheet0.ListObjects("tblTotals")
        X = i - 1
        i = i + 1
        For r = 1 To lObj.DataBodyRange.Rows.Count
            If lObj.DataBodyRange(r, 7) = "Lower" Then
                If ows.Cells(13, 3).Value = "Job Cost" And lObj.DataBodyRange(r, 10).Value <> lObj.DataBodyRange(r - 1, 10).Value Then
                    With ows.Cells(i, 2)
                        .Value = lObj.DataBodyRange(r, 10).Value & "-" & lObj.DataBodyRange(r, 11).Value
                        .Font.Size = 12
                        .Font.Bold = True
                        .Font.Color = -16777216
                    End With
                    With ows.Range(ows.Cells(i, 2), ows.Cells(i, iGTCol))
                         .Interior.Pattern = xlSolid
                         .Interior.PatternColorIndex = xlAutomatic
                         .Interior.ThemeColor = xlThemeColorDark2
                         .Interior.TintAndShade = 0.799981688894314
                         .Interior.PatternTintAndShade = 0
                    End With
                    i = i + 1
                End If
                With ows.Cells(i, 3)
                    .Value = lObj.DataBodyRange(r, 4).Value
                    .Font.Size = 12
                    .InsertIndent 2
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 18)
                    .FormulaR1C1 = 0
                    .NumberFormat = "_(#,##0_);_((#,##0);_(""-""_);_(@_)"
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                With ows.Cells(i, iGTCol - 17)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],10)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 16)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],20)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                With ows.Cells(i, iGTCol - 15)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],21)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 14)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],22)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                With ows.Cells(i, iGTCol - 13)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],30)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 12)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],35)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 11)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],40)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                With ows.Cells(i, iGTCol - 10)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],45)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                With ows.Cells(i, iGTCol - 9)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],49)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 8)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],50)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                With ows.Cells(i, iGTCol - 7)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],51)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 6)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],52)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                With ows.Cells(i, iGTCol - 5)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],53)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 4)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],60)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                With ows.Cells(i, iGTCol - 3)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],61)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol - 2)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],62)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                With ows.Cells(i, iGTCol - 1)
                    .FormulaR1C1 = "=SUMIFS(tblTotals[Amount],tblTotals[Name],RC3,tblTotals[JobCostCategory],70)"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, iGTCol)
                    .FormulaR1C1 = "=SUM(RC[-11]:RC[-1])"
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .Font.Size = 12
                    .Font.Color = -16777216
'                    .Interior.Pattern = xlSolid
'                    .Interior.PatternColorIndex = xlAutomatic
'                    .Interior.ThemeColor = xlThemeColorAccent1
'                    .Interior.TintAndShade = 0.799981688894314
'                    .Interior.PatternTintAndShade = 0
                End With
                i = i + 1
            End If
        Next r
        
'Other Project Costs Totals
        Y = X - i
        ows.Cells(i, 2).Value = "TOTAL"
        With ows.Range(ows.Cells(i, 2), ows.Cells(i, z))
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0
            End With
            With .Font
                .Size = 12
                .Bold = True
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End With
        For C = 1 To 17
            With ows.Cells(i, iGTCol - C)
                .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
                .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                With .Font
                    .Size = 12
                    .Bold = True
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
            End With
        Next C
        With ows.Cells(i, iGTCol - 18)
            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
            .NumberFormat = "_(#,##0_);_((#,##0);_(""-""_);_(@_)"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With .Font
                .Size = 12
                .Bold = True
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End With
        With ows.Cells(i, iGTCol)
            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
            .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With .Font
                .Size = 12
                .Bold = True
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End With
        i = i + 1
        ActiveSheet.PageSetup.PrintArea = ows.Range(ows.Cells(1, 2), ows.Cells(i, iGTCol)).Address
        ActiveSheet.PageSetup.PrintTitleRows = "$1:$10"
    Application.ScreenUpdating = True
End Sub

Sub clearAllAddons()
    For Each ows In ActiveWorkbook.Worksheets
        If ows.PivotTables.Count > 0 Then
            Set pt = ows.PivotTables(1)
            With pt.TableRange1
                iGTRow = .Cells(.Cells.Count).row
                iGTCol = .Cells(.Cells.Count).Column
            End With
            iAddRow = ActualUsedRange(ows).Rows.Count
            If iAddRow > iGTRow Then
                ows.Range(ows.Cells(iGTRow, 2).Offset(1, 0), ows.Cells(iAddRow, iGTCol)).EntireRow.Delete
            End If
        ElseIf ows.CodeName = "Sheet3" Then
            iGTRow = ows.Range("SysEnd").row
            iGTCol = ows.Range("H1").Column
            iAddRow = ActualUsedRange(ows).Rows.Count
            If iAddRow > iGTRow Then
                ows.Range(ows.Cells(iGTRow, 2).Offset(1, 0), ows.Cells(iAddRow, iGTCol)).EntireRow.Delete
            End If
        End If
    Next
End Sub

Sub ReApplyAddons()
On Error GoTo e1

    For Each ows In ActiveWorkbook.Worksheets
        If ows.PivotTables.Count > 0 Then
            Set pt = ows.PivotTables(1)
            If pt.name Like "XTab*" Then
                Call XTabAddons
            ElseIf pt.name Like "Control Estimate*" Then
                Call CEstAddons
            ElseIf pt.name Like "Variance*" Then
                Call Addons_VAR
            Else
                Call Addons
            End If
        ElseIf ows.CodeName = "Sheet3" Then
            Call Addons
        End If
    Next
    Call SheetFormatingAll

Exit Sub
e1:
    logError "failed to apply markups"

End Sub

Sub Addons_VAR()
    If ows.name = "" Then
        Set ows = ActiveSheet
    End If
    Set pt = ows.PivotTables(1)
    i = 0
    sJobUM = Range("rngJobUnitName").Value
    If ows.PivotTables.Count > 0 Then
        If pt = "" Then
            Set pt = ows.PivotTables(1)
        End If
        With pt.TableRange1
            iGTRow = .Cells(.Cells.Count).row
            iGTCol = .Cells(.Cells.Count).Column
        End With
    Else
        iGTRow = ows.Range("SysEnd").row
        iGTCol = ows.Range("H1").Column
    End If
    Application.ScreenUpdating = False
    iAddRow = ActualUsedRange(ows).Rows.Count
    If iAddRow > iGTRow Then
        Exit Sub
    End If
    
    Set lObj = Sheet0.ListObjects("tblTotals")
        If lObj.ListRows.Count < 1 Then Exit Sub
        i = iGTRow + 2
        X = iGTCol
        For r = 1 To lObj.DataBodyRange.Rows.Count
            If lObj.DataBodyRange(r, 7) = "Upper" Then
                With ows.Cells(i, 3)
                    .Value = lObj.DataBodyRange(r, 4).Value
                    .InsertIndent 2
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, X)
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .FormulaR1C1 = lObj.DataBodyRange(r, 14)
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, X - 1)
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .FormulaR1C1 = lObj.DataBodyRange(r, 13)
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, X - 2)
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .FormulaR1C1 = lObj.DataBodyRange(r, 6)
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                i = i + 1
            End If
        Next r
        
        With ows.Cells(i, 2)
            .Value = StrConv("Projected Construction Costs", vbUpperCase)
            .Font.Size = 12
            .Font.Bold = True
            .Font.ThemeColor = xlThemeColorDark1
            .Font.TintAndShade = 0
        End With

        With ows.Cells(i, X)
            Y = iGTRow - i
            .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
            .Font.Size = 12
            .Font.Bold = True
            .Font.ThemeColor = xlThemeColorDark1
            .Font.TintAndShade = 0
        End With
        With ows.Cells(i, X - 2)
            .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
            .Font.Size = 12
            .Font.Bold = True
            .Font.ThemeColor = xlThemeColorDark1
            .Font.TintAndShade = 0
        End With
        With ows.Cells(i, X - 1)
            .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
            .Font.Size = 12
            .Font.Bold = True
            .Font.ThemeColor = xlThemeColorDark1
            .Font.TintAndShade = 0
        End With
        With ows.Range(ows.Cells(i, 2), ows.Cells(i, iGTCol)).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
'Set Allocate False
    r = 0
    Set lObj = Sheet0.ListObjects("tblTotals")
        Y = i
        i = i + 2
        X = iGTCol
        For r = 1 To lObj.DataBodyRange.Rows.Count
            If lObj.DataBodyRange(r, 7) = "Lower" Then
                With ows.Cells(i, 3)
                    .Value = lObj.DataBodyRange(r, 4).Value
                    .InsertIndent 2
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, X)
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .FormulaR1C1 = lObj.DataBodyRange(r, 14)
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, X - 1)
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .FormulaR1C1 = lObj.DataBodyRange(r, 13)
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                With ows.Cells(i, X - 2)
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .FormulaR1C1 = lObj.DataBodyRange(r, 6)
                    .Font.Size = 12
                    .Font.Color = -16777216
                End With
                i = i + 1
            End If
        Next r
        
        With ows.Cells(i, 2)
            .Value = StrConv("TOTAL", vbUpperCase)
            .Font.Size = 12
            .Font.Bold = True
            .Font.ThemeColor = xlThemeColorDark1
            .Font.TintAndShade = 0
        End With

        With ows.Cells(i, X)
            Y = Y - i
            .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
            .Font.Size = 12
            .Font.Bold = True
            .Font.ThemeColor = xlThemeColorDark1
            .Font.TintAndShade = 0
        End With
        With ows.Cells(i, X - 2)
            .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
            .Font.Size = 12
            .Font.Bold = True
            .Font.ThemeColor = xlThemeColorDark1
            .Font.TintAndShade = 0
        End With
        With ows.Cells(i, X - 1)
            .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
            .Font.Size = 12
            .Font.Bold = True
            .Font.ThemeColor = xlThemeColorDark1
            .Font.TintAndShade = 0
        End With
        With ows.Range(ows.Cells(i, 2), ows.Cells(i, iGTCol)).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

    Application.ScreenUpdating = True
    Call SheetFormatting
End Sub

