Attribute VB_Name = "modExecSummary"
Option Explicit

Sub ExecSummary()
On Error GoTo e1

    Dim rngSeries As Excel.Range
    Dim rngCategory As Excel.Range
    Dim dProjCosts As Double
    Dim dMarkup As Double
    Dim RngToCover As Range
    Dim ChtOb As ChartObject

    sJobUM = Range("rngJobUnitName").Value
    Set ows = Sheet2
    'On Error GoTo errHndlr
    ows.visible = xlSheetVisible
    ows.Activate
    iGTRow = ows.Range("exConstCosts").row
    iGTCol = ows.Range("exTotal").Offset(0, 5).Column
    iAddRow = Range("exTotal").row
    If iAddRow > iGTRow + 2 Then
        ows.Range(Cells(iGTRow, 2).Offset(2, 0), Cells(iAddRow - 1, iGTCol)).EntireRow.Delete
    End If
        
'Calc Project Costs ****ABOVE**** Markups
    Set lObj = Sheet0.ListObjects("tblTotals")
    dProjCosts = Sheet3.Range("SysEnd").Offset(0, 6).Value
    
    If lObj.ListRows.count > 0 Then
        dMarkup = Application.WorksheetFunction.SumIf(lObj.ListColumns(7).DataBodyRange, "Upper", lObj.ListColumns(6).DataBodyRange)
        If dMarkup = 0 Then
            dProjCosts = dProjCosts
        Else
            dProjCosts = dProjCosts + dMarkup
        End If
        With Cells(iGTRow, 5)
            .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
            .FormulaR1C1 = "=IFERROR(RC[2]/rngJobSize,0)"
        End With
        With Cells(iGTRow, 7)
            .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
            .Value = dProjCosts
        End With
        i = ows.Range("exTotal").row
    
    'Load Systems Summary ****BELOW**** Markups
        For r = 1 To lObj.DataBodyRange.Rows.count
            If lObj.DataBodyRange(r, 7).Value = "Lower" Then
                Cells(i, 1).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                With Cells(i, 2)
                    .Value = lObj.DataBodyRange(r, 4).Value
                    .Font.Size = 12
                    .Font.color = -16777216
                End With
                If lObj.DataBodyRange(r, 9).Value <> "" Then
                    With Cells(i, 4)
                        .Style = "Percent"
                        .FormulaR1C1 = lObj.DataBodyRange(r, 5) / 100
                        .NumberFormat = "0.00%"
                        .Font.Size = 12
                        .Font.color = -16777216
                    End With
                End If
                With Cells(i, 7)
                    .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
                    .FormulaR1C1 = lObj.DataBodyRange(r, 6)
                    .Font.Size = 12
                    .Font.color = -16777216
                End With
                With Cells(i, 5)
                    .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
                    .FormulaR1C1 = "=IFERROR(RC[2]/rngJobSize,0)"
                    .Font.Size = 12
                    .Font.color = -16777216
                End With
                i = i + 1
            End If
        Next r
    
        With Cells(i, 5)
            .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
            .FormulaR1C1 = "=IFERROR(RC[2]/rngJobSize,0)"
        End With
        With Cells(i, 7)
            Y = iGTRow - i
            .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
            .FormulaR1C1 = "=SUM(R[" & Y & "]C:R[-1]C)"
        End With
        
        Range(Cells(57, 1), Cells(1048576, 1)).EntireRow.Delete
    Else
        With Cells(iGTRow, 5)
            .NumberFormat = Range("rngNewCur_2").NumberFormatLocal
            .FormulaR1C1 = "=IFERROR(RC[2]/rngJobSize,0)"
        End With
        With Cells(iGTRow, 7)
            .NumberFormat = Range("rngNewCur_0").NumberFormatLocal
            .Value = dProjCosts
        End With
    End If
setCharts:
'******Update Chart data series******
    Set rngSeries = Nothing
    Set rngCategory = Nothing
    Set rngSeries = Sheet3.Range(Sheet3.Range("SysStart").Offset(1, 1), Sheet3.Range("SysEnd").Offset(-1, 1))
    Set rngCategory = Sheet3.Range(Sheet3.Range("SysStart").Offset(1, 6), Sheet3.Range("SysEnd").Offset(-1, 6))
    ActiveSheet.ChartObjects("chrtExecSummary").Activate
    ActiveChart.SetSourceData Source:=Sheets("Systems Summary").Range(rngSeries.address & "," & rngCategory.address)

    Set RngToCover = ActiveSheet.Range("$H$27:$K$55")
    Set ChtOb = ActiveChart.Parent
    ChtOb.Height = RngToCover.Height
    ChtOb.Width = RngToCover.Width
    ChtOb.Top = RngToCover.Top
    ChtOb.Left = RngToCover.Left
errHndlr:
    'Call SheetFormatting
    Range("A1").Activate

Exit Sub
e1:
    logError "failed to build executive summary"

End Sub

