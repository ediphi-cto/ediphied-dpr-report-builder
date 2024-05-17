Attribute VB_Name = "modApplyFx"
Option Explicit

Dim i As Integer
Dim g, v, w, x, Y, z
Dim ows As Worksheet
Dim col As Integer
Dim rGT, c1, c2, c3, c4, c5
Dim rngGT, rng1, rng2, rng3, rng4, rng5, rngMU, rngT As Range
Dim iGT, iL1, iL2, iL3, iL4 As Integer
Dim sGT, sST1, sST2, sST3, sST4 As String
Dim nm, name

Sub ConvertFxL1() 'Convert Level 1 report to range with formulas
    Set ows = ActiveSheet
    On Error Resume Next
    For Each nm In ActiveWorkbook.Names
        name = nm.name
        ActiveWorkbook.Names(name).Delete
    Next nm
    On Error GoTo 0
    col = 10 'Level 1 Total Column position
    sGT = "=SUM("
    sST1 = "=SUM("
    'Find Total row for Level Report
    Set rngGT = ows.Range("B14:B1048576").Find(" SUBTOTAL", LookAt:=xlPart)
    If Not rngGT Is Nothing Then
        rGT = rngGT.row
    End If
    For i = 14 To ows.Range("B1048576").End(xlUp).row
        If i < rGT Then
            If ows.Cells(i, 3).Value <> "" And InStr(ows.Cells(i, 3).Value, "Subtotal") = False Then
                With ows.Range("C" & i)
                    Set rng1 = Range(ows.Range("C" & i), ows.Range("C" & i).End(xlDown))
                    z = rng1.Rows.count - 2
                    Cells(rng1.End(xlDown).row, col).FormulaR1C1 = "=SUM(R[-" & z & "]C:R[-1]C)" ' Level 1 Item Subtotal
                    iGT = iGT + 1
                    g = rngGT.row - rng1.End(xlDown).row
                    If iGT = 1 Then
                        sGT = sGT & "R[-" & g & "]C"
                    Else
                        sGT = sGT & ",R[-" & g & "]C"
                    End If
                End With
            ElseIf ows.Cells(i, 4).Value <> "" Then
                ows.Cells(i, 10).FormulaR1C1 = "=IFERROR(RC[-1]*RC[-3],0)"  'Adds formula at Item level
            End If
        Else
            'Create Report Levels Grand Total
            Cells(rGT, col).FormulaR1C1 = sGT & " )" 'Grand Total
            'Create subtotal for construction costs
            Set rngMU = ows.Range("B" & rGT & ":B1048576").Find(" CONSTRUCTION COSTS", LookAt:=xlPart)
            If Not rngMU Is Nothing Then
                Cells(rngMU.row, col).FormulaR1C1 = "=SUM(R[-" & rngMU.row - rGT & "]C:R[-1]C)"
            End If
            'Create Report Total
            Set rngT = ows.Range("B" & rGT & ":B1048576").Find("TOTAL", LookAt:=xlWhole)
            If Not rngT Is Nothing Then
                Cells(rngT.row, col).FormulaR1C1 = "=SUM(R[-" & rngT.row - rngMU.row & "]C:R[-1]C)"
            End If
            Exit Sub
        End If
    Next
End Sub

Sub ConvertFxL2() 'Convert Level 2 report to range with formulas
    Set ows = ActiveSheet
    On Error Resume Next
    For Each nm In ActiveWorkbook.Names
        name = nm.name
        ActiveWorkbook.Names(name).Delete
    Next nm
    On Error GoTo 0
    col = 12 'Level 2 Total Column position
    sGT = "=SUM("
    sST1 = "=SUM("
    'Find Total row for Level Report
    Set rngGT = ows.Range("B14:B1048576").Find(" SUBTOTAL", LookAt:=xlPart)
    If Not rngGT Is Nothing Then
        rGT = rngGT.row
    End If
    For i = 14 To ows.Range("B1048576").End(xlUp).row
        If i < rGT Then
            If ows.Cells(i, 3).Value <> "" And InStr(ows.Cells(i, 3).Value, "Subtotal") = False Then
                With ows.Range("C" & i)
                    Set rng1 = Range(ows.Range("C" & i), ows.Range("C" & i).End(xlDown))
                    For Each c1 In rng1
                        If c1.Offset(0, 2).Value <> "" And InStr(c1.Offset(0, 2).Value, "Subtotal") = False Then
                            Set rng2 = Range(c1.Offset(0, 2), (c1.Offset(0, 2).End(xlDown)))
                            z = rng2.Rows.count - 2
                            Cells(rng2.End(xlDown).row, col).FormulaR1C1 = "=SUM(R[-" & z & "]C:R[-1]C)" ' Level 2 Item Subtotal
                            Y = rng1.End(xlDown).row - rng2.End(xlDown).row
                            iL1 = iL1 + 1
                            If iL1 = 1 Then
                                sST1 = sST1 & "R[-" & Y & "]C"
                            Else
                                sST1 = sST1 & ",R[-" & Y & "]C"
                            End If
                        ElseIf c1.Offset(0, 3).Value <> "" Then
                            c1.Offset(0, 9).FormulaR1C1 = "=IFERROR(RC[-1]*RC[-3],0)" 'Adds formula at Item level
                        End If
                    Next c1
                    iL1 = 0
                    Cells(rng1.End(xlDown).row, col).FormulaR1C1 = sST1 & " )" 'Level 1 Subtotal
                    sST1 = "=SUM("
                    iGT = iGT + 1
                    g = rngGT.row - rng1.End(xlDown).row
                    If iGT = 1 Then
                        sGT = sGT & "R[-" & g & "]C"
                    Else
                        sGT = sGT & ",R[-" & g & "]C"
                    End If
                End With
            End If
        Else
            'Create Report Levels Grand Total
            Cells(rGT, col).FormulaR1C1 = sGT & " )" 'Grand Total
            'Create subtotal for construction costs
            Set rngMU = ows.Range("B" & rGT & ":B1048576").Find(" CONSTRUCTION COSTS", LookAt:=xlPart)
            If Not rngMU Is Nothing Then
                Cells(rngMU.row, col).FormulaR1C1 = "=SUM(R[-" & rngMU.row - rGT & "]C:R[-1]C)"
            End If
            'Create Report Total
            Set rngT = ows.Range("B" & rGT & ":B1048576").Find("TOTAL", LookAt:=xlWhole)
            If Not rngT Is Nothing Then
                Cells(rngT.row, col).FormulaR1C1 = "=SUM(R[-" & rngT.row - rngMU.row & "]C:R[-1]C)"
            End If
            Exit Sub
        End If
    Next
End Sub

Sub ConvertFxL3() 'Convert Level 3 report to range with formulas
    Set ows = ActiveSheet
    On Error Resume Next
    For Each nm In ActiveWorkbook.Names
        name = nm.name
        ActiveWorkbook.Names(name).Delete
    Next nm
    On Error GoTo 0
    col = 14 'Level 3 Total Column position
    sGT = "=SUM("
    sST1 = "=SUM("
    sST2 = "=SUM("
    'Find Total row for Level Report
    Set rngGT = ows.Range("B14:B1048576").Find(" SUBTOTAL", LookAt:=xlPart)
    If Not rngGT Is Nothing Then
        rGT = rngGT.row
    End If
    For i = 14 To ows.Range("B1048576").End(xlUp).row
        If i < rGT Then
            If ows.Cells(i, 3).Value <> "" And InStr(ows.Cells(i, 3).Value, "Subtotal") = False Then
                With ows.Range("C" & i)
                Set rng1 = Range(ows.Range("C" & i), ows.Range("C" & i).End(xlDown))
                For Each c1 In rng1
                    If c1.Offset(0, 2).Value <> "" And InStr(c1.Offset(0, 2).Value, "Subtotal") = False Then
                        Set rng2 = Range(c1.Offset(0, 2), (c1.Offset(0, 2).End(xlDown)))
                        For Each c2 In rng2
                            If c2.Offset(0, 2).Value <> "" And InStr(c2.Offset(0, 2).Value, "Subtotal") = False Then
                                Set rng3 = Range(c2.Offset(0, 2), (c2.Offset(0, 2).End(xlDown)))
                                z = rng3.Rows.count - 2
                                Cells(rng3.End(xlDown).row, col).FormulaR1C1 = "=SUM(R[-" & z & "]C:R[-1]C)" ' Level 3 Item Subtotal
                                Y = rng2.End(xlDown).row - rng3.End(xlDown).row
                                iL2 = iL2 + 1
                                If iL2 = 1 Then
                                    sST2 = sST2 & "R[-" & Y & "]C"
                                Else
                                    sST2 = sST2 & ",R[-" & Y & "]C"
                                End If
                            ElseIf c2.Offset(0, 3).Value <> "" Then
                                c2.Offset(0, 9).FormulaR1C1 = "=IFERROR(RC[-1]*RC[-3],0)" 'Adds formula at Item level
                            End If
                        Next c2
                        iL2 = 0
                        Cells(rng2.End(xlDown).row, col).FormulaR1C1 = sST2 & " )" 'Level 2 Subtotal
                        sST2 = "=SUM("
                        iL1 = iL1 + 1
                        x = rng1.End(xlDown).row - rng2.End(xlDown).row
                        If iL1 = 1 Then
                            sST1 = sST1 & "R[-" & x & "]C"
                        Else
                            sST1 = sST1 & ",R[-" & x & "]C"
                        End If
                    End If
                Next c1
                iL1 = 0
                Cells(rng1.End(xlDown).row, col).FormulaR1C1 = sST1 & " )" 'Level 1 Subtotal
                sST1 = "=SUM("
                iGT = iGT + 1
                g = rngGT.row - rng1.End(xlDown).row
                If iGT = 1 Then
                    sGT = sGT & "R[-" & g & "]C"
                Else
                    sGT = sGT & ",R[-" & g & "]C"
                End If
            End With
        End If
    Else
        'Create Report Levels Grand Total
        Cells(rGT, col).FormulaR1C1 = sGT & " )" 'Grand Total
        'Create subtotal for construction costs
        Set rngMU = ows.Range("B" & rGT & ":B1048576").Find(" CONSTRUCTION COSTS", LookAt:=xlPart)
        If Not rngMU Is Nothing Then
            Cells(rngMU.row, col).FormulaR1C1 = "=SUM(R[-" & rngMU.row - rGT & "]C:R[-1]C)"
        End If
        'Create Report Total
        Set rngT = ows.Range("B" & rGT & ":B1048576").Find("TOTAL", LookAt:=xlWhole)
        If Not rngT Is Nothing Then
            Cells(rngT.row, col).FormulaR1C1 = "=SUM(R[-" & rngT.row - rngMU.row & "]C:R[-1]C)"
        End If
        Exit Sub
    End If
    Next
End Sub

Sub ConvertFxL4() 'Convert Level 4 report to range with formulas
    Set ows = ActiveSheet
    On Error Resume Next
    For Each nm In ActiveWorkbook.Names
        name = nm.name
        ActiveWorkbook.Names(name).Delete
    Next nm
    On Error GoTo 0
    col = 16 'Level 4 Total Column position
    sGT = "=SUM("
    sST1 = "=SUM("
    sST2 = "=SUM("
    sST3 = "=SUM("
    'Find Total row for Level Report
    Set rngGT = ows.Range("B14:B1048576").Find(" SUBTOTAL", LookAt:=xlPart)
    If Not rngGT Is Nothing Then
        rGT = rngGT.row
    End If
    For i = 14 To ows.Range("B1048576").End(xlUp).row
        If i < rGT Then
            If ows.Cells(i, 3).Value <> "" And InStr(ows.Cells(i, 3).Value, "Subtotal") = False Then
                With ows.Range("C" & i)
                Set rng1 = Range(ows.Range("C" & i), ows.Range("C" & i).End(xlDown))
                For Each c1 In rng1
                    If c1.Offset(0, 2).Value <> "" And InStr(c1.Offset(0, 2).Value, "Subtotal") = False Then
                        Set rng2 = Range(c1.Offset(0, 2), (c1.Offset(0, 2).End(xlDown)))
                        For Each c2 In rng2
                            If c2.Offset(0, 2).Value <> "" And InStr(c2.Offset(0, 2).Value, "Subtotal") = False Then
                                Set rng3 = Range(c2.Offset(0, 2), (c2.Offset(0, 2).End(xlDown)))
                                For Each c3 In rng3
                                    If c3.Offset(0, 2).Value <> "" And InStr(c3.Offset(0, 2).Value, "Subtotal") = False Then
                                        Set rng4 = Range(c3.Offset(0, 2), (c3.Offset(0, 2).End(xlDown)))
                                        iL3 = iL3 + 1
                                        z = rng4.Rows.count - 2
                                        Cells(rng4.End(xlDown).row, col).FormulaR1C1 = "=SUM(R[-" & z & "]C:R[-1]C)" ' Level 4 Item Subtotal
                                        Y = rng3.End(xlDown).row - rng4.End(xlDown).row
                                        If iL3 = 1 Then
                                            sST3 = sST3 & "R[-" & Y & "]C"
                                        Else
                                            sST3 = sST3 & ",R[-" & Y & "]C"
                                        End If
                                    ElseIf c3.Offset(0, 3).Value <> "" Then
                                        c3.Offset(0, 9).FormulaR1C1 = "=IFERROR(RC[-1]*RC[-3],0)" 'Adds formula at Item level
                                    End If
                                Next c3
                                iL3 = 0
                                Cells(rng3.End(xlDown).row, col).FormulaR1C1 = sST3 & " )" ' Level 3 Subtotal
                                sST3 = "=SUM("
                                iL2 = iL2 + 1
                                x = rng2.End(xlDown).row - rng3.End(xlDown).row
                                If iL2 = 1 Then
                                    sST2 = sST2 & "R[-" & x & "]C"
                                Else
                                    sST2 = sST2 & ",R[-" & x & "]C"
                                End If
                            End If
                        Next c2
                        iL2 = 0
                        Cells(rng2.End(xlDown).row, col).FormulaR1C1 = sST2 & " )" 'Level 2 Subtotal
                        sST2 = "=SUM("
                        iL1 = iL1 + 1
                        w = rng1.End(xlDown).row - rng2.End(xlDown).row
                        If iL1 = 1 Then
                            sST1 = sST1 & "R[-" & w & "]C"
                        Else
                            sST1 = sST1 & ",R[-" & w & "]C"
                        End If
                    End If
                Next c1
                iL1 = 0
                Cells(rng1.End(xlDown).row, col).FormulaR1C1 = sST1 & " )" 'Level 1 Subtotal
                sST1 = "=SUM("
                iGT = iGT + 1
                g = rngGT.row - rng1.End(xlDown).row
                If iGT = 1 Then
                    sGT = sGT & "R[-" & g & "]C"
                Else
                    sGT = sGT & ",R[-" & g & "]C"
                End If
            End With
        End If
    Else
        'Create Report Levels Grand Total
        Cells(rGT, col).FormulaR1C1 = sGT & " )" 'Grand Total
        'Create subtotal for construction costs
        Set rngMU = ows.Range("B" & rGT & ":B1048576").Find(" CONSTRUCTION COSTS", LookAt:=xlPart)
        If Not rngMU Is Nothing Then
            Cells(rngMU.row, col).FormulaR1C1 = "=SUM(R[-" & rngMU.row - rGT & "]C:R[-1]C)"
        End If
        'Create Report Total
        Set rngT = ows.Range("B" & rGT & ":B1048576").Find("TOTAL", LookAt:=xlWhole)
        If Not rngT Is Nothing Then
            Cells(rngT.row, col).FormulaR1C1 = "=SUM(R[-" & rngT.row - rngMU.row & "]C:R[-1]C)"
        End If
        Exit Sub
    End If
    Next
End Sub

Sub ConvertFxL5() 'Convert Level 5 report to range with formulas
    Set ows = ActiveSheet
    On Error Resume Next
    For Each nm In ActiveWorkbook.Names
        name = nm.name
        ActiveWorkbook.Names(name).Delete
    Next nm
    On Error GoTo 0
    col = 18 'Level 5 Total Column position
    sGT = "=SUM("
    sST1 = "=SUM("
    sST2 = "=SUM("
    sST3 = "=SUM("
    sST4 = "=SUM("
    'Find Total row for Level Report
    Set rngGT = ows.Range("B14:B1048576").Find(" SUBTOTAL", LookAt:=xlPart)
    If Not rngGT Is Nothing Then
        rGT = rngGT.row
    End If
    For i = 14 To ows.Range("B1048576").End(xlUp).row
        If i < rGT Then
            If ows.Cells(i, 3).Value <> "" And InStr(ows.Cells(i, 3).Value, "Subtotal:") = False Then
                With ows.Range("C" & i)
                Set rng1 = Range(ows.Range("C" & i), ows.Range("C" & i).End(xlDown))
                For Each c1 In rng1
                    If c1.Offset(0, 2).Value <> "" And InStr(c1.Offset(0, 2).Value, "Subtotal") = False Then
                        Set rng2 = Range(c1.Offset(0, 2), (c1.Offset(0, 2).End(xlDown)))
                        For Each c2 In rng2
                            If c2.Offset(0, 2).Value <> "" And InStr(c2.Offset(0, 2).Value, "Subtotal") = False Then
                                Set rng3 = Range(c2.Offset(0, 2), (c2.Offset(0, 2).End(xlDown)))
                                For Each c3 In rng3
                                    If c3.Offset(0, 2).Value <> "" And InStr(c3.Offset(0, 2).Value, "Subtotal") = False Then
                                        Set rng4 = Range(c3.Offset(0, 2), (c3.Offset(0, 2).End(xlDown)))
                                        For Each c4 In rng4
                                        If c4.Offset(0, 2).Value <> "" And InStr(c4.Offset(0, 2).Value, "Subtotal") = False Then
                                            Set rng5 = Range(c4.Offset(0, 2), (c4.Offset(0, 2).End(xlDown)))
                                            iL4 = iL4 + 1
                                            z = rng5.Rows.count - 2
                                            Cells(rng5.End(xlDown).row, col).FormulaR1C1 = "=SUM(R[-" & z & "]C:R[-1]C)" 'Level 5 (Item Level Subtotal)
                                            Y = rng4.End(xlDown).row - rng5.End(xlDown).row
                                            If iL4 = 1 Then
                                                sST4 = sST4 & "R[-" & Y & "]C"
                                            Else
                                                sST4 = sST4 & ",R[-" & Y & "]C"
                                            End If
                                        ElseIf c4.Offset(0, 3).Value <> "" Then
                                            c4.Offset(0, 9).FormulaR1C1 = "=IFERROR(RC[-1]*RC[-3],0)" 'Adds formula at Item level
                                        End If
                                        Next c4
                                        iL4 = 0
                                        Cells(rng4.End(xlDown).row, col).FormulaR1C1 = sST4 & " )" 'Level 4 Subtotal
                                        sST4 = "=SUM("
                                        iL3 = iL3 + 1
                                        x = rng3.End(xlDown).row - rng4.End(xlDown).row
                                        If x = 0 Then x = 1
                                        If iL3 = 1 Then
                                            sST3 = sST3 & "R[-" & x & "]C"
                                        Else
                                            sST3 = sST3 & ",R[-" & x & "]C"
                                        End If
                                    End If
                                Next c3
                                iL3 = 0
                                Cells(rng3.End(xlDown).row, col).FormulaR1C1 = sST3 & " )" 'Level 3 Subtotal
                                sST3 = "=SUM("
                                iL2 = iL2 + 1
                                w = rng2.End(xlDown).row - rng3.End(xlDown).row
                                If w = 0 Then w = 1
                                If iL2 = 1 Then
                                    sST2 = sST2 & "R[-" & w & "]C"
                                Else
                                    sST2 = sST2 & ",R[-" & w & "]C"
                                End If
                            End If
                        Next c2
                        iL2 = 0
                        Cells(rng2.End(xlDown).row, col).FormulaR1C1 = sST2 & " )" 'Level 2 Subtotal
                        sST2 = "=SUM("
                        iL1 = iL1 + 1
                        v = rng1.End(xlDown).row - rng2.End(xlDown).row
                        If v = 0 Then v = 1
                        If iL1 = 1 Then
                            sST1 = sST1 & "R[-" & v & "]C"
                        Else
                            sST1 = sST1 & ",R[-" & v & "]C"
                        End If
                    End If
                Next c1
                iL1 = 0
                Cells(rng1.End(xlDown).row, col).FormulaR1C1 = sST1 & " )" 'Level 1 Subtotal
                sST1 = "=SUM("
                iGT = iGT + 1
                g = rngGT.row - rng1.End(xlDown).row
                
                If iGT = 1 Then
                    sGT = sGT & "R[-" & g & "]C"
                Else
                    sGT = sGT & ",R[-" & g & "]C"
                End If
            End With
        End If
    Else
        'Create Report Levels Grand Total
        Cells(rGT, col).FormulaR1C1 = sGT & " )" 'Grand Total
        'Create subtotal for construction costs
        Set rngMU = ows.Range("B" & rGT & ":B1048576").Find(" CONSTRUCTION COSTS", LookAt:=xlPart)
        If Not rngMU Is Nothing Then
            Cells(rngMU.row, col).FormulaR1C1 = "=SUM(R[-" & rngMU.row - rGT & "]C:R[-1]C)"
        End If
        'Create Report Total
        Set rngT = ows.Range("B" & rGT & ":B1048576").Find("TOTAL", LookAt:=xlWhole)
        If Not rngT Is Nothing Then
            Cells(rngT.row, col).FormulaR1C1 = "=SUM(R[-" & rngT.row - rngMU.row & "]C:R[-1]C)"
        End If
        Exit Sub
    End If
    Next
End Sub
