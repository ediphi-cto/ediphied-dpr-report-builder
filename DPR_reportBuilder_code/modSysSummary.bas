Attribute VB_Name = "modSysSummary"
Option Explicit

Sub SummaryDetail()
Dim rSys As Range
Dim sFnd As String
Dim C, firstAddress
Dim k

    Set owb = ActiveWorkbook
    sJobUM = Range("rngJobUnitName").Value
    Set lObj = Sheet0.ListObjects("tblRptTrack")
    'sXpath1 = lObj.DataBodyRange(1, 10).Value
    'bCkb1 = lObj.DataBodyRange(1, 11).Value
    sLvl1xNd = lObj.DataBodyRange(1, 12).Value
    sLvl1Code = lObj.DataBodyRange(1, 13).Value

    Set ows = Sheet3
    ows.Activate
'    ows.Range(ows.Range("SysStart").Offset(1, 0), ows.Range("SysEnd").Offset(-2, 0)).EntireRow.ClearContents
    If ows.Range(ows.Range("SysStart").Offset(1, 0), ows.Range("SysEnd").Offset(-2, 0)).Rows.count > 2 Then
        ows.Range(ows.Range("SysStart").Offset(1, 0), ows.Range("SysEnd").Offset(-2, 0)).EntireRow.Delete
    End If
    
'    On Error Resume Next
'    C = 0
'    For Each k In dict.Keys
'        C = C + 1
'        ows.Range("sysEnd").Offset(-1, 0).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
'        X = ows.Range("sysEnd").Offset(-2, 1).row
'        If InStr(1, sLvl1Code, "Name") = 0 Then
'            ows.Cells(X, 2).Value = dict.Item(k)(0)
'        End If
'        ows.Cells(X, 3).Value = dict.Item(k)(1)
'        ows.Cells(X, 7).Formula = "=IFERROR(RC[1]/rngJobSize,0)"
'        ows.Cells(X, 8).Value = dict.Item(k)(2)
'    Next k
'    On Error GoTo 0
'Sort Summary
    ows.Sort.SortFields.Clear
    ows.Sort.SortFields.Add key:=Range(ows.Range("SysStart").Offset(1, 0), ows.Range("SysEnd").Offset(-2, 0)), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ows.Sort.SortFields.Add key:=Range(ows.Range("SysStart").Offset(1, 1), ows.Range("SysEnd").Offset(-2, 1)), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ows.Sort
        .SetRange Range(ows.Range("SysStart").Offset(1, 0), ows.Range("SysEnd").Offset(-2, 6))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Set Total formula
    r = Range(Range("SysStart"), Range("SysEnd")).count
    ows.Range("SysEnd").Offset(0, 6).FormulaR1C1 = "=SUM(R[-" & r - 1 & "]C:R[-1]C)"
    ows.Range("SysEnd").Offset(0, 6).NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    Call Addons
    Exit Sub
errHndlr:
    Exit Sub
End Sub



