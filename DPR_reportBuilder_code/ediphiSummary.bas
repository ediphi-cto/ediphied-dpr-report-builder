Attribute VB_Name = "ediphiSummary"
Option Explicit

Sub createSummary(groupByHeader As String)

    Dim summaryWS As Worksheet
    Set summaryWS = ThisWorkbook.Worksheets("Systems Summary")
    
    Dim summaryRows As Collection
    Set summaryRows = createSummaryDictColl(groupByHeader)
    
    Dim printRan As Range
    Set printRan = summaryWS.[\print]
    printRan.Offset(1, 0).Resize(summaryRows.count - 2, 1).EntireRow.Insert
    Set printRan = printDictList2Range(dictColl:=summaryRows, startCell:=printRan, noHeaders:=True)
    
    With summaryWS
        .Sort.SortFields.Add key:=Range(.Range("SysStart").Offset(1, 0), .Range("SysEnd").Offset(-2, 0)), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .SetRange printRan
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        .visible = xlSheetVisible
    End With

End Sub

Function createSummaryDictColl(groupByHeader As String) As Collection

    Dim pivotDataWS As Worksheet
    Set pivotDataWS = ThisWorkbook.Worksheets("pivot data")
    
    Dim tbl As ListObject
    Set tbl = pivotDataWS.ListObjects("tblEdiphiPivotData")
    
    Dim headerRowRan As Range
    Dim sortCodeCol As Range
    Dim sortNameCol As Range
    Dim totalRan As Range
    
    Set headerRowRan = tbl.HeaderRowRange
    Set sortCodeCol = headerRowRan.Find(What:=groupByHeader & "_code", LookIn:=xlValues, LookAt:=xlWhole)
    Set sortNameCol = headerRowRan.Find(What:=groupByHeader, LookIn:=xlValues, LookAt:=xlWhole)
    Set totalRan = headerRowRan.Find(What:="GrandTotal", LookIn:=xlValues, LookAt:=xlWhole)
    
    If sortCodeCol Is Nothing Or sortNameCol Is Nothing Then GoTo e1
    If totalRan Is Nothing Then GoTo e2
    
    Set sortCodeCol = sortCodeCol.EntireColumn
    Set sortNameCol = sortNameCol.EntireColumn
    Set totalRan = totalRan.Offset(1, 0)
    
    Dim sortCodeRan As Range, sortNameRan As Range
    Dim summaryDict As New Dictionary
    Dim rowDict As Dictionary
    Do Until totalRan.row > tbl.Range.Rows.count
        Set sortCodeRan = Intersect(sortCodeCol.EntireColumn, totalRan.EntireRow)
        Set sortNameRan = Intersect(sortNameCol.EntireColumn, totalRan.EntireRow)
        If Not summaryDict.Exists(sortCodeRan.Value) Then
            Set rowDict = New Dictionary
            With rowDict
                .Add "code", "'" & sortCodeRan.Value
                .Add "description", sortNameRan.Value
                .Add "blank1", ""
                .Add "blank2", ""
                .Add "blank3", ""
                .Add "cost per sf", "=IFERROR(RC[1]/rngJobSize,0)"
                .Add "total", 0
            End With
            summaryDict.Add sortCodeRan.Value, rowDict
        End If
        summaryDict(sortCodeRan.Value)("total") = summaryDict(sortCodeRan.Value)("total") + cDbl_safe(totalRan.Value)
        Set totalRan = totalRan.Offset(1, 0)
    Loop
    
    Set createSummaryDictColl = New Collection
    Dim k
    For Each k In summaryDict.Keys
        createSummaryDictColl.Add summaryDict(k)
    Next
    
Exit Function
e1:
    logError "Could not find WBS column " & groupByHeader & " in tblEdiphiPivotData when creating Summary"

Exit Function
e2:
    logError "Could not find total cost column in tblEdiphiPivotData"
    
End Function

Sub ghj()

    Dim pivotDataWS As Worksheet
    Set pivotDataWS = ThisWorkbook.Worksheets("pivot data")
    
    Dim tbl As ListObject
    Set tbl = pivotDataWS.ListObjects("tblEdiphiPivotData")

    pp tbl.Range

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

End Sub
