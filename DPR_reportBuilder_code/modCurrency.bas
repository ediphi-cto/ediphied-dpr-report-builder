Attribute VB_Name = "modCurrency"
Sub CurrencyFormat()
Dim sOrgCur_0, sOrgCur_2 As String
Dim sNewCur_0, sNewCur_2 As String
Dim ws As Worksheet, C As Range

sOrgCur_0 = Application.Range("rngOrgCur_0").NumberFormatLocal
sOrgCur_2 = Application.Range("rngOrgCur_2").NumberFormatLocal
sNewCur_0 = Application.Range("rngNewCur_0").NumberFormatLocal
sNewCur_2 = Application.Range("rngNewCur_2").NumberFormatLocal
'Convert 0 Decimal
Application.ScreenUpdating = False
    For Each ws In ActiveWorkbook.Sheets
        If ws.CodeName <> "Sheet1" Then
            With Application
                .FindFormat.Clear
                .ReplaceFormat.Clear
                .FindFormat.NumberFormat = sOrgCur_0
                .ReplaceFormat.NumberFormat = sNewCur_0
            End With
            ws.Cells.Replace What:="", Replacement:="", Lookat:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=True, ReplaceFormat:=True
        End If
    Next ws
    With Application
        .FindFormat.Clear
        .ReplaceFormat.Clear
    End With
    'Convert 2 decimal
    For Each ws In ActiveWorkbook.Sheets
        If ws.CodeName <> "Sheet1" Then
            With Application
                .FindFormat.Clear
                .ReplaceFormat.Clear
                .FindFormat.NumberFormat = sOrgCur_2
                .ReplaceFormat.NumberFormat = sNewCur_2
            End With
            ws.Cells.Replace What:="", Replacement:="", Lookat:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=True, ReplaceFormat:=True
        End If
    Next ws
    With Application
        .FindFormat.Clear
        .ReplaceFormat.Clear
        .ScreenUpdating = True
    End With
    Range("rngOrgCur_0").NumberFormat = Range("rngNewCur_0").NumberFormatLocal
    Range("rngOrgCur_2").NumberFormat = Range("rngNewCur_2").NumberFormatLocal
End Sub
