Attribute VB_Name = "modRefresh"
'TODO- kill this file

'Public MainBar As ProgressBar
'Public SubBar As ProgressBar
'Public SubCounter As Long
'Public SubCount As Long
'Public sMsg As String
'Public bRefresh As Boolean
'
'Sub refreshReports()
''Dim oMod As Object
''Dim C
''Dim TotalCount As Long
''    bRefresh = True
'''Initialize a New Instance of the Progressbars
''    Set MainBar = New ProgressBar
''    Set SubBar = New ProgressBar
''    TotalCount = 10
''    SubCount = 6
''    With MainBar
''        .Title = "DPR Reporter Data Refresh"
''        .ExcelStatusBar = True
''        .StartColour = rgbGreen
''        .EndColour = rgbGreen 'rgbRed
''        .TotalActions = TotalCount
''    End With
''
''    With SubBar
''        .Title = "Refreshing Reports"
''        .ExcelStatusBar = True
''        .StartColour = rgbBlue
''        .EndColour = rgbBlue
''    End With
''    MainBar.ShowBar
''
''    For Counter = 1 To TotalCount
''        Select Case Counter
''        Case 1
''            MainBar.NextAction "Gathering data from WinEst......", True
'''            Call Get_EstData 'Run Refresh API
''        Case 2
''            MainBar.NextAction "Updating Project Information......", True
''            Call clearAllAddons
''        Case 3
''            MainBar.NextAction "Updating Markups......", True
''            Call xmlProjectInfo
''        Case 4
''            MainBar.NextAction "Updating Alternates Tables......", True
''            Call xmlTotalsTableRefresh
''        Case 5
''            MainBar.NextAction "Updating WBS Tables......", True
''            Call UpdateAlternates
''        Case 6
''            MainBar.NextAction "Updating Systems Summary......", True
''            Call WBSDefinition
''        Case 7
''            MainBar.NextAction "Updating Executive Summary......", True
''            Call SummaryDetail
''        Case 8
''            MainBar.NextAction "Updating Reports......", True
''            Call ExecSummary
''        Case 9
''             For Each ows In ActiveWorkbook.Worksheets
''                If ptExists = True Then
''                    sRprt = ows.PivotTables(1).Name
''                    Set lObj = Sheet0.ListObjects("tblRptTrack")
''                    Set C = lObj.ListColumns(1).DataBodyRange.Find(sRprt, LookIn:=xlValues)
''                    If Not C Is Nothing Then
''                        i = C.row - 1
''                        If InStr(lObj.DataBodyRange(i, 1), "XTab") = 1 Then
''                            sMod = "xmlXTabLevel" & lObj.DataBodyRange(i, 2).Value
''                            Y = 0
''
''                        ElseIf InStr(lObj.DataBodyRange(i, 1), "Control Estimate") = 1 Then
''                            sMod = "xmlCtrlEst" & lObj.DataBodyRange(i, 2).Value
''                            Y = 1
'''                        ElseIf InStr(lObj.DataBodyRange(i, 1), "Variance Report") = 1 Then
'''                            sMod = "xml_VAR_Level" & lObj.DataBodyRange(i, 2).Value
'''                            Y = 1
''                        Else
''                            sMod = "xmlLevel" & lObj.DataBodyRange(i, 2).Value
''                            Y = 1
''                        End If
''                        sRprt = lObj.DataBodyRange(i, 1)
''                        iLvl = lObj.DataBodyRange(i, 2)
''                        bCkbAll = lObj.DataBodyRange(i, 3)
''                        bCkbSub = lObj.DataBodyRange(i, 4)
'''                        SubBar.ShowBar
'''                        SubBar.Top = MainBar.Top + MainBar.Height + 10
'''                        SubBar.Left = MainBar.Left
'''                        SubBar.Title = "Refreshing " & sRprt & "...."
'''                        SubBar.TotalActions = 0
''                        For X = Y To lObj.DataBodyRange(i, 2)
''                            Select Case X
''                                Case 0
''                                    sXpath0 = lObj.DataBodyRange(i, 5).Value
''                                    bCkb0 = lObj.DataBodyRange(i, 6)
''                                    sLvl0xNd = lObj.DataBodyRange(i, 7).Value
''                                    sLvl0Code = lObj.DataBodyRange(i, 8).Value
''                                    sLvl0Item = lObj.DataBodyRange(i, 9).Value
''                                Case 1
''                                    sXpath1 = lObj.DataBodyRange(i, 10).Value
''                                    bCkb1 = lObj.DataBodyRange(i, 11)
''                                    sLvl1xNd = lObj.DataBodyRange(i, 12).Value
''                                    sLvl1Code = lObj.DataBodyRange(i, 13).Value
''                                    sLvl1Item = lObj.DataBodyRange(i, 14).Value
''                                Case 2
''                                    sXpath2 = lObj.DataBodyRange(i, 15).Value
''                                    bCkb2 = lObj.DataBodyRange(i, 16)
''                                    sLvl2xNd = lObj.DataBodyRange(i, 17).Value
''                                    sLvl2Code = lObj.DataBodyRange(i, 18).Value
''                                    sLvl2Item = lObj.DataBodyRange(i, 19).Value
''                                Case 3
''                                    sXpath3 = lObj.DataBodyRange(i, 20).Value
''                                    bCkb3 = lObj.DataBodyRange(i, 21)
''                                    sLvl3xNd = lObj.DataBodyRange(i, 22).Value
''                                    sLvl3Code = lObj.DataBodyRange(i, 23).Value
''                                    sLvl3Item = lObj.DataBodyRange(i, 24).Value
''                                Case 4
''                                    sXpath4 = lObj.DataBodyRange(i, 25).Value
''                                    bCkb4 = lObj.DataBodyRange(i, 26)
''                                    sLvl4xNd = lObj.DataBodyRange(i, 27).Value
''                                    sLvl4Code = lObj.DataBodyRange(i, 28).Value
''                                    sLvl4Item = lObj.DataBodyRange(i, 29).Value
''                                Case 5
''                                    sXpath5 = lObj.DataBodyRange(i, 30).Value
''                                    bCkb5 = lObj.DataBodyRange(i, 31)
''                                    sLvl5xNd = lObj.DataBodyRange(i, 32).Value
''                                    sLvl5Code = lObj.DataBodyRange(i, 33).Value
''                                    sLvl5Item = lObj.DataBodyRange(i, 34).Value
''                            End Select
''                        Next X
''                        Application.Run sMod
''                    End If
''                    With ows.PivotTables(1).PivotCache
'''                        sMsg = "Refreshing Pivot Cache...."
'''                        SubBar.NextAction sMsg, True
''                        Set .Recordset = rsNew
''                        .Refresh
''                    End With
''                    'Sort Report Levels
''                    If iLvl >= 1 Then ows.PivotTables(1).PivotFields(sLvl1Code).AutoSort xlAscending, sLvl1Code
''                    If iLvl >= 2 Then ows.PivotTables(1).PivotFields(sLvl2Code).AutoSort xlAscending, sLvl2Code
''                    If iLvl >= 3 Then ows.PivotTables(1).PivotFields(sLvl3Code).AutoSort xlAscending, sLvl3Code
''                    If iLvl >= 4 Then ows.PivotTables(1).PivotFields(sLvl4Code).AutoSort xlAscending, sLvl4Code
''                    If iLvl = 5 Then ows.PivotTables(1).PivotFields(sLvl5Code).AutoSort xlAscending, sLvl5Code
''                    'Sort Item level
''                    If InStr(1, sRprt, "XTab") = 0 Then ows.PivotTables(1).PivotFields("ItemCode").AutoSort xlAscending, "ItemCode"
'''                SubBar.Terminate
''                End If
''            Next ows
''        Case 10
''            MainBar.NextAction "Re-applying Markups......", True
''            Call ReApplyAddons
''            MainBar.NextAction "Data Refresh Finalizing...", True
''        End Select
''    Next Counter
''TerminateSub:
'    MainBar.Complete 3
'    If bFileSave = True Then Call FileSaveAs
'    bRefresh = False
'    MsgBox "Data refresh is complete.", vbInformation, "Data Refresh"
'End Sub
'Function ptExists() As Boolean
'    On Error Resume Next
'    ptExists = Not (IsError(ows.PivotTables(1)))
'    On Error GoTo 0
'End Function
'
'Function FolderFromPath(ByRef strFullPath As String) As String
'
'     FolderFromPath = Left(strFullPath, InStrRev(strFullPath, "\"))
'
'End Function
'
