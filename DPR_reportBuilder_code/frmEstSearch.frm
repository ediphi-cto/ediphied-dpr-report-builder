VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstSearch 
   Caption         =   "DPR Report Builder"
   ClientHeight    =   11640
   ClientLeft      =   924
   ClientTop       =   3612
   ClientWidth     =   10116
   OleObjectBlob   =   "frmEstSearch.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sCriteria As String
Dim sBusUnit As String
Dim sProjSrch, sEstimatorSrch, sFileSrch As String
Dim i As Long, arr As Variant

Sub SVRConnect()
    sUser = Environ("UserName")
    sCompID = Environ("COMPUTERNAME")
    On Error GoTo errSQLConn
    sConnection = "Provider=SQLOLEDB;Data Source=" & sSvr & ";Initial Catalog=" & sDb & ";User ID=" & sDbUser & ";Password=" & sDbPass & ";"
    Set objCon = CreateObject("ADODB.Connection")
    Set objRst = CreateObject("ADODB.Recordset")
    On Error GoTo 0
    Exit Sub
errSQLConn:
    MsgBox "Could not connect to the server. Check your Internet connections and try again.", vbCritical, "Server Connection Failed"
    Exit Sub
End Sub

Sub loadTree()
Dim nLvl1, nLvl2, nLvl3, nLvl4, nLvl5 As Node
Dim sKey1, sKey2, sKey3, sKey4, sKey5, sKey6 As String
Dim X, Y As Integer

'Connect to Estimate Project Info Table
    If optBU.Value = True Then
        sCriteria = "(Business_Unit = '" & Range("rngRegion").Value & "') AND (Project_Name LIKE '%" & sProjSrch & "%') AND (Lead_Estimator LIKE '%" & sEstimatorSrch & "%') AND (Est_File_Name LIKE '%" & sFileSrch & "%')"
    Else
        sCriteria = "(Project_Name LIKE '%" & sProjSrch & "%') AND (Lead_Estimator LIKE '%" & sEstimatorSrch & "%') AND (Est_File_Name LIKE '%" & sFileSrch & "%')"
    End If
    Call SVRConnect
    objCon.Open sConnection
    sSql = "SELECT TOP (100) PERCENT Business_Unit, Lead_Estimator, Project_Name, Est_File_Name, Est_File_Location, ids_Est_GUID"
    sSql = sSql & " FROM     PRECON.tblEST_Project_Info"
    sSql = sSql & " WHERE " & sCriteria
    sSql = sSql & " ORDER BY Business_Unit, Lead_Estimator, Project_Name"
    objRst.Open sSql, objCon
    tv.Nodes.Clear
    Y = 0
    On Error Resume Next
    If Not objRst.BOF And Not objRst.EOF Then
        Do
        sKey1 = objRst(0)
        tv.Nodes.Add , , sKey1, objRst(0)
        tv.Nodes(sKey1).Expanded = True
        tv.Nodes(sKey1).Bold = True
'LEVEL ONE***************************************************************************
        sKey2 = sKey1 & "-" & objRst(1)
        Set nLvl1 = tv.Nodes.Add(sKey1, tvwChild, sKey2, "ESTIMATOR: " & objRst(1))
        nLvl1.Expanded = True
        nLvl1.ForeColor = &H814901
        tv.Nodes(sKey2).Bold = True
'        nLvl1.Tag = objRst(5)
        If Err.Number <> 0 Then
            Err.Clear
            Set nLvl1 = tv.Nodes(objRst(1))
        End If
'LEVEL TWO***************************************************************************
        If objRst(2) <> "" Then
            sKey3 = sKey2 & "-" & objRst(2)
            Set nLvl2 = tv.Nodes.Add(sKey2, tvwChild, sKey3, objRst(2))
            Err.Clear
        End If
'LEVEL THREE
        If objRst(3) <> "" Then
            sKey4 = sKey3 & "-" & objRst(3)
            Set nLvl3 = tv.Nodes.Add(sKey3, tvwChild, sKey4, objRst(3))
            nLvl3.FontStyle.Italic = True
            nLvl3.Tag = objRst(5)
            Err.Clear
        End If
        Y = Y + 1
        objRst.MoveNext
        Loop Until objRst.EOF
    End If
    lblResults.Caption = "Found " & Y & " record(s)"
    On Error GoTo 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Range("rngVarGUID").Value = sVarGuid
    Range("rngVarEstID").Value = txtSelEstimate.Value
    frmReportLevel.txtVarXML.Value = txtSelEstimate.Value
    Call xmlVarTotalsTable
    Unload Me
End Sub

Private Sub optAllBU_Click()
    loadTree
End Sub

Private Sub optBU_Click()
    loadTree
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
Dim nodX As Node
     sVarGuid = Node.Tag
     Set nodX = Node.Parent
     txtSelEstimate.Value = Node
End Sub

Private Sub txtSrchEstFile_Change()
    sFileSrch = txtSrchEstFile.Value
    loadTree
End Sub

Private Sub txtSrchEstimator_Change()
    sEstimatorSrch = txtSrchEstimator.Value
    loadTree
End Sub

Private Sub txtSrchProjName_Change()
    sProjSrch = txtSrchProjName.Value
    loadTree
End Sub

Private Sub txtSrchProjName_Enter()
'    Label38.visible = False
End Sub

Private Sub UserForm_Activate()
    optBU.Caption = "Show Estimates for " & Range("rngRegion").Value
    txtCurEstimate.Value = Range("rngEstimateID").Value
    loadTree
    With txtSrchProjName
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 115
    Me.Left = Application.Left + 482
End Sub

Private Sub CheckChildren(Node As Node)
'Dim i As Integer, nodX As Node, C As Range
'    If Node.Children <> 0 Then 'If node has children
'        Set nodX = Node.Child 'Catch first child
'        For i = 1 To Node.Children 'Loop through each child
'            nodX.Checked = Node.Checked 'Set as checked
'            Set C = lObj.ListColumns(4).DataBodyRange.Find(nodX, LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True)
'            If Not C Is Nothing Then
'                C.Offset(0, 1).Value = nodX.Checked
'            End If
'            CheckChildren nodX 'Check to see if this node has children
'            Set nodX = nodX.Next 'Catch next child
'        Next 'Loop
'    Else
'        Set C = lObj.ListColumns(4).DataBodyRange.Find(Node, LookIn:=xlValues, Lookat:=xlWhole, MatchCase:=True)
'        If Not C Is Nothing Then
'            C.Offset(0, 1).Value = Node.Checked
'        End If
'    End If
End Sub

Function removeAlpha(r As String) As String
    With CreateObject("vbscript.regexp")
        .Pattern = "[A-Za-z]"
        .Global = True
        removeAlpha = .Replace(r, "")
    End With
End Function
