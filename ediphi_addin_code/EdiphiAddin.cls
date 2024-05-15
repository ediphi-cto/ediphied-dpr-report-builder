VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EdiphiAddin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents APP As Application
Attribute APP.VB_VarHelpID = -1
Private Const EDIPHI_TAG As String = "b863abbf-c25a-4a29-97a5-4acdbc9c65ab"

Private Sub APP_WorkbookOpen(ByVal wb As Workbook)
    
    Dim debugModeOn As Boolean
    If wb.Worksheets(1).[a1].Value = EDIPHI_TAG Then
        debugModeOn = (getEnv("DEBUG") = 1)
        If debugModeOn Then Stop
        triggerReportBuilder ediphiCSVfilename:=wb.Name
    End If
    
End Sub

Private Sub Class_Initialize()
    
    'this is how the addin listens to every workbook open event
    Set Me.APP = Application
    
End Sub
