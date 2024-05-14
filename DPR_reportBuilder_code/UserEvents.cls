VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const TEAMS_WEBHOOK As String = ""
Const SLACK_WEBHOOK As String = "https://hooks.slack.com/triggers/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"


Function slackBody(msg As String, Optional url As String)
        Dim user As String, tenant As String
        user = Application.UserName
        tenant = getEnv("TENANT")
        
        slackBody = "{""username"": " & quote(user) & _
              ", ""url"": " & quote(url) & _
              ", ""message"": " & quote(msg) & _
              ", ""tenant"": " & quote(tenant) & "}"

End Function

Function teamsPost(msg As String) As Boolean
    
    Dim data As String
    Dim RESTcall As New MSXML2.XMLHTTP60
    data = toJSON(msg)
    
    With RESTcall
        .Open "POST", TEAMS_WEBHOOK, False
        .setRequestHeader "Content-type", "application/json"
        .send (data)
        teamsPost = (.responseText = "ok")
    End With

Exit Function
e1:
    teamsPost = False
End Function

Function slackPost(msg As String, Optional url As String) As Boolean
    On Error GoTo e1
    Dim data As String
    Dim RESTcall As Object
    msg = thisReportBuilder.metaDataAsString & msg
    data = slackBody(msg, url)
    
    Set RESTcall = CreateObject("MSXML2.XMLHTTP")
    With RESTcall
        .Open "POST", SLACK_WEBHOOK, False
        .setRequestHeader "Content-type", "application/json"
        .send (data)
        slackPost = (.responseText = "ok")
    End With

Exit Function
e1:
    slackPost = False
End Function

Function toJSON(msg As String) As String
    
    toJSON = "{""text"":" & quote(msg) & "}"

End Function





