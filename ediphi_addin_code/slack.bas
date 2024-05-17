Attribute VB_Name = "slack"
Option Explicit

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

Function slackPost(msg As String, Optional url As String) As Boolean
    On Error GoTo e1
    Dim Data As String
    Dim RESTcall As Object
    Data = slackBody(msg, url)
    
    Set RESTcall = CreateObject("MSXML2.XMLHTTP")
    With RESTcall
        .Open "POST", SLACK_WEBHOOK, False
        .setRequestHeader "Content-type", "application/json"
        .send (Data)
        slackPost = (.responseText = "ok")
    End With

Exit Function
e1:
    slackPost = False
End Function

Function quote(str As String) As String

    quote = Chr(34) & str & Chr(34)
    
End Function

