Function GetHTML(URL As String) As String
    Dim Html As String
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", URL, False
        .Send
        GetHTML = .ResponseText
    End With
End Function
Function SplitText(text As String, Regex As String, Index As Integer) As String
    Dim Data() As String
    Data = Split(text, Regex)
    SplitText = Data(Index)
End Function
Private Sub ChatVIew_Click()

End Sub
Private Sub LoadChatView_Click()
    Dim Chat As String
    Chat = TextInputBox.Value
    Dim Value As String
    Value = GetHTML("https://builder.pingpong.us/api/builder/pingpong/chat/demo?query=" & ENDECODEURL(Chat))
    Value = SplitText(Value, "reply"":""", 1)
    Value = SplitText(Value, """,""type", 0)
    If CheckContains(Value, "\") Then
        Value = SplitText(Value, "\", 0)
    End If
    Value = "[BOT] - " & Value
    Dim PreChat As String
    PreChat = ChatVIew.Caption
    Dim NewChat As String
    NewChat = PreChat & vbCrLf & "[나] - " & Chat & vbCrLf & Value
    ChatVIew.Caption = NewChat
    TextInputBox.Value = ""
End Sub
Private Sub TextInputBox_Change()

End Sub
Function ENDECODEURL(varText As Variant, Optional blnEncode = True)
Static objHtmlfile As Object
    If objHtmlfile Is Nothing Then
      Set objHtmlfile = CreateObject("htmlfile")
      With objHtmlfile.parentWindow
        .execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
        .execScript "function decode(s) {return decodeURIComponent(s)}", "jscript"
      End With
    End If
    If blnEncode Then
      ENDECODEURL = objHtmlfile.parentWindow.encode(varText)
    Else
      ENDECODEURL = objHtmlfile.parentWindow.decode(varText)
    End If
End Function
Function CheckContains(text As String, find As String)
    If InStr(text, find) > 0 Then
        CheckContains = True
    Else
        CheckContains = False
    End If
End Function

