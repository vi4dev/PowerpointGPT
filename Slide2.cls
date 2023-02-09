VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Slide2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub GetOpenAIResponse()

Dim type_area_value As String
Dim response_gpt As String
Dim obj As Object


type_area_value = type_area.Value

'Send type_area_value to OpenAI API and store the response in response_gpt
response_gpt = GetOpenAIAPIResponse(type_area_value)

Set obj = JsonConverter.ParseJSON(response_gpt)

'Debug.Print response_gpt
Set dic = Module1.ParseJSON(response_gpt)
'Debug.Print dic("obj.choices(0).text")
End Sub

Sub Start_User_Copy()
    Dim type_area_value As String
    Dim response_gpt As String
    Dim obj As Object
    type_area_value = type_area.Value
    response_gpt = GetOpenAIAPIResponse(type_area_value)
    Set obj = JsonConverter.ParseJSON(response_gpt)
    Set dic = Module1.ParseJSON(response_gpt)
    ActivePresentation.Slides(2).Shapes("msg_box").TextFrame.TextRange = "Thinking.."
    ActivePresentation.Slides(2).Shapes("user_box").TextFrame.TextRange = type_area.Value
    ActivePresentation.Slides(2).Shapes("response_box").TextFrame.TextRange = Replace(dic("obj.choices(0).text"), "\n", vbCrLf)
    ActivePresentation.Slides(2).Shapes("msg_box").TextFrame.TextRange = " "
    Debug.Print Replace(dic("obj.choices(0).text"), "\n", vbCrLf)
End Sub





Function GetOpenAIAPIResponse(text As String) As String

Dim request As Object
Set request = CreateObject("MSXML2.XMLHTTP")
'Replace API_KEY and API_ENDPOINT with your actual API key and endpoint
request.Open "POST", "https://api.openai.com/v1/completions", False
request.setRequestHeader "Content-Type", "application/json"
request.setRequestHeader "Authorization", "Bearer " & ActivePresentation.Slides(1).Shapes("api_box").OLEFormat.Object.text
request.send "{ ""model"": ""text-davinci-003"", ""prompt"": """ & text & """, ""max_tokens"": 500 }"


GetOpenAIAPIResponse = request.ResponseText

End Function
Sub ClearConvo()
    ActivePresentation.Slides(2).Shapes("user_box").TextFrame.TextRange = " "
    ActivePresentation.Slides(2).Shapes("response_box").TextFrame.TextRange = " "
    ActivePresentation.Slides(2).Shapes("msg_box").TextFrame.TextRange = "Conversation Cleared.."


End Sub

Sub ExitPresentation()
   Application.Quit
End Sub

Sub ExportSlide2AsPNG()
    Dim Slide As Slide
    Dim SlideCount As Integer
    Dim FileName As String
    Dim Path As String
    
    Set Slide = ActivePresentation.Slides(2)
    SlideCount = 1
    Path = ActivePresentation.Path & "\saves\"
    FileName = "save" & SlideCount & ".png"
    
    If (Dir(Path, vbDirectory) = "") Then
        MkDir Path
    End If
    
    While Dir(Path & FileName) <> ""
        SlideCount = SlideCount + 1
        FileName = "save" & SlideCount & ".png"
    Wend
    
    Slide.Export Path & FileName, "PNG"
End Sub


Private Sub type_area_Change()

End Sub
