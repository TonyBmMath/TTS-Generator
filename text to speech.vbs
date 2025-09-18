Dim voice, fileStream, fso, timestamp, filePath, userText
'161
'160
'159
'158
'157
'155
'154
'153
'151
'150
'149
'148
'147
'146
'145
'143
'142
'141
'140
'139
'138
'137
'136
'135
'134
'132
'131
'130
'129
'128
'127
'126
'125
'124
'123
'122
'121
'120
'119
'118
'117
'116
'115
'114
'112
'110
'109
'106
'104
'102
'101
'100
'99
'98
'97
'96
'95
'94
'93
'92
'91
'90
'89
'87
'86
'84
'83
'82
'81
'80
'79
'78
' Text to speech generator
' By TonyBmMath on github
' https://www.youtube.com/@LearnMathWithTonyBM
' https://www.facebook.com/LearnMathWithTonyBM
' https://classroom.google.com/c/NzgyNjIzNzA1MTYy?cjc=pou4ig4b
' github.com/tonybmmath
' Website: tonybmmath.github.io
' Prompt user for input text
userText = InputBox("Enter the text to convert to speech:", "Text to Speech")

' Exit if user cancels or leaves input empty
If Trim(userText) = "" Then
    WScript.Echo "No input provided. Exiting."
    WScript.Quit
End If

Set voice = CreateObject("SAPI.SpVoice")
Set fileStream = CreateObject("SAPI.SpFileStream")
Set fso = CreateObject("Scripting.FileSystemObject")

' Create a safe timestamp string for filename
timestamp = Replace(Replace(Replace(Now, ":", "-"), "/", "-"), " ", "_")
filePath = fso.GetAbsolutePathName(".") & "\tts_" & timestamp & ".wav"

' Open stream for writing (3 = SSFMCreateForWrite)
fileStream.Open filePath, 3, False

' Assign the stream and speak the input text
Set voice.AudioOutputStream = fileStream
voice.Speak userText

' Close the stream
fileStream.Close

' Notify user of success
MsgBox "Speech exported to:" & vbCrLf & filePath, vbInformation, "Done"
