Dim voice, fileStream, fso, timestamp, filePath, userText
' Text to speech generator
' By TonyBmMath on github
' https://www.youtube.com/@LearnMathWithTonyBM
' https://www.facebook.com/LearnMathWithTonyBM
' https://classroom.google.com/c/NzgyNjIzNzA1MTYy?cjc=pou4ig4b
' github.com/tonybmmath
' Website: tonybmmath.github.io
' Prompt user for input text
'1
'2
'3
'4
'5
'34
'32
'31
'11
'10
'9
'8
'7
'6
'16
'15
'14
'13
'12
'11
'10
'9
'8
'7
'6
'31
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
