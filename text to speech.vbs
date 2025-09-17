Dim voice, fileStream, fso, timestamp, filePath, userText
'56
'55
'54
'53
'52
'51
'50
'49
'48
'47
'46
'45
'43
'42
'41
'40
'39
'38
'36
'35
'34
'33
'32
'31
'30
'29
'28
'27
'26
'25
'24
'23
'22
'21
'20
'19
'18
'17
'15
'14
'12
'11
'10
'9
'8
'6
'5
'4
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
