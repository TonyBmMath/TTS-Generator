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
'92
'91
'90
'89
'87
'86
'85
'84
'83
'82
'81
'80
'79
'78
'77
'76
'75
'74
'73
'72
'71
'70
'69
'67
'66
'65
'64
'63
'62
'59
'58
'56
'52
'51
'60
'59
'58
'56
'54
'53
'52
'51
'42
'41
'40
'39
'38
'37
'6
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
