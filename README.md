# 🗣️ VBScript Text-to-Speech Converter

This script allows users to convert any text input into spoken audio and automatically saves it as a `.wav` file using Windows' built-in Speech API (SAPI).

## 📋 Features

- Prompts user to enter text via an input box  
- Converts the entered text to speech  
- Saves the spoken output as a `.wav` file with a timestamped filename  
- Notifies the user when the file is successfully created  

## 🛠️ Requirements

- Windows OS  
- Windows Script Host (WSH) enabled  
- SAPI (Speech API) — included by default in most Windows installations  

## 🚀 How to Use

1. **Save the script** as a `.vbs` file (e.g., `TextToSpeech.vbs`)
2. **Double-click** the file to run it
3. **Enter your text** in the input box when prompted
4. The script will:
   - Convert the text to speech
   - Save the audio as a `.wav` file in the same directory
   - Show a message box with the file path

## 📁 Output

The audio file will be saved in the same folder as the script, with a filename like:
```
tts_YYYY-MM-DD_HH-MM-SS.wav
```

Example:
```
tts_2025-08-26_12-27-00.wav
```


## ⚠️ Notes

- If no text is entered, the script will exit gracefully.
- The script uses `SAPI.SpVoice` and `SAPI.SpFileStream` COM objects, which are native to Windows.
- This script does **not** require internet access or external libraries.

*It does not have a female voice yet.*
