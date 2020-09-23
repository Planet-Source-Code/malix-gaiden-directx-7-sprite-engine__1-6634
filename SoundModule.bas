Attribute VB_Name = "SoundModule"
Const snd_async = &H1
Const snd_nodefault = &H2
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long


Sub LoadMidi(Filename As String, alias As String)

Dim result As Integer
result = mciSendString("open " & Filename & ".mid type sequencer alias " & alias, 0&, 0, 0)
End Sub
Sub PlayMidi(alias As String)
Dim result As Integer
result = mciSendString("play " & alias, 0&, 0, 0)
End Sub
Sub StopMidi(alias As String)

Dim result As Integer
result = mciSendString("close " & alias, 0&, 0, 0)
End Sub


Sub PlayWave(alias As String)
Dim result As Integer
result = sndPlaySound(alias & ".wav", snd_async Or snd_nodefault)
End Sub
Sub CloseSound()
Dim result As Integer
result = mciSendString("close all", 0&, 0, 0)

End Sub
