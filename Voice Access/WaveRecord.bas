Attribute VB_Name = "WaveRecord"
Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrrtning As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'This is the essential code for recording a wave file. I didn't need many other interesting features.

Public Sub RecordWave()
Dim Rec As String, Settings As String

Rec = Space$(260)
ResetWave

'Set wave
'Modify these settings if you need different formats. I need only this one.
Settings = "set capture alignment 4 bitspersample 16 samplespersec 22500 channels 2 bytespersec 88200"
mciSendString "seek capture to start", Rec, Len(Rec), 0
mciSendString Settings, Rec, Len(Rec), 0

'Record
mciSendString "record capture", Rec, Len(Rec), 0

End Sub

Public Sub ResetWave()
Dim Rec As String
mciSendString "close all", Rec, Len(Rec), 0
mciSendString "open new type waveaudio alias capture", Rec, Len(Rec), 0
End Sub

Public Sub StopWave()
Dim Rec As String
mciSendString "stop capture", Rec, Len(Rec), 0
End Sub

Public Sub SaveWave(WName As String)
Dim Rec As String, WaveShortName As String
        
     If Spaces(WName) Then
        WaveShortName = GetShortName(WName)
        mciSendString "save capture " & WaveShortName, Rec, Len(Rec), 0
     Else
        mciSendString "save capture " & WName, Rec, Len(Rec), 0
     End If
     
End Sub

Function Spaces(WName As String) As Boolean
Dim a As Long
    Spaces = False
    a = InStr(WName, " ")
    If a <> 0 Then Spaces = True
End Function

'Thanks E. de Vries for this function
Function GetShortName(ByVal sLongFileName As String) As String
Dim lRetVal As Long, sShortPathName As String, iLen As Integer, FilNum As Integer
    'Set up buffer area for API function call return
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)
    FilNum = FreeFile
    'Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    If lRetVal = 0 Then 'The file does not exist, first create it!
        Open sLongFileName For Output As #FilNum
        Close #FilNum
        lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
        'Now another try!
        Kill (sLongFileName)
        'Delete file now!
    End If
    'Strip away unwanted characters.
    GetShortName = Left(sShortPathName, lRetVal)
End Function
