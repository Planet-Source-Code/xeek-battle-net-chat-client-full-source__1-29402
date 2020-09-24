VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private mvarFileName As String

Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData
End Property

Public Property Get FileName() As String
    FileName = mvarFileName
End Property

Public Function IsPlay() As Boolean
    Static s As String * 30
    mciSendString "status MP3Play mode", s, Len(s), 0
    MsgBox s
    IsPlay = (Mid(s, 1, 7) = "playing")
End Function

Public Function mmPlay() As String
    Dim cmd As String * 255
    Dim dwret As Long
    Dim ret As String * 128
    Dim tmp As String * 255
    Dim dwLen As Long
    Dim spn As String
    If Dir(mvarFileName) = "" Then
        mmPlay = "error: file doesn't exist"
        Exit Function
    End If
    dwLen = GetShortPathName(mvarFileName, tmp, 255)
    spn = Left(tmp, dwLen)
    cmd = "open " & spn & " type MPEGVideo Alias MP3Play"
    dwret = mciSendString(cmd, 0, 0, 0)
    If dwret <> 0 Then
        mciGetErrorString dwret, ret, 128
        mmPlay = "error: " & ret
        Exit Function
    End If
    mmPlay = "success"
    mciSendString "play MP3Play", 0, 0, 0
End Function

Public Sub mmPause()
    mciSendString "pause MP3Play", 0, 0, 0
End Sub

Public Sub mmStop()
    mciSendString "stop MP3Play", 0, 0, 0
    mciSendString "close MP3Play", 0, 0, 0
End Sub

Public Function PosInSec() As Integer
    Static s As String * 30
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play position", s, Len(s), 0
    PosInSec = Round(s / 1000)
End Function

Public Function LenInSec() As Integer
    Static s As String * 30
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play length", s, Len(s), 0
    LenInSec = Round(s / 1000)
End Function

Public Function LenInMs() As Integer
    Static s As String * 30
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play length", s, Len(s), 0
End Function

Public Function Length() As String
    Static s As String * 30
    Dim sec As Integer
    Dim min As Integer
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play length", s, Len(s), 0
    sec = Round(s / 1000)
    If sec < 60 Then
        Length = "0:" & Format(sec, "00")
    Else
        min = Int(sec / 60)
        sec = sec - (min * 60)
        Length = min & ":" & Format(sec, "00")
    End If
End Function

Public Sub SeekTo(Second As Integer)
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    If IsPlay = True Then
        mciSendString "play MP3Play from " & Second, 0, 0, 0
    Else
        mciSendString "seek MP3Play to " & Second, 0, 0, 0
    End If
End Sub

Private Sub Class_Terminate()
    mmStop
End Sub
