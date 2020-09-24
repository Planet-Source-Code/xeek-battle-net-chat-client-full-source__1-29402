Attribute VB_Name = "mod"
Private Declare Sub Sleep Lib "kernel32" ( _
     ByVal dwMilliseconds As Long)
Private Declare Function SetTimer Lib "user32" ( _
     ByVal hWnd As Long, _
     ByVal nIDEvent As Long, _
     ByVal uElapse As Long, _
     ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" ( _
     ByVal hWnd As Long, _
     ByVal nIDEvent As Long) As Long

Public Const mp3Finish = 15

'these are to remember which section in the registry to save/read from
Public Const dApp As String = "Battle.net Bot"
Public Const dSec As String = "ops"

'globals
Public un As String 'username
Public chan As String 'channel
Public buddy As Integer
Public mp3 As clsMP3
Public list As mp3List

Public Sub NextMp3(ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    Select Case nIDEvent
        Case mp3Finish
            mp3.mmStop
            mp3.FileName = list.ListNext
            mp3.mmPlay
            SetTimer frmMain.hWnd, mp3Finish, mp3.LenInMs, AddressOf NextMp3
    End Select
End Sub

Public Sub Owner(dCom As String)
    'direct commands
    Select Case LCase(dCom)
        'start help
            Case "help"
                tell dApp & " v" & App.Major & "." & App.Minor & "." & App.Revision
                tell "for each mp3 command, you may type " & Chr(34) & "help [command]" & Chr(34) & _
                    " (don't type the brackets when identifying a command)... here are the commands:"
                tell "[dir] [play] [stop] [pause] [ff] [rw] [next] [prev] [random] [shuffle] [rdir]"
                tell "[dir *] [play *] [load *] [remove] [add *] [remove *] [clear] [list]"
                tell "note: mp3 play is continuous"
            Case "help dir"
                tell "the dir command tells you which directory is set, and how many mp3 are in it"
            Case "help play"
                tell "the play command starts playing mp3 at the current index"
            Case "help stop"
                tell "stops mp3 from playing"
            Case "help pause"
                tell "pauses the mp3"
            Case "help ff"
                tell "fast-forwards the mp3 5 seconds"
            Case "help rw"
                tell "rewinds the mp3 5 seconds"
            Case "help next"
                tell "starts playing the next mp3"
            Case "help prev"
                tell "starts playing the previous mp3"
            Case "help random"
                tell "plays a random mp3 in the list"
            Case "help shuffle"
                tell "randomizes the list of mp3"
            Case "help rdir"
                tell "resets the list of mp3 to the mp3s that are in the current directory"
            Case "help dir *"
                tell "clears the list, sets the directory, and adds the mp3 in the directory to the list"
            Case "help play *"
                tell "will play a certain song, you replace the * with part of the song name to play"
            Case "help load *"
                tell "this will load *.m3u files only (mp3 play lists), clear the list of mp3, and add the playlist to it"
            Case "help remove"
                tell "this will remove the current song from the list of mp3"
            Case "help add *"
                tell "this will add mp3s with * in their names"
            Case "help remove *"
                tell "this will remove mp3s with * in their names"
            Case "help clear"
                tell "this will clear the list of mp3 in the list"
            Case "help list"
                tell "this will tell you how many mp3 are currently in the list"
        'end help
        'start commands
            Case "dir"
                If cDir = "" Then
                    tell "no directory is set, to set a directory, type " & Chr(34) & "dir [directory]" & Chr(34) & ", replacing [directory] with the folder path to use"
                    Exit Sub
                End If
                tell "current directory: " & frmMain.tMp3Dir.Text & " | mp3 count: " & countMp3
            Case "play"
                
                mp3.mmPlay
            Case "stop"
            Case "pause"
            Case "ff"
            Case "rw"
            Case "next"
            Case "prev"
            Case "random"
            Case "shuffle"
            Case "rdir"
            Case "remove"
            Case "clear"
            Case "list"
        'end commands
    End Select
    
    'dynamic commands
    If dCom Like "dir *" Then
    End If
    If dCom Like "play *" Then
    End If
    If dCom Like "load *" Then
    End If
    If dCom Like "add *" Then
    End If
End Sub

Public Sub tell(what As String)
    On Error GoTo offline
    'send our whisper to the owner whatever we want to say
    frmMain.ws.SendData "/w " & frmMain.tOwner.Text & " " & what
    'keep from flooding the server and getting disconnected
    Sleep 750
    Exit Sub
offline:
    If Err.Number = 40006 Then Disc
End Sub

Public Function countMp3(Optional addThem As Boolean) As Integer
    Dim s As String, c As Integer
    s = Dir(frmMain.tMp3Dir.Text & "\*.mp3")
    If Trim(s) = "" Then Exit Function
    If addThem = True Then list.AddMp3 frmMain.tMp3Dir.Text & "\" & s
    c = c + 1
    Do: DoEvents
        s = Dir
        If Trim(s) = "" Then Exit Do
        If addThem = True Then list.AddMp3 frmMain.tMp3Dir.Text & "\" & s
        c = c + 1
    Loop
    countMp3 = c
    mp3.FileName = list.getCurrent
    mp3.mmPlay
End Function

Public Sub Display(Title As String, daMsg As String)
    'load the display form
    Load frmDisp
    'set the title
    frmDisp.Caption = Title
    'set what text to display to the user
    frmDisp.txt.Text = daMsg
    'show the display with vbModal so the main form (owner) is disabled until the display is closed
    frmDisp.Show vbModal, frmMain
End Sub

Public Sub Disc()
    With frmMain
        'if we get disconnected, set the tabs to the general tab
        .daTab.Tab = 0
        'disable the other tabs
        .daTab.TabEnabled(1) = False
        .daTab.TabEnabled(2) = False
        'reset the login button's caption
        .bLogin.Caption = "connect"
        Cap
    End With
End Sub

Public Sub Conn()
    With frmMain
        'once we connect, enable the tabs
        .daTab.TabEnabled(1) = True
        .daTab.TabEnabled(2) = True
        'show the chat tab
        .daTab = 2
        'no caption on the login button
        'changing the value property 'clicks' it,
        'and it goes on what the caption
        'is when it's clicked
        .bLogin.Caption = ""
        'set the value so the button is raised
        .bLogin.Value = 0
        .bLogin.Caption = "disconnect"
        Cap "connected"
    End With
End Sub

Public Sub Cap(Optional dCap As String)
    'if there's no caption use a default caption
    If dCap = "" Then _
        frmMain.Caption = dApp & " v" & App.Major & "." & App.Minor & "." & App.Revision: Exit Sub
    'otherwise use the default & the caption wanted
    frmMain.Caption = dApp & " v" & App.Major & "." & App.Minor & "." & App.Revision & " - " & dCap
End Sub
