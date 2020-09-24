VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mp3List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private cMp3() As String
Private current As Integer

Public Function ListShuffle() As Boolean
    'basically to shuffle an array, just loop through the array, and switch values
    On Error GoTo messed
    Dim i As Integer, f As Integer, b As String
    'loop through the array
    For i = LBound(cMp3) To UBound(cMp3)
        Randomize
        'find a random number, so we can grab the contents of the array
        f = Int((Rnd * UBound(cMp3)) + LBound(cMp3))
        'make sure we know what it used to be
        b = cMp3(i)
        'make it equal the random index
        cMp3(i) = cMp3(f)
        'set the random index in the array to what the other one was
        cMp3(f) = b
    Next
    Shuffle = True
messed:
End Function

Public Function getRandom() As String
    On Error GoTo messed
    Randomize
    'get a random mp3
    getRandom = cMp3(Int((Rnd * UBound(cMp3)) + LBound(cMp3)))
messed:
End Function

Public Function ListNext() As String
    On Error GoTo messed
    'increment our current position
    current = current + 1
    'if our position is higher than the highest, set it to the lowest
    If current > UBound(cMp3) Then current = LBound(cMp3)
    'return the next mp3
    ListNext = cMp3(current)
messed:
End Function

Public Function ListPrev() As String
    On Error GoTo messed
    'decrement our current position
    current = current - 1
    'if our position is lower than the lowest, set it to the highest
    If current < LBound(cMp3) Then current = UBound(cMp3)
    'return the previous mp3
    ListPrev = cMp3(current)
messed:
End Function

Public Function removeCurrent() As Boolean
    On Error GoTo messed
    Dim i As Integer
    'replace each position by the one above it
    For i = current To (UBound(cMp3) - 1)
        cMp3(i) = cMp3(i + 1)
    Next
    'take out the last position
    ReDim Preserve cMp3(UBound(cMp3) - 1)
    removeCurrent = True
messed:
End Function

Public Function AddMp3(what As String) As Boolean
    'if there's an error that means the array isn't initialized
    On Error GoTo init
    'if the file isn't there, don't do nothing
    If Dir(what) = "" Then Exit Function
finish:
    'if the first position is empty, fill it, otherwise make room for another, and fill it
    If cMp3(0) = "" Then
        cMp3(0) = what
    Else
        ReDim Preserve cMp3(UBound(cMp3) + 1)
        cMp3(UBound(cMp3)) = what
    End If
    AddMp3 = True
    Exit Function
init:
    'initialize the array
    ReDim cMp3(0)
    GoTo finish
End Function

Public Function getCurrent() As String
    On Error GoTo messed
    'retrieve the current mp3
    getCurrent = cMp3(current)
    Exit Function
messed:
End Function

Public Function Amount() As Integer
    On Error GoTo messed
    'return the amount of mp3 in the list
    Amount = UBound(cMp3) - 1
messed:
End Function
