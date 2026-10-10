Attribute VB_Name = "Platform"
'==============================================================================
' Platform: the only module that talks to the Windows API.
' Keyboard state, a high-resolution clock, sleep and window focus.
'==============================================================================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private mFrequency As Currency

' True while the key with virtual-key code vk is physically held down.
Public Function KeyDown(ByVal vk As Long) As Boolean
    KeyDown = (GetAsyncKeyState(vk) And &H8000) <> 0
End Function

' Seconds from an arbitrary origin, with sub-millisecond resolution.
' Currency is a 64-bit integer scaled by 10000, so the ratio is exact enough.
Public Function Clock() As Double
    Dim ticks As Currency
    If mFrequency = 0 Then QueryPerformanceFrequency mFrequency
    QueryPerformanceCounter ticks
    Clock = ticks / mFrequency
End Function

Public Sub Nap(ByVal milliseconds As Long)
    If milliseconds > 0 Then Sleep milliseconds
End Sub

' True when the Excel window is the one receiving keystrokes, so the game
' does not react to keys typed in another application.
Public Function HasFocus() As Boolean
    HasFocus = (GetForegroundWindow() = Application.hWnd)
End Function

' While the game runs, Excel must ignore the game keys: otherwise typing "w"
' would start editing a cell. OnKey with "" disables a key, and OnKey without
' a procedure restores it. GetAsyncKeyState still sees the physical key.
Public Sub CaptureKeys(ByVal capture As Boolean)
    Dim keys As Variant
    Dim i As Long

    keys = Array("w", "a", "s", "d", "m", "+w", "+a", "+s", "+d", "+m", _
                 "{UP}", "{DOWN}", "{LEFT}", "{RIGHT}")

    On Error Resume Next
    For i = LBound(keys) To UBound(keys)
        If capture Then
            Application.OnKey keys(i), ""
        Else
            Application.OnKey keys(i)
        End If
    Next i
    On Error GoTo 0
End Sub
