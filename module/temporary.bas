Attribute VB_Name = "temporary"
'Option Explicit
'
Private Declare Function GetTickCount Lib "kernel32" () As Long
'
''선언부 시작
'Private Const LWA_COLORKEY As Long = &H1
'Private Const GWL_EXSTYLE As Long = -20
'Private Const WS_EX_LAYERED As Long = &H80000
'
'Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
''선언부 끝
'
'Private Sub Form_Load()
'
'    Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, -20) Or WS_EX_LAYERED)
'    Call SetLayeredWindowAttributes(Me.hwnd, 0, 255 * 0.4, 2)
'End Sub
'
'
'
'
''Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
''Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
''Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
''Private Const GWL_EXSTYLE = (-20)
''Private Const WS_EX_LAYERED = &H80000
''Private Const LWA_ALPHA = &H2
''Private Sub Form_Load()
'''    Dim WndStyle As Long
'''     WndStyle = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
'''     WndStyle = WndStyle Or ws_ex_layered
'''     Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, WndStyle)
'''     SetLayeredWindowAttributes Me.hWnd, 0, 255 * (0.8), LWA_ALPHA
''
''    FadeIN Me
''End Sub
''
''Private Sub Form_Unload(Cancel As Integer)
''    FadeOUT Me
''End Sub
'
'
'
''선언부 시작
'Private Const LWA_COLORKEY As Long = &H1
'Private Const GWL_EXSTYLE As Long = -20
'Private Const WS_EX_LAYERED As Long = &H80000
'
'Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
''선언부 끝
'
'Private Sub Form_Load()
'
'    Dim ret As Long
'    ret = GetWindowLong(Me.hwnd, -20) Or WS_EX_LAYERED                  '투명스타일 적용
'    Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, ret)
'    Call SetLayeredWindowAttributes(Me.hwnd, vbMagenta, 255 * 0.4, 2) '자홍색(vbMagenta)을 투명으로 변경
'End Sub
'

'--- 나머지는 클립보드 데이타 헨들링할때 필요한것들
'Private Declare Function SetClipboardViewer Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function ChangeClipboardChain Lib "user32" (ByVal hwnd As Long, ByVal hWndNext As Long) As Long
'Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function CloseClipboard Lib "user32" () As Long
'Private Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
'Private Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
'Private Declare Function EmptyClipboard Lib "user32" () As Long
'Private Declare Function GetClipboardData Lib "user32" Alias "GetClipboardDataA" (ByVal wFormat As Long) As Long
'Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
'Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
'Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
'Private Const GMEM_DDESHARE = &H2000
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Const MB_ As Boolean = True



'### 딜레이
Public Sub Delay(ByVal Value As Long)
    Dim before As Long
    before = GetTickCount()
    Do Until GetTickCount() >= before + Value
        DoEvents
    Loop
End Sub

'### SplitEX
Function splitex(ByVal s$, ByVal S1$, ByVal S2$) As String
    splitex = Split(Split(s, S1)(1), S2)(0)
End Function

'### CopyIt
Public Sub copyit(ByVal str$)
    Clipboard.Clear
    Clipboard.SetText str
End Sub

'### N과 M사이의 난수 발생 ###
Function RndNum(ByVal n As Long, ByVal m As Long) As Long
    Randomize
    RndNum = Int(Rnd * ((m + 1) - n)) + n
End Function




'### 테스트용 메세지박스
Sub MB(ByVal str$)
    If Not MB_ Then Exit Sub
    MsgBox str$
End Sub


'### 5개 integer검사해서 최대값
Function max5(ByVal a%, ByVal b%, ByVal c%, ByVal d%, ByVal e%) As Integer
    Dim i%, arr(5) As Integer, max%
    arr(0) = a
    arr(1) = b
    arr(2) = c
    arr(3) = d
    arr(4) = e
    
    max = 0
    
    For i = 0 To 4
        If max < arr(i) Then max = arr(i)
    Next i
    
    max5 = max
    
    
End Function

