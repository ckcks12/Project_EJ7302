Attribute VB_Name = "mod_hook"
Option Explicit

'---���콺/Ű���� ��
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hHook As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'---

'---Ŭ������ü�� ����
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
'---

'---Ŭ�������ڵ鸵
Private Declare Function SetClipboardViewer Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ChangeClipboardChain Lib "user32" (ByVal hwnd As Long, ByVal hWndNext As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_DDESHARE = &H2000
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)
'---

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type


'---Ű���� ����ü
Public Type KeyStruct
    vkcode As Long
    scancode As Long
    flags As Long
    time As Long
    dwExtrainfo As Long
End Type
'---

Private hMouse As Long
Private hKeyboard As Long
Public myhwnd As Long '�ʱ�ȭ��������� �� hwnd�� ��� frmMain Load�ɶ� hwnd�θŹ��ʱ�ȭ
Public CBEXFolder As String
Public cbexDoing As Boolean 'cbex���� Ŭ�����庹�����϶� ����Ÿū��� �ߺ�������������Ԥ���Ǥ���������������������������
Private cbexIndex As Integer
Public keyrecognizing As Boolean
Private OldProc As Long


Private suicide As Boolean

Dim bCtrl As Boolean
Dim bAlt As Boolean
Dim bKey(256) As Boolean 'Ű����Ű���������

Public Function MouseHookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

End Function

Public Function KeyboardHookProc(ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As KeyStruct) As Long: On Error GoTo h
    If cbexDoing Then GoTo h
    'If keyrecognizing Then GoTo h
    
    With lParam
        Select Case wParam
            Case 257 '�� ������ �´ٸ�... 257�� ��� Ű�� KeyUp message���..
                'If .vkcode = 162 Then bCtrl = False
                'If .vkcode = 164 Then bAlt = False
                If .vkcode > 0 And .vkcode < 256 Then bKey(.vkcode) = False
            Case Else '�� ���� ��� wParam���� KeyPress�� ���ƴѰ�����
                'If .vkcode = 162 Then bCtrl = True
                'If .vkcode = 164 Then bAlt = True
                If .vkcode > 0 And .vkcode < 256 Then bKey(.vkcode) = True
                
        End Select
    End With
    
    If GetAsyncKeyState(162) And GetAsyncKeyState(164) Then
        keyrecognizing = True
        If bKey(90) Then 'z
            Dim tmpCursor As POINTAPI
            GetCursorPos tmpCursor
            SetAlphaEX WindowFromPoint(tmpCursor.X, tmpCursor.Y), value_SetAlpha_Alpha
        ElseIf bKey(88) Then 'x
            GetCursorPos tmpCursor
            SetAlphaEX WindowFromPoint(tmpCursor.X, tmpCursor.Y), 1
        ElseIf bKey(160) Then 'shift
            frmOCRmaina.Show
            suicide = True
        ElseIf bKey(80) Then
            frmMain.Form1ShowFromfrmMain
            suicide = True
        ElseIf bKey(65) Then
            frmZoom.Show
        ElseIf bKey(83) Then
            frmSearch.Show
        ElseIf bKey(112) Then
            frmMain.SC.Run "f1"
            'frmcblist.show
        ElseIf bKey(113) Then
            frmMain.SC.Run "f2"
            'frmcblist.show
        ElseIf bKey(114) Then
            frmMain.SC.Run "f3"
            'frmcblist.show
        ElseIf bKey(115) Then
            frmMain.SC.Run "f4"
            'frmcblist.show
        ElseIf bKey(116) Then
            frmMain.SC.Run "f5"
            'frmcblist.show
        ElseIf bKey(117) Then
            frmMain.SC.Run "f6"
            'frmcblist.show
        ElseIf bKey(118) Then
            frmMain.SC.Run "f7"
            'frmcblist.show
        ElseIf bKey(119) Then
            frmMain.SC.Run "f8"
            'frmcblist.show
        ElseIf bKey(120) Then
            frmMain.SC.Run "f9"
            'frmcblist.show
        ElseIf bKey(121) Then
            frmMain.SC.Run "f10"
            'frmcblist.show
        ElseIf bKey(122) Then
            frmMain.SC.Run "f11"
            'frmcblist.show
        ElseIf bKey(123) Then
            frmMain.SC.Run "f12"
            'frmcblist.show
        ElseIf bKey(81) Then 'Q
            frmMM.Show
        End If
        '��Ʈ��+��Ʈ ���������� ó��
    End If
    
    '---Ŭ������ Ư��ó��
    If GetAsyncKeyState(162) Then
        keyrecognizing = True
        If bKey(86) Then 'v
            cbexDoing = True
            bKey(86) = False
            'cbexindex���� ��絥��Ÿ�������ͼ� Set�Ѵ�����
            'cbexindex���������ϰ�
            'cbexindex -1
            
            Dim tmp As String, b() As Byte, bLen As Long, ff As Integer: ff = FreeFile
            Dim nowtmp As String
            Dim hMem(999) As Long, hIndex As Integer
            
            tmp = CBEXFolder & "\" & cbexIndex & "\"
            Debug.Print "load>>" & cbexIndex & "///" & cbexDoing & "///" & bKey(86)
            
            If LenB(Dir$(tmp)) Then '������ ������쿣
                
                Clipboard.Clear
                nowtmp = OpenFile(tmp & "1.cbex")
                If value_CB_Script Then nowtmp = frmMain.SC.Run("cb", nowtmp)
                Clipboard.SetText nowtmp
            
'                OpenClipboard myhwnd
'
'                EmptyClipboard
'
'                Open tmp & Dir$(tmp) For Binary Access Read As #FF
'                    ReDim b(FileLen(tmp & Dir$(tmp))) As Byte
'                    Get #FF, , b
'                    DoEvents
'                Close #FF
'
'                hMem(hIndex) = GlobalAlloc(GMEM_DDESHARE, UBound(b) * 2)
'                GlobalLock hMem(hIndex)
'                CopyMemory ByVal hMem(hIndex), b(0), UBound(b) * 2
'                GlobalUnlock hMem(hIndex)
'
'                SetClipboardData CLng(Replace$(Dir$(tmp), ".cbex", vbNullString)), hMem(hIndex): DoEvents
'
'
'                hIndex = hIndex + 1
'
'                nowtmp = Dir$
'                Do While LenB(nowtmp)
'                    Open tmp & nowtmp For Binary Access Read As #FF
'                        ReDim b(FileLen(tmp & nowtmp)) As Byte
'                        Get #FF, , b
'                        DoEvents
'                    Close #FF
'
'                    hMem(hIndex) = GlobalAlloc(GMEM_DDESHARE, UBound(b) * 2)
'                    GlobalLock hMem(hIndex)
'                    CopyMemory ByVal hMem(hIndex), b(0), UBound(b) * 2
'                    GlobalUnlock hMem(hIndex)
'
'                    SetClipboardData CLng(Replace$(nowtmp, ".cbex", vbNullString)), hMem(hIndex): DoEvents
'
'                    hIndex = hIndex + 1
'                    nowtmp = Dir$
'                Loop
'
'                CloseClipboard
            End If
            
            On Error Resume Next
            'Debug.Print tmp
            Kill tmp & "*.*"
            keyrecognizing = False
            On Error GoTo h
            
            'bKey(86) = False
            cbexDoing = False
            cbexIndex = cbexIndex - 1
            If cbexIndex < -1 Then cbexIndex = -1
        End If
    End If
    keyrecognizing = False
      
h:
    Dim i%
    For i = 0 To 256
        bKey(i) = False
    Next i
    Debug.Print frmMain.SC.Error.Description
    KeyboardHookProc = CallNextHookEx(hKeyboard, nCode, wParam, ByVal lParam)
    If suicide Then Uninstall
End Function

Public Function newProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long: On Error GoTo h
    If uMsg = &H308 Then 'Ŭ������ ��ȭ ����
        'Ŭ�����庯ȭó��
        
        If cbexDoing Then GoTo h
        cbexDoing = True
        
        Dim LastFormat As Boolean
        Dim hMem As Long, bLen As Long, b() As Byte
        Dim ff%: ff = FreeFile
        
        
        
        'OpenClipboard myhwnd
        
        'LastFormat = EnumClipboardFormats(LastFormat)
        
        LastFormat = Clipboard.GetFormat(1)
        
        If LastFormat Then '������ ����Ǿ�����
            cbexIndex = cbexIndex + 1
            If cbexIndex >= value_CB_Max Then  '���۳���
                cbexIndex = value_CB_Max - 1
                cbexDoing = False
                GoTo h
            End If
            
            PrintFile CBEXFolder & "\" & cbexIndex & "\1.cbex", Clipboard.GetText
            
            
        End If
        On Error Resume Next
        If Clipboard.GetFormat(2) And value_CB_PictureAutoSave = True Then
            If isDir(App.Path & "\pic\", True) = False Then MkDir (App.Path & "\pic\")
            SavePicture Clipboard.GetData, App.Path & "\pic\" & GetTickCount & ".bmp"
        End If
        On Error GoTo h
        
'        Do While LastFormat
'
'            If LastFormat = 1 Then
'
'                hMem = GetClipboardData(LastFormat)
''                GlobalLock hMem
''                bLen = GlobalSize(hMem)
''                If bLen > 0 Then
''                    ReDim b(bLen - 1) As Byte
''                    CopyMemory b(0), ByVal hMem, bLen
''                    GlobalLock hMem 'b���ٰ� ����Ÿ�����ͼ�
''
''
''
''                    Open CBEXFolder & "\" & cbexIndex & "\" & LastFormat & ".cbex" For Binary Access Write As #FF '���˸����� ����
''                        Put #FF, , b
''                        DoEvents
''                    Close #FF
''                End If
'
'
'                DoEvents
'            End If
'            LastFormat = EnumClipboardFormats(LastFormat)
'        Loop
        
        'CloseClipboard
        Debug.Print "copied>>" & cbexIndex
        cbexDoing = False
        
        
    End If
h:
    CloseClipboard '�����ҵ� ? �ٽú���!
    newProc = CallWindowProc(OldProc, hwnd, uMsg, wParam, lParam)
End Function


Public Function Install() As Boolean: On Error GoTo h
    
    If hMouse > 0 Or hKeyboard > 0 Or OldProc > 0 Then  '�̹� ��ŷ���ϼ���������
        Uninstall
    End If
    
    'hMouse = SetWindowsHookEx(14, AddressOf MouseHookProc, App.hInstance, ByVal 0&)
    hMouse = 1 '�ƹ����� ���콺��ŷ�� �������°Ű�����... �ϴ� ����
    
    hKeyboard = SetWindowsHookEx(13, AddressOf KeyboardHookProc, App.hInstance, ByVal 0&)
    OldProc = SetWindowLong(myhwnd, -4, AddressOf newProc)
    
    If hMouse > 0 Or hKeyboard > 0 Or OldProc > 0 Then
        '��ŷ&��Ŭ���̼���
        Install = True
        suicide = False
        cbexIndex = -1
        cbexDoing = False
        keyrecognizing = False
        Exit Function
    End If
    
h:
End Function

Public Function Uninstall() As Boolean: On Error GoTo h
    UnhookWindowsHookEx hMouse
    UnhookWindowsHookEx hKeyboard
    SetWindowLong myhwnd, -4, OldProc
    Uninstall = True
    Exit Function
h:
End Function
