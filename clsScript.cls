VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
''--- Script language for project EJ7302 ---
'
''---Here is how to use this script
''1. execute [var]
''    for a file or folder. calling explorer with file or folder path.
''2. url [var]
''    open internet and connect to url.
''3. minimize/maximize
''    set all of windows to minimize or maximize.
''4. input var/ouput var
''    input and output to var. var is variation.
''    maybe, input will be InputBox() and output MsgBox()
''5. fileopen [var]/fileprint [var] [var]
''    open file with var
''    print file with var to var.
''6. capture
''    capture the window now!
''    and it will place into clipboard.
''7. mouse
''    mouseclick [left/right]
''    MouseMove [x, y]
''    mousewheel [Count]
''8. keyboard
''    keyboardpress [Var]
''9. clipboard
''    cbclear
''    cbset [Var]
''    cbget [Var]
'
''---Here is how to use this script
''1. 실행 [변수]
''    파일이나 폴더를 실행합니다.
''2. 인터넷 [변수]
''    인터넷을 접속합니다.
''3. 최소화/최대화
''    모든 윈도우를 최소화/최대화합니다.
''4. 입력 [변수] /출력 [변수]
''    변수로 입력을 받거나 출력을 합니다.
''    아마 입력은 InputBox(), 출력은 MsgBox() 일겁니다.
''5. 읽기 [변수] /쓰기 [변수]
''    파일을 읽고 씁니다.
''6. 캡쳐
''    현재 윈도우를 캡쳐합니다. 클립보드에 저장됩니다.
''7. 마우스
''    클릭 [좌/우]
''    이동 [x, y]
''    휠 [횟수]
''8. 키보드
''    누르기 [키]
''9. 클립보드
''    초기화
''    복사 [Var]
''    붙혀넣기 [Var]
'
'
'
'
''스크립트 최적화
''스크립트 run
''script typical input
''script custom input
'
'
'Public Source As String
'
'
'
'
'
'Function check() As String: On Error GoTo h
'
'    Dim i As Long
'    Dim tmp() As String
'    Dim result As String
'
'    tmp = Split(Source, vbCrLf)
'
'    For i = 0 To UBound(tmp())
'        if instR(tmp(i), "실행 ")
'    Next i
'
'
'    Exit Function
'h:
'    check = vbNullString
'End Function
