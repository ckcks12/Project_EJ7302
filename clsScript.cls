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
''1. ���� [����]
''    �����̳� ������ �����մϴ�.
''2. ���ͳ� [����]
''    ���ͳ��� �����մϴ�.
''3. �ּ�ȭ/�ִ�ȭ
''    ��� �����츦 �ּ�ȭ/�ִ�ȭ�մϴ�.
''4. �Է� [����] /��� [����]
''    ������ �Է��� �ްų� ����� �մϴ�.
''    �Ƹ� �Է��� InputBox(), ����� MsgBox() �ϰ̴ϴ�.
''5. �б� [����] /���� [����]
''    ������ �а� ���ϴ�.
''6. ĸ��
''    ���� �����츦 ĸ���մϴ�. Ŭ�����忡 ����˴ϴ�.
''7. ���콺
''    Ŭ�� [��/��]
''    �̵� [x, y]
''    �� [Ƚ��]
''8. Ű����
''    ������ [Ű]
''9. Ŭ������
''    �ʱ�ȭ
''    ���� [Var]
''    �����ֱ� [Var]
'
'
'
'
''��ũ��Ʈ ����ȭ
''��ũ��Ʈ run
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
'        if instR(tmp(i), "���� ")
'    Next i
'
'
'    Exit Function
'h:
'    check = vbNullString
'End Function