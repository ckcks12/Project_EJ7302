VERSION 5.00
Begin VB.UserControl newOCR 
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5850
   ScaleHeight     =   6315
   ScaleWidth      =   5850
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   10
      Height          =   2535
      Left            =   720
      ScaleHeight     =   2475
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   1680
      Width           =   2655
      Begin VB.Shape Box 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFF80&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape Box 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFF80&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   600
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape Box 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFF80&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   0
         Top             =   480
         Width           =   135
      End
      Begin VB.Shape Box 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFF80&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   3
         Left            =   600
         Top             =   480
         Width           =   135
      End
   End
End
Attribute VB_Name = "newOCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MDown As Boolean '���콺 �ٿ���� ���
Public BoxEnabled As Boolean  '�ڽ� ǥ�û��� ���
Private SEN As Integer ' ���� //06.27 ������ ������������. �ظ��ϸ� ����ں��ٴ� �ý��ۿ� ����ȭ�� ���ð��� �� ���ҰŰ���.
Private SENarr As Long ' �迭 ũ�� //06.27 �Ϻη� ����ڿ��� �޴ٰ� ���� ���̳� �������� ���� ���������������Ͱ���..

Private Buffer() As Byte

Private X1 As Long, X2 As Long, dX As Long
Private Y1 As Long, Y2 As Long, dY As Long


'Event Recognized(ByVal Title As String)
Event Pointed()



Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then MDown = True
    If Button = 1 Then Cls
    pic.CurrentX = X: pic.CurrentY = Y
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): On Error Resume Next
    If MDown Then
        pic.Line -(X, Y)
    End If
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MDown = False
    Point
End Sub

Private Sub UserControl_Initialize()
    '������Ʈ�� �ε�.
    SEN = 50
    SENarr = 300
    
    Cls '�ʱ�ȭ
End Sub

Private Sub UserControl_Resize()
    ResizePic
End Sub








'�ƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢ�
'�ƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢ�
'�ƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢ�
'---------------- Property Get Set Let
Public Property Let BoxColor(ByVal tmp As Long): On Error Resume Next
    Dim i%
    For i = 0 To 3
        Box(i).BorderColor = tmp
        Box(i).FillColor = tmp
    Next i
End Property

Public Property Get BoxColor() As Long: On Error Resume Next
    BoxColor = Box(0).BorderColor
End Property

Public Property Let BGI(ByVal Path As String): On Error GoTo a
    
    pic.Picture = LoadPicture(Path)
    
    Exit Property
a:
End Property

'06.27 ����ڿ��Ժ��ٴ� �ý��ۿ��� ����ȭ�� ���ð����� �ʿ���. �ϴ� ��������.
'Public Property Let Sensitivity(ByVal tmp As Integer): On Error Resume Next
'    If tmp <= 50 Or tmp >= 3000 Then
'        '�ʹ��۰ų� �ʹ�Ŀ���� �ȵȴ�!
'    Else
'        SEN = tmp
'    End If
'End Property
'
'Public Property Get Sensitivity() As Integer: On Error Resume Next
'    Sensitivity = SEN
'End Property
'-----------------








'�ƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢ�
'�ƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢ�
'�ƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢ�
'----------------- My Func and Sub
Sub ResizePic()
    pic.Left = 0
    pic.Top = 0
    pic.Width = UserControl.Width
    pic.Height = UserControl.Height
End Sub




'���� �ʱ⿡ �ν� ǥ����
'����� �ΰ�����. 1. �ʱ���ġ�� �����˱�
'                 2. Box�� ǥ���ϱ�. <<-- �̰Ŵ� �������� ǥ�þ��ϰ� �Ҽ�������.
Function Point() As Boolean
    '�ʱ� ��ġ �˾Ƴ���
    Dim i%, j%
    
    X1 = 999999
    Y1 = 999999
    X2 = 0
    Y2 = 0
    
    For i = 0 To pic.Width Step SEN
        For j = 0 To pic.Height Step SEN
            If pic.Point(i, j) = vbBlack Then
                If X1 > i Then X1 = i
                If Y1 > j Then Y1 = j
                If X2 < i Then X2 = i
                If Y2 < j Then Y2 = j
            End If
        Next j
    Next i
    
    X1 = X1 - SEN: X2 = X2 + SEN
    Y1 = Y1 - SEN: Y2 = Y2 + SEN ' �̰� SEN��ŭ ��ٺ��� �̰� ���ڶ��� ���ݾ� ©��.. �װ� �����ϱ�����.
    
    dX = X2 - X1
    dY = Y2 - Y1
    
    If dX <= SENarr Or dY <= SENarr Then '�迭�� ������ ������������ŭ ���� ���ϰ��..
        Point = False
        Exit Function
    End If
    
    ReDim Buffer(SENarr, SENarr) As Byte
    
    dX = dX / SENarr: dY = dY / SENarr '������...
    
    For i = 0 To SENarr - 1
        For j = 0 To SENarr - 1
            If pic.Point(i * dX + X1, j * dY + Y1) = vbBlack Then
                Buffer(j, i) = 1
            Else
                Buffer(j, i) = 0
            End If
        Next j
    Next i
    
    '---Boxǥ���ϴºκ�
        'If BoxEnabled Then
            Box(0).Left = X1 - Box(0).Width: Box(0).Top = Y1 - Box(0).Height
            Box(1).Left = X2: Box(1).Top = Y1 - Box(1).Height
            Box(2).Left = X1 - Box(2).Width: Box(2).Top = Y2
            Box(3).Left = X2: Box(3).Top = Y2
        'End If
    '---
    
    Point = True
    RaiseEvent Pointed
End Function


'Ŭ������� ���� Buffer���� ���.
Sub PrintBufferToClipboard()
    Dim i%, j%, tmp$
    
    'ReDim Preserve Buffer(SENarr, SENarr) As Byte
    
    For i = 0 To SENarr - 1
        For j = 0 To SENarr - 1
            tmp = tmp & Buffer(i, j)
        Next j
        tmp = tmp & vbCrLf
    Next i
a:
    On Error GoTo a
    Clipboard.Clear
    Clipboard.SetText tmp
End Sub

'���� Buffer�� ��ȯ
Function Peek() As Byte(): On Error Resume Next
    Peek = Buffer
End Function

'���� Buffer�� ���Ϸ�
Function Save(ByVal Path As String) As Boolean: On Error GoTo a
    Dim ff As Integer
    ff = FreeFile
    Close #ff
    Open Path For Binary Access Write As #ff
        Put #ff, , Buffer
    Close #ff
    
    Save = True
    Exit Function
a:
    Save = False
End Function


'�ʱ�ȭ
Sub Cls()
    Dim i%

    For i = 0 To 3
        Box(i).Left = -Box(i).Width
        Box(i).Top = -Box(i).Top
    Next i
    
    ReDim Buffer(SENarr, SENarr) As Byte
    pic.Cls
End Sub




'�ν�
'�� Buffer����Ÿ�� �ָ� �װŶ� ����.
Function Recognize(ByRef newBuffer() As Byte, Optional ByRef percent As Integer) As Boolean: On Error GoTo h
    Dim i As Integer, j As Integer
    Dim both As Long, First As Long, Second As Long, either As Long
    Dim tmp As Long
    'ReDim Preserve newBuffer(SENarr, SENarr) As Byte
    'ReDim Preserve Buffer(SENarr, SENarr) As Byte
    
    For i = 0 To SENarr - 1
        either = 0: First = 0: Second = 0: both = 0
        For j = 0 To SENarr - 1
        
            Select Case Buffer(i, j) + (newBuffer(i, j) * 2)
                '0 either 1 first 2 second 3 both
                Case 0: either = either + 1&
                Case 1: First = First + 1&
                Case 2: Second = Second + 1&
                Case 3: both = both + 1&
            End Select
        Next j
        If (both * 4) + (First * 2) - (Second * 3) + either > 230 Then tmp = tmp + 1
        
    Next i
    
    
    percent = Int(tmp / 3)
    Recognize = True
    Exit Function
h:
    Recognize = False
End Function


'���Ϸ�����
Function SaveTo(ByVal Path As String) As Boolean: On Error GoTo h
    
    SavePicture pic.Image, Path
    SaveTo = True
    Exit Function
h:
    SaveTo = False
End Function

'�ƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢ�
'�ƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢ�
'�ƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢ�
