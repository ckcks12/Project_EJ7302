VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmOCRmaina 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5430
   LinkTopic       =   "Form2"
   ScaleHeight     =   3435
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSScriptControlCtl.ScriptControl SC 
      Left            =   4800
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin projectEJ7302.newOCR newOCR 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6165
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   3000
      Pattern         =   "*.dat"
      TabIndex        =   1
      Top             =   1800
      Width           =   735
   End
End
Attribute VB_Name = "frmOCRmaina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Sub Form_Load()
    Unload frmMain
    FadeIN Me
    File1.Path = App.Path & "\ocr\"
    File1.Refresh
End Sub

Private Sub newOCR_Pointed(): On Error GoTo h
    '파일리스트에잇는거 다가꼬옴
    Dim i%, b(300, 300) As Byte, B1(300, 300) As Byte, B2(300, 300) As Byte, B0(300, 300) As Byte, ff%: ff = FreeFile: Dim percent As Integer, percent0%, percent1%, percent2%, percent3%, percent4%
    Dim B3(300, 300) As Byte, B4(300, 300) As Byte
    Dim source As String
    Dim best As String
    Dim bestPercent As Integer
    Dim thisPercent As Integer
    
    For i = 0 To File1.ListCount - 1&
        
        Open File1.Path & "\" & File1.List(i) For Binary Access Read As #ff
            Get #ff, , B0
        Close #ff
        
        If Not newOCR.Recognize(B0, percent0) Then GoTo h
        
        If percent0 > 40 Then
            
'            Open File1.Path & "\" & Replace$(File1.List(i), ".dat", vbNullString) & "\0.dat" For Binary Access Read As #FF
'                Get #FF, , B0
'            Close #FF
'
'            If Not newOCR.Recognize(B0, percent0) Then GoTo h
            
            Open File1.Path & "\" & Replace$(File1.List(i), ".dat", vbNullString) & "\1.dat" For Binary Access Read As #ff
                Get #ff, , B1
            Close #ff
            If Not newOCR.Recognize(B1, percent1) Then GoTo h
            
            Open File1.Path & "\" & Replace$(File1.List(i), ".dat", vbNullString) & "\2.dat" For Binary Access Read As #ff
                Get #ff, , B2
            Close #ff
            If Not newOCR.Recognize(B2, percent2) Then GoTo h
            
            Open File1.Path & "\" & Replace$(File1.List(i), ".dat", vbNullString) & "\3.dat" For Binary Access Read As #ff
                Get #ff, , B3
            Close #ff
            If Not newOCR.Recognize(B3, percent3) Then GoTo h
            
            Open File1.Path & "\" & Replace$(File1.List(i), ".dat", vbNullString) & "\4.dat" For Binary Access Read As #ff
                Get #ff, , B4
            Close #ff
            If Not newOCR.Recognize(B4, percent4) Then GoTo h
            
            
            
            
            thisPercent = max5(percent0, percent1, percent2, percent3, percent4)
            If percent0 > 70 Then
                'If percent1 > 70 Or percent2 > 70 Then
                    If thisPercent > bestPercent Then
                        best = File1.Path & "\" & Replace$(File1.List(i), ".dat", vbNullString) & "\command.txt"
                        bestPercent = thisPercent
                    End If
                'End If
            'ElseIf percent1 > 70 And percent2 > 70 Then
            ElseIf percent1 > 70 Or percent2 > 70 Or percent3 > 70 Or percent4 > 70 Then
                If thisPercent > bestPercent Then
                    best = File1.Path & "\" & Replace$(File1.List(i), ".dat", vbNullString) & "\command.txt"
                    bestPercent = thisPercent
                End If
            End If
            
'            If percent0 >= 90 Or percent1 >= 90 Or percent2 >= 90 Then
'                If thisPercent >= bestPercent Then
'                    best = File1.Path & "\" & Replace$(File1.List(i), ".dat", vbNullString) & "\command.txt"
'                    bestPercent = thisPercent
'                End If
'            End If



'MsgBox i + 1 & vbCrLf & "1 " & percent0 & vbCrLf & "2 " & percent1 & vbCrLf & "3 " & percent2
        End If
        
        
        
        
    Next i
    
    If best = vbNullString Then GoTo w
    
    '---여기서 미리 작업 ..
                    source = OpenFile(App.Path & "\scriptex.dll")
                    source = source & vbCrLf & OpenFile(App.Path & "\scriptdiy.dll")
    '---
             
    Debug.Print best
    source = source & vbCrLf & OpenFile(best)
    
    SC.AddCode source
    SC.Run "main"
    
w:
    frmMain.Show
    Unload Me
    Exit Sub
h:
    MsgBox "인식 중 오류가 발생했습니다", vbCritical + vbOKOnly + vbSystemModal, "오류"
    If SC.Error.Number <> 0 Then MsgBox SC.Error.Description, vbCritical + vbSystemModal + vbOKOnly, SC.Error.Number
        
    frmMain.Show
    Unload Me
End Sub
