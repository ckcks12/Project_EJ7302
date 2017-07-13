VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmMain"
   ClientHeight    =   8640
   ClientLeft      =   2190
   ClientTop       =   -390
   ClientWidth     =   10590
   LinkTopic       =   "frmMain"
   ScaleHeight     =   8640
   ScaleWidth      =   10590
   Tag             =   "8640 11535"
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6360
      Top             =   1440
   End
   Begin MSScriptControlCtl.ScriptControl SC 
      Left            =   2040
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin projectEJ7302.newButton newButton3 
      Height          =   3255
      Left            =   3600
      TabIndex        =   2
      Top             =   2400
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5741
   End
   Begin projectEJ7302.newButton newButton2 
      Height          =   3255
      Left            =   4920
      TabIndex        =   1
      Top             =   4320
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5741
   End
   Begin projectEJ7302.newButton newButton1 
      Height          =   3255
      Left            =   1920
      TabIndex        =   0
      Top             =   4080
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5741
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Declare Function SetClipboardViewer Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ChangeClipboardChain Lib "user32" (ByVal hwnd As Long, ByVal hWndNext As Long) As Long

 
Dim hCBChain As Long
Dim source As String

Private Sub Form_Activate()
    Me.Hide
    
    Uninstall
    Install
    
    
    '스크립트 불러옹기
    SC.Reset
    source = OpenFile(App.Path & "\scriptex.dll")
    source = source & vbCrLf & OpenFile(App.Path & "\scriptdiy.dll")
    SC.AddCode source
End Sub

Private Sub Form_Load()
    Me.Hide
    'FadeIN Me
    
    
    
    'myhwnd초기화 -- 클립보드체인감지용
    myhwnd = hwnd
    
    
    hCBChain = SetClipboardViewer(myhwnd)
    Install 'hook and subclassing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ChangeClipboardChain hwnd, hCBChain
    Uninstall
End Sub

Public Sub Form1ShowFromfrmMain()
    ChangeClipboardChain hwnd, hCBChain
    Uninstall
    Form1.Show
    Unload Me
End Sub

Private Sub Timer1_Timer()
    mod_hook.keyrecognizing = False
    mod_hook.cbexDoing = False
    Uninstall
    Install
End Sub
