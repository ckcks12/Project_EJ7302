VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmCBList 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmCBList"
   ClientHeight    =   8640
   ClientLeft      =   2190
   ClientTop       =   -390
   ClientWidth     =   10590
   LinkTopic       =   "frmCBList"
   ScaleHeight     =   8640
   ScaleWidth      =   10590
   Tag             =   "8640 11535"
   WindowState     =   2  'Maximized
   Begin MSScriptControlCtl.ScriptControl SC 
      Left            =   3120
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
End
Attribute VB_Name = "frmCBList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim source As String

Public f As String

Private Sub Form_Load(): On Error GoTo h
    
    
    source = OpenFile(App.Path & "\scriptex.dll")
    source = source & vbCrLf & OpenFile(App.Path & "\scriptdiy.dll")
    SC.AddCode source
    
    SC.Run f
    
    frmMain.Show
    Unload Me
    
    Exit Sub
    
    
h:
    If SC.Error.Number <> 0 Then
        MsgBox SC.Error.Description, vbCritical + vbOKOnly + vbSystemModal, SC.Error.Number
    End If
    frmMain.Show
    frmMain.Hide
    Unload Me
End Sub
