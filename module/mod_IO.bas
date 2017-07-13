Attribute VB_Name = "mod_IO"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


'###폴더와 파일 있는지없는지 검사
Function isDir(ByVal Path As String, Optional ByVal Directory As Boolean = False) As Boolean: On Error GoTo whaterr
    If Directory Then '폴더검사
        If Len(Dir(Path & "\", vbDirectory)) = 0 Then GoTo whaterr
    Else '일반파일검사 (읽기전용이나 시스템파일의 경우엔 직접수정해서사용
        If Len(Dir(Path)) = 0 Then GoTo whaterr
    End If
    
    isDir = True
    Exit Function
    
whaterr:
    isDir = False
    Exit Function
End Function

'### 파일 열기
Function OpenFile(ByVal Path As String, Optional ByVal isList As Boolean, Optional ByRef List As ListBox, Optional ByVal Delimiter As String) As String: On Error GoTo whaterr
    If Not isDir(Path) Then GoTo whaterr '파일유무검사
    Dim tmp                 As String
    Dim tmp2                As String
    Dim tmp3()              As String
    Dim i                   As Long
    
    With CreateObject("Scripting.FileSystemObject") '파일 읽어오기
        With .opentextfile(Path)
            tmp2 = .readall
        End With
    End With
    
    If isList Then '리스트에 추가시켜줘야한다면
        tmp3 = Split(tmp2, Delimiter)
        
        For i = 0 To UBound(tmp3)
            List.AddItem tmp3(i)
        Next i
        OpenFile = tmp2
        Exit Function
    End If
    
    OpenFile = tmp2
    Exit Function
    
whaterr:
    OpenFile = vbNullString
    Exit Function
End Function

'### 파일 쓰기
Function PrintFile(ByVal Path As String, ByVal str As String, Optional ByVal isTime As Boolean = False) As Boolean: On Error GoTo whaterr
    If Len(Path) = 0 Then GoTo whaterr
    
    Dim FF                  As Integer: FF = FreeFile
    Open Path For Append As #FF
    
    If isTime Then '시간기록하기
      
        Print #FF, time & vbTab & vbTab & str
            
    Else
        
        Print #FF, str
        
    End If
    
    Close #FF
    
    PrintFile = True
    Exit Function
    
whaterr:
    PrintFile = False
    Exit Function
End Function

'### INI 파일 쓰기
Function PrintINI(ByVal Path As String, ByVal First As String, ByVal Second As String, ByVal content As String) As Boolean: On Error GoTo whaterr
    If WritePrivateProfileString(First, Second, content, Path) Then
        PrintINI = True
    Else
whaterr:
        PrintINI = False
    End If
End Function

'### INI 파일 읽기
Function OpenINI(ByVal Path As String, ByVal First As String, ByVal Second As String, Optional ByVal Default As String = "") As String: On Error GoTo whaterr
    If Not isDir(Path) Then GoTo whaterr
    
    Dim BufLen              As Long
    Dim tmp$
    tmp = Space$(FileLen(Path))
    BufLen = GetPrivateProfileString(First, Second, Default, tmp, 20, Path)
    If BufLen = 0 Then GoTo whaterr

    OpenINI = Left$(tmp, BufLen + 1&)
    Exit Function
whaterr:
    OpenINI = Default
End Function
