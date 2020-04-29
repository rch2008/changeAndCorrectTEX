Attribute VB_Name = "INI"
'文件名SourceDB.ini文件
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

Public bookPath As String

'Private Function getIniPath()
    'bookPath = "D:\test\book1\" '高三复习
    'bookPath = "D:\test\book2" '初中
    'bookPath = "D:\test\book3" '高中分册习题
'End Function

Public Function GetIniLong(ByVal Section As String, ByVal In_Key As String) As Long
    On Error GoTo GetIniLongErr
    GetIniLong = GetPrivateProfileInt(Section, In_Key, -1, bookPath) ' & "SourceDB.ini")
    If GetIniLong = -1 Then
        'MsgBox "读取数值失败"
        GoTo GetIniLongErr
    End If
    Exit Function
GetIniLongErr:
    Err.Clear
    GetIniLong = -1
End Function

Public Function WriteIniLong(ByVal Section As String, ByVal In_Key As String, ByVal In_Data As String) As Boolean
    On Error GoTo WriteIniLongErr
    WriteIniLong = True
    WritePrivateProfileString Section, In_Key, In_Data, bookPath ' & "\SourceDB.ini"
    Exit Function
WriteIniLongErr:
    MsgBox "writeerror"
    Err.Clear
    WriteIniLong = False
End Function

Public Function GetAppINI(ByVal Section As String, ByVal In_Key As String) As String
    Dim stmp As String * 1024
    Dim l As Long
    On Error GoTo GetAppINIErr
    l = GetPrivateProfileString(Section, In_Key, "", stmp, 1024, App.Path & "\setup.ini")
    If l = -1 Then
        MsgBox "读取数值失败"
        GoTo GetAppINIErr
    End If
    GetAppINI = Left(stmp, l)
    Exit Function
GetAppINIErr:
    Err.Clear
    GetAppINI = ""
End Function

Public Function WriteAppINI(ByVal Section As String, ByVal In_Key As String, ByVal In_Data As String) As Boolean
    On Error GoTo WriteAppINIErr
    WriteAppINI = True
    WritePrivateProfileString Section, In_Key, In_Data, App.Path & "\setup.ini"
    Exit Function
WriteAppINIErr:
    MsgBox "writeerror"
    Err.Clear
    WriteAppINI = False
End Function

Public Function testini(ByVal Section As String, ByVal In_Key As String, ByVal In_Data As Long)
    Dim l, m As Long
    l = 1
    If WriteIniLong(Section, In_Key, In_Data) Then
        MsgBox "Write OK!"
    End If
    m = GetIniLong(Section, In_Key)
    If m <> -1 Then
        MsgBox "m is " & m
    End If
End Function



