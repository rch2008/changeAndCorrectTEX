Attribute VB_Name = "文件夹选择"
Option Explicit '文件夹选择框

Private Const BIF_STATUSTEXT = &H4&

Private Const BIF_USENEWUI = &H40

Private Const BIF_RETURNONLYFSDIRS = 1

Private Const BIF_DONTGOBELOWDOMAIN = 2

Private Const MAX_PATH = 260

Private Const WM_USER = &H400

Private Const BFFM_INITIALIZED = 1

Private Const BFFM_SELCHANGED = 2

Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)

Private Const BFFM_SETSelectION = (WM_USER + 102)

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList _
                Lib "shell32" (ByVal pidList As Long, _
                               ByVal lpBuffer As String) As Long

Private Declare Function lstrcat _
                Lib "kernel32" _
                Alias "lstrcatA" (ByVal lpString1 As String, _
                                  ByVal lpString2 As String) As Long

Private Type BrowseInfo

    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long

End Type

Private m_CurrentDirectory As String   'The current directory

Public Function BrowseForFolder(owner As Form, _
                                Title As String, _
                                StartDir As String) As String

    Dim lpIDList    As Long

    Dim szTitle     As String

    Dim sBuffer     As String

    Dim tBrowseInfo As BrowseInfo

    m_CurrentDirectory = StartDir & vbNullChar
    szTitle = Title

    With tBrowseInfo
        .hWndOwner = owner.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_STATUSTEXT '如果希望对话框中有“新建文件夹”,那么就给.ulFlags 加上BIF_USENEWUI属性,BIF_RETURNONLYFSDIRS 的意思是仅仅返回文件夹
        .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)  'get address of function.
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        BrowseForFolder = sBuffer
    Else
        BrowseForFolder = ""
    End If

End Function

Private Function BrowseCallbackProc(ByVal hWnd As Long, _
                                    ByVal uMsg As Long, _
                                    ByVal lp As Long, _
                                    ByVal pData As Long) As Long

    Dim lpIDList As Long

    Dim ret      As Long

    Dim sBuffer  As String

    On Error Resume Next

    Select Case uMsg

        Case BFFM_INITIALIZED
            Call SendMessage(hWnd, BFFM_SETSelectION, 1, m_CurrentDirectory)

        Case BFFM_SELCHANGED
            sBuffer = Space(MAX_PATH)
            ret = SHGetPathFromIDList(lp, sBuffer)

            If ret = 1 Then
                Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
            End If

    End Select

    BrowseCallbackProc = 0
End Function

Private Function GetAddressofFunction(add As Long) As Long

    GetAddressofFunction = add
End Function
