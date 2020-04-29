Attribute VB_Name = "selectFilesAPI"
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOPENFILENAME As OPENFILENAME) As Long

Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_EXPLORER = &H80000
Public allSelectFiles() As String

Public Function selectFile(ByVal index As Integer) As String()
    On Error GoTo myError
    Dim rtn As Long, pos As Integer
    Dim file As OPENFILENAME
    file.lStructSize = Len(file)
    file.hInstance = App.hInstance
    file.lpstrFile = String$(255, 0)
    file.nMaxFile = 255
    file.lpstrFileTitle = String$(255, 0)
    file.nMaxFileTitle = 255
    file.lpstrInitialDir = ""
    file.lpstrFilter = "tex Files(*.tex)" & Chr$(0) & "*.tex" & Chr$(0) & _
                       "Word (*.docx)" & Chr$(0) & "*.docx" & Chr$(0) & _
                       "txt (*.txt)" & Chr$(0) & "*.txt" & Chr$(0) & _
                       "INI (*.ini)" & Chr$(0) & "*.ini" & Chr$(0) & _
                       "所有文件(*.*)" & Chr$(0) & "*.*" & Chr$(0)
    file.nFilterIndex = index
    file.lpstrTitle = "选择文件"

    Dim str() As String
    Dim sNothing(0) As String
    Dim strTemp As String
    sNothing(0) = ""
    selectFile = sNothing
    'fileDlg.flags = &H80200

    file.flags = OFN_EXPLORER + OFN_ALLOWMULTISELECT

    rtn = GetOpenFileName(file)

    If rtn > 0 Then
        strTemp = Replace(file.lpstrFile, Chr(0) + Chr(0), "")
        If Right(strTemp, 1) = Chr(0) Then
            strTemp = Left(strTemp, Len(strTemp) - 1)
        End If
        str = Split(strTemp, Chr(0))      ' 取得附件路径
        t = UBound(str)
        If t > 0 Then
            For i = 1 To t
                str(i) = str(0) & "\" & str(i)
            Next i
            str(0) = ""
            strTemp = Join(str, "?")
            If Right(strTemp, 1) = "?" Then
                strTemp = Mid(strTemp, 2, Len(strTemp) - 2)
            Else
                strTemp = Mid(strTemp, 2, Len(strTemp) - 1)
            End If
            str = Split(strTemp, "?")
        End If
        selectFile = str
    End If

    Exit Function

myError:
    MsgBox "操作失败!", vbCritical + vbOKOnly, APP_NAME

End Function
