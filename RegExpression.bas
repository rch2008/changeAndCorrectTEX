Attribute VB_Name = "RegExpression"
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public questionID As Long
Public tabID As Long
Public figID As Long
Public tempID(2) As Long
Public strFullName() As String
Public strReplaceList() As String
Public strReplaceSymbolList() As String
Public strTestName As String
Public replaceListFile As String
Public replaceSymbolListFile As String
Public braceCMDList As String
Public questionTypeXZ As String
Public questionTypeTK As String
Public questionTypeJD As String
Public answerBoundary As String
Public questionAnswerBoundary As String
Public correctEnvironments() As String
Public mathScriptList() As String
Public delDollerList() As String

Public answerSolutionBoundary As String
Public solutionStart As String
Public solutionEnd As String

Public docxToTexPath As String
Public addDollorFlag As Boolean
Public correctLeftRightFlag As Boolean
Public Counter As Long
Public ifReadINI As Boolean
Public fileSelect As Boolean
Public needLeftRightList As String
Public Unidentified As Long
Public strUnidentified As String
Private firstQuestionFlag As Boolean
Function init()
    Unidentified = 0
    fileSelect = False
End Function

Function onlyToTex()
    Dim PID As Long
    Dim arr() As String
    Dim larr As Integer
    Dim strPathName, strName As String
    Dim docxFileName
    init
    strFullName = selectFile(2)
    If strFullName(0) = "" Then
        'MsgBox "未选择docx文件"
        Exit Function
    End If
    For Each docxFileName In strFullName
        arr = Split(docxFileName, ".")
        strPathName = arr(0)
        arr = Split(arr(0), "\")
        strName = arr(UBound(arr))
            
        PID = Shell(docxToTexPath + "convert.bat " + docxFileName, 1)
    
        If PID <> 0 Then
            hProcess = OpenProcess(&H100000, True, PID)
            WaitForSingleObject hProcess, -1
            CloseHandle hProcess
        End If
        Shell docxToTexPath + "CopyAndDel.bat " + strPathName
    Next
End Function

Function changeToTex() As String
    Dim PID As Long
    Dim arr() As String
    Dim larr As Integer
    Dim strPathName, strName As String
    Dim docxFileName
    init
    'strFullName = selectFile(2)
    beforeChange
    If strFullName(0) = "" Then
        'MsgBox "未选择docx文件"
        Exit Function
    End If
    On Error GoTo err1
    For Each docxFileName In strFullName
        arr = Split(docxFileName, ".")
        strPathName = arr(0)
        arr = Split(arr(0), "\")
        strName = arr(UBound(arr))
            
        PID = Shell(docxToTexPath + "convert.bat " + docxFileName, 1)
        t = Timer()
        If PID <> 0 Then
            hProcess = OpenProcess(&H100000, True, PID)
            WaitForSingleObject hProcess, -1
            CloseHandle hProcess
        End If
        changeToTex = Timer() - t
        Shell docxToTexPath + "CopyAndDel.bat " + strPathName
    Next
    Exit Function
err1:
    MsgBox "请查看convert.bat CopyAndDel.bat 是否存在，路径是否正确！"
End Function

Function convertToTex()
    Dim PID As Long
    Dim arr() As String
    Dim larr As Integer
    Dim strPathName, strName As String
    Dim docxFileName
    init
    'strFullName = selectFile(2)
    beforeChange
    If strFullName(0) <> "" Then
        fileSelect = True
    Else
        'MsgBox "未选择docx文件"
        Exit Function
    End If
    For Each docxFileName In strFullName
        arr = Split(docxFileName, ".")
        strPathName = arr(0)
        arr = Split(arr(0), "\")
        strName = arr(UBound(arr))
            
        PID = Shell(docxToTexPath + "convert.bat " + docxFileName, 1)
    
        If PID <> 0 Then
            hProcess = OpenProcess(&H100000, True, PID)
            WaitForSingleObject hProcess, -1
            CloseHandle hProcess
        End If
        Shell docxToTexPath + "CopyAndDel.bat " + strPathName
        
        Main
        
    Next
    fileSelect = False
End Function

Function Main() As String
    Dim doc As String
    Dim str As String
    Dim arr() As String
    Dim finalStr As String
    Dim allFilesName As String
    Dim texFileName
    allFilesName = ""
    strUnidentified = ""
    If fileSelect = False Then
        strFullName = selectFile(1)
    End If
    If strFullName(0) <> "" Then
        tempID(0) = questionID
        tempID(1) = figID
        tempID(2) = tabID
        t = Timer()
        For Each texFileName In strFullName
            finalStr = ""
            readUTF8 texName(texFileName), doc
            cutDocument doc        '去头尾，去双换行，删除带括号命令
            replaceSymbol doc, texFileName       '替换，删除一些符号
            correctUDscript doc    '双上下标
            cutXTJ doc, finalStr, texFileName
            arr = Split(texFileName, ".")
            writeTex finalStr, Left(texFileName, Len(texFileName) - Len(arr(UBound(arr))) - 1) + "_VBA.tex"
            'Debug.Print finalStr
            If Unidentified <> 0 Then
                writeTex strUnidentified, Left(texFileName, Len(texFileName) - Len(arr(UBound(arr))) - 1) + "_VBA.txt"
            End If
        Next
        Main = Timer() - t
        If ifReadINI = True Then
            If MsgBox("是否将ID号写入INI文件？" + Chr(13) + "是，ID号写入INI文件。" + Chr(13) + "否，ID号不写入INI文件", vbYesNo, "是否写入INI文件") = vbYes Then
                For Each texFileName In strFullName
                    allFilesName = allFilesName + Chr(13) + texFileName
                Next
                writeINI allFilesName
            Else
                questionID = tempID(0)
                figID = tempID(1)
                tabID = tempID(2)
            End If
        End If
    End If
End Function

Function JoinTest(Optional joinFlag As Boolean = False) As String
    Dim doc As String
    Dim str As String
    Dim arr() As String
    Dim finalStr As String
    Dim allFilesName As String
    Dim texFileName
    allFilesName = ""
    strUnidentified = ""
    If fileSelect = False And joinFlag = False Then
        strFullName = selectFile(1)
    End If
    If strFullName(0) <> "" Then
        For Each texFileName In strFullName
            finalStr = ""
            readUTF8 texName(texFileName), doc
            cutDocument doc        '去头尾，去双换行，删除带括号命令
            replaceSymbol doc, texFileName       '替换，删除一些符号
            correctUDscript doc    '双上下标
            joinXTJ doc, finalStr, texFileName, joinFlag
            arr = Split(texFileName, ".")
            If joinFlag Then
                writeTex finalStr, Left(texFileName, Len(texFileName) - Len(arr(UBound(arr))) - 1) + "_Join.tex"
            End If
        Next
    End If
    JoinTest = finalStr
End Function

Function delBraceCMDInFiles(Optional joinFlag As Boolean = False) As String
    Dim doc As String
    Dim str As String
    Dim arr() As String
    Dim finalStr As String
    Dim allFilesName As String
    Dim texFileName
    allFilesName = ""
    strUnidentified = ""
    If fileSelect = False And joinFlag = False Then
        strFullName = selectFile(1)
    End If
    If strFullName(0) <> "" Then
        For Each texFileName In strFullName
            finalStr = ""
            readUTF8 texName(texFileName), doc
            delBraceCMD doc      '替换，删除一些符号
            writeTex doc, CStr(texFileName)
        Next
    End If
    delBraceCMDInFiles = doc
End Function

Function joinXTJ(ByRef doc As String, ByRef finalStr As String, ByVal texFileName As String, joinFlag As Boolean)
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim strTemp, str As String
    Dim strXZ As String
    Dim strTK As String
    Dim strJD As String
    Dim arr() As String
    Dim t() As String
    Dim docQuestion As String
    Dim arrAnswer() As String
    Dim index As Long
    Dim jFlag As String
    
    If joinFlag Then
        jFlag = "Not empty"
    Else
        jFlag = ""
    End If
    
    Set re = New RegExp
    strXZ = ""
    strTK = ""
    strJD = ""
    
    re.Global = True
    
    re.Pattern = "(\n|\r|" + Chr(13) + ")" + answerBoundary
    doc = re.Replace(doc, Chr(0))
    t = Split(doc, Chr(0))
    docQuestion = t(0)
    arrAnswer = cutApart(t(1), jFlag)
    index = 1
    
    re.Pattern = "(\n|\r|" + Chr(13) + ")" + "\S{0,2}(" + questionTypeXZ + "|" + questionTypeTK + "|" + questionTypeJD + ")"

    If re.test(docQuestion) Then
        firstQuestionFlag = True
        Set mMatches = re.Execute(docQuestion)
        strTestName = Mid(docQuestion, 1, mMatches(0).FirstIndex)
        For i = 0 To mMatches.Count - 1
            strTemp = mMatches(i).Value
            If QuestionType(strTemp, questionTypeXZ) = True Then
            '选择题
                If i + 1 < mMatches.Count Then
                    strXZ = Mid(docQuestion, mMatches(i).FirstIndex + 1, mMatches(i + 1).FirstIndex - mMatches(i).FirstIndex)
                Else
                    strXZ = Mid(docQuestion, mMatches(i).FirstIndex + 1)
                End If
                joinQA cutApart(strXZ, ""), finalStr, arrAnswer, index, joinFlag
            ElseIf QuestionType(strTemp, questionTypeTK) = True Then
            '填空题
                If i + 1 < mMatches.Count Then
                    strTK = Mid(docQuestion, mMatches(i).FirstIndex + 1, mMatches(i + 1).FirstIndex - mMatches(i).FirstIndex)
                Else
                    strTK = Mid(docQuestion, mMatches(i).FirstIndex + 1)
                End If
                joinQA cutApart(strTK, ""), finalStr, arrAnswer, index, joinFlag
            ElseIf QuestionType(strTemp, questionTypeJD) = True Then
                '解答题
                If i + 1 < mMatches.Count Then
                    strJD = Mid(docQuestion, mMatches(i).FirstIndex + 1, mMatches(i + 1).FirstIndex - mMatches(i).FirstIndex)
                Else
                    strJD = Mid(docQuestion, mMatches(i).FirstIndex + 1)
                End If
                joinQA cutApart(strJD, ""), finalStr, arrAnswer, index, joinFlag
            End If
            DoEvents
        Next i
        finalStr = "\begin{document}" + Chr(13) + strTestName + finalStr + "\end{document}"
    Else
    
    End If
End Function

Function QuestionType(ByVal str As String, ByVal strPattern As String) As Boolean
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Set re = New RegExp
    re.Pattern = strPattern
    Set mMatches = re.Execute(TrimEnter(str))
    If mMatches.Count > 0 Then
        If mMatches(0).FirstIndex < 3 Then
            QuestionType = True
            Exit Function
        End If
    Else
        QuestionType = False
    End If
End Function
Function joinQA(ByRef strQuestionAndAnswer() As String, ByRef finalStr As String, ByRef strAnswer() As String, ByRef index As Long, Optional joinFlag As Boolean = False)
    Dim jx As String
    If joinFlag Then
        jx = questionAnswerBoundary
    Else
        jx = ""
    End If
    finalStr = finalStr + strQuestionAndAnswer(0)
    For i = 1 To UBound(strQuestionAndAnswer)
        finalStr = finalStr + Chr(13) + strQuestionAndAnswer(i) + Chr(13) + jx + strAnswer(index)
        index = index + 1
    Next

End Function

Function beforeChange()
    On Error GoTo err1
    Dim mySelection As Word.Selection
    Dim wApp As Word.Application
    Dim docxFileName
    Set wApp = CreateObject("Word.Application")
    wApp.Visible = False
    strFullName = selectFile(2)
    If strFullName(0) <> "" Then
        For Each docxFileName In strFullName
            Err.Clear
            Set wDoc = wApp.Documents.Open(docxFileName)
            Set mySelection = wApp.Documents.Application.Selection
            mySelection.Find.MatchWildcards = False
            mySelection.Find.MatchWildcards = True
            Call mySelection.Find.Execute("([0-9]@．)（[0-9]{1,2}分）", False, False, True, False, False, True, wdFindContinue, False, "\1", wdReplaceAll, False, False, False, False)
            Call mySelection.Find.Execute("声明：*25151492", False, False, True, False, False, True, wdFindContinue, False, "", wdReplaceAll, False, False, False, False)
'            mySelection.Find.Font.Underline = wdUnderlineSingle
'            Call mySelection.Find.Execute("([! |　])@", False, False, True, False, False, True, wdFindContinue, True, " ", wdReplaceAll, False, False, False, False)
            mySelection.Find.Replacement.Font.Underline = wdUnderlineNone
            Call mySelection.Find.Execute("([ 　])@", False, False, True, False, False, True, wdFindContinue, True, "_", wdReplaceAll, False, False, False, False)
            mySelection.Find.MatchWildcards = False
            If Err.Number <> 0 Then
                Debug.Print Err
            End If
            mySelection.WholeStory          '选择文档全部内容
            mySelection.Font.Bold = False   '去粗体
            mySelection.Font.Italic = False '去斜体
        
            mySelection.Font.Color = wdColorBlack
            wApp.ActiveDocument.SaveAs FileName:=docxFileName
            wApp.ActiveDocument.Close
            wApp.RecentFiles(1).Delete
        Next
    End If
    Set wApp = Nothing
    Exit Function
err1:
    MsgBox Err.Number
    strFullName = selectFile(2)
End Function
Function texName(ByVal str As String) As String
    Dim arr() As String
    arr = Split(str, ".")
    arr(UBound(arr)) = "tex"
    texName = Join(arr, ".")
End Function
Function cutDocument(ByRef str As String) As String
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    Dim s As String
    
    Set re = New RegExp
    re.Pattern = "begin\{document\}|\\end\{document\}"
    re.Global = True
    s = str
    Call delBraceCMD(s)
    DelTensor s
    Set mMatches = re.Execute(s)
    If mMatches.Count = 2 Then
        s = Mid(s, mMatches(0).FirstIndex + mMatches(0).Length + 1, mMatches(1).FirstIndex - mMatches(0).FirstIndex - mMatches(0).Length)
        re.Pattern = "\$"
        s = re.Replace(s, "")
        re.Pattern = "(\n|\r|" + Chr(13) + ")+"  '
        s = re.Replace(s, Chr(13))
        re.Pattern = "声明：(.)+25151492"
        s = re.Replace(s, "")
    Else
        MsgBox "document环境不匹配！"
    End If
    str = s
End Function

Function delBraceCMD(ByRef s As String) As String
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    Dim str As String
    Dim prev As Long
    
    prev = 1
    Set re = New RegExp
    re.Pattern = braceCMDList '"\\textbf\{|\\text\{|\\textit\{|\\mathrm\{|\\boldsymbol\{|\\textcolor\{color-[0-9]\}\{|\\underline\{"
    
    re.Global = True
    Set mMatches = re.Execute(s)
    Do While mMatches.Count > 0
        For Each mMatch In mMatches
            If prev < mMatch.FirstIndex Then
                str = str + Mid(s, prev, mMatch.FirstIndex + 1 - prev)
                prev = nextRightBrace(mMatch.FirstIndex + mMatch.Length + 1, s)
                str = str + Mid(s, mMatch.FirstIndex + mMatch.Length + 1, prev - mMatch.FirstIndex - mMatch.Length - 1)
                prev = prev + 1
            End If
            DoEvents
        Next
        str = str + Mid(s, prev)
        s = str
        Set mMatches = re.Execute(s)
        prev = 1
        str = ""
   Loop
End Function

Function Redistribution()
    Dim doc As String
    Dim allFilesName As String
    Dim texFileName
    allFilesName = ""
    strUnidentified = ""
    If fileSelect = False Then
        strFullName = selectFile(1)
    End If
    If strFullName(0) <> "" Then
        tempID(0) = questionID
        tempID(1) = figID
        tempID(2) = tabID
        For Each texFileName In strFullName
            doc = ""
            readUTF8 CStr(texFileName), doc
            
            'doc = redistributeQuestionID(doc)
            'doc = redistributeF_T_ID(doc, "FigID")
            'doc = redistributeF_T_ID(doc, "TabID")
            redistributeQuestionID doc
            If redistributeF_T_ID(doc, "FigID") = False Then
                writeTex figID & doc, Split(CStr(texFileName), ".")(0) & ".txt"
                Exit Function
            End If
            If redistributeF_T_ID(doc, "TabID") = False Then
                writeTex tabID & doc, Split(CStr(texFileName), ".")(0) & ".txt"
                Exit Function
            End If

            writeTex doc, CStr(texFileName)
        Next
        If ifReadINI = True Then
            If MsgBox("是否将ID号写入INI文件？" + Chr(13) + "是，ID号写入INI文件。" + Chr(13) + "否，ID号不写入INI文件", vbYesNo, "是否写入INI文件") = vbYes Then
                For Each texFileName In strFullName
                    allFilesName = allFilesName + Chr(13) + texFileName
                Next
                writeINI allFilesName
            Else
                questionID = tempID(0)
                figID = tempID(1)
                tabID = tempID(2)
            End If
        End If
    End If
End Function

Function redistributeQuestionID(ByRef s As String) As String
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    Dim str As String
    Dim prev As Long
    
    prev = 1
    Set re = New RegExp
    re.Pattern = "\\item[XTJF]\{"
    
    re.Global = True
    Set mMatches = re.Execute(s)
    If mMatches.Count > 0 Then
        For Each mMatch In mMatches
            If prev < mMatch.FirstIndex Then
                str = str + Mid(s, prev, mMatch.FirstIndex + mMatch.Length + 1 - prev) + returnID(questionID)
                prev = nextRightBrace(mMatch.FirstIndex + mMatch.Length + 1, s)
            End If
            DoEvents
        Next
        str = str + Mid(s, prev)
        s = str
        redistributeQuestionID = str
    End If
End Function

Function redistributeF_T_ID(ByRef s As String, ID As String) As Boolean
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    Dim str As String
    Dim F_T_ID As Long
    Dim prev As Long
    
    redistributeF_T_ID = True
    prev = 1
    Set re = New RegExp
    If ID = "FigID" Then
        re.Pattern = "\\FigMinipage\{"
    ElseIf ID = "TabID" Then
        re.Pattern = "\\TabMinipage\{"
    Else
        redistributeF_T_ID = False
        s = "ID类型错误！"
        MsgBox "ID类型错误！"
        Exit Function
    End If
    
    re.Global = True
    Set mMatches = re.Execute(s)
    If mMatches.Count > 0 Then
        For Each mMatch In mMatches
            If prev < mMatch.FirstIndex Then
                If ID = "FigID" Then
                    F_T_ID = returnID(figID)
                ElseIf ID = "TabID" Then
                    F_T_ID = returnID(tabID)
                End If
                str = str + Mid(s, prev, mMatch.FirstIndex + mMatch.Length + 1 - prev)
                prev = nextRightBrace(mMatch.FirstIndex + mMatch.Length + 1, s)
                
                If nextLBrace(prev, s) = False Then
                    MsgBox ID & "有错误！"
                    Exit Function
                Else
                    prev = prev + 1
                End If
                
                str = str + Mid(s, mMatch.FirstIndex + mMatch.Length + 1, _
                                   prev - mMatch.FirstIndex - mMatch.Length - 1) & F_T_ID
                prev = nextRightBrace(prev, s)
            End If
            DoEvents
        Next
        s = str + Mid(s, prev)
    End If
End Function

Function correctUDscript(ByRef str As String)
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim strTemp As String
    Dim stIndex, edIndex As Long
    Dim mutiflag As Boolean
    Set re = New RegExp
    re.Global = True
    '修正双下标
    stIndex = 1
    tempXiuZheng = ""
    re.Pattern = "_\{"
    If re.test(str) Then
        Set mMatches = re.Execute(str)
        If mMatches.Count > 1 Then
            For i = 0 To mMatches.Count - 1
                If stIndex = 1 Or stIndex <= mMatches(i).FirstIndex Then
                    edIndex = mMatches(i).FirstIndex + mMatches(i).Length
                    tempXiuZheng = tempXiuZheng + Mid(str, stIndex, edIndex - stIndex + 1)
                Else
                    GoTo nextDscript
                End If
                mutiflag = True
                Do
                    stIndex = edIndex + 1
                    edIndex = nextRightBrace(stIndex, str)
                    tempXiuZheng = tempXiuZheng + Mid(str, stIndex, edIndex - stIndex)
                    If i + 1 < mMatches.Count Then
                        If mMatches(i + 1).FirstIndex = edIndex Then
                            edIndex = edIndex + mMatches(i + 1).Length
                            tempXiuZheng = tempXiuZheng + " "
                            i = i + 1
                        Else
                            mutiflag = False
                        End If
                    Else
                        mutiflag = False
                    End If
                Loop While mutiflag
                stIndex = edIndex
nextDscript:
            Next
            tempXiuZheng = tempXiuZheng + Mid(str, edIndex)
            str = tempXiuZheng
        End If
    End If
    '修正双上标
    stIndex = 1
    tempXiuZheng = ""
    re.Pattern = "\^\{"
    If re.test(str) Then
        Set mMatches = re.Execute(str)
        If mMatches.Count > 1 Then
            For i = 0 To mMatches.Count - 1
                If stIndex = 1 Or stIndex <= mMatches(i).FirstIndex Then
                    edIndex = mMatches(i).FirstIndex + mMatches(i).Length
                    tempXiuZheng = tempXiuZheng + Mid(str, stIndex, edIndex - stIndex + 1)
                Else
                    GoTo nextUscript
                End If
                mutiflag = True
                Do
                    stIndex = edIndex + 1
                    edIndex = nextRightBrace(stIndex, str)
                    tempXiuZheng = tempXiuZheng + Mid(str, stIndex, edIndex - stIndex)
                    If i + 1 < mMatches.Count Then
                        If mMatches(i + 1).FirstIndex = edIndex Then
                            edIndex = edIndex + mMatches(i + 1).Length
                            tempXiuZheng = tempXiuZheng + " "
                            i = i + 1
                        Else
                            mutiflag = False
                        End If
                    Else
                        mutiflag = False
                    End If
                Loop While mutiflag
                stIndex = edIndex
nextUscript:
            Next
            tempXiuZheng = tempXiuZheng + Mid(str, edIndex)
            str = tempXiuZheng
        End If
    End If
End Function

Function cutXTJ(ByRef doc As String, ByRef finalStr As String, ByVal texFileName As String)
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim strTemp, str As String
    Dim strXZ As String
    Dim strTK As String
    Dim strJD As String
    Dim arr() As String
    
    Set re = New RegExp
    strXZ = ""
    strTK = ""
    strJD = ""
    
    re.Global = True
    re.Pattern = "(\n|\r|" + Chr(13) + ")" + "\S{0,2}(" + questionTypeXZ + "|" + questionTypeTK + "|" + questionTypeJD + ")"

    If re.test(doc) Then
        firstQuestionFlag = True
        Set mMatches = re.Execute(doc)
        strTestName = Mid(doc, 1, mMatches(0).FirstIndex)
        For i = 0 To mMatches.Count - 1
            strTemp = mMatches(i).Value
            If QuestionType(strTemp, questionTypeXZ) = True Then
                '选择题
                If i + 1 < mMatches.Count Then
                    strXZ = Mid(doc, mMatches(i).FirstIndex + 1, mMatches(i + 1).FirstIndex - mMatches(i).FirstIndex)
                Else
                    strXZ = Mid(doc, mMatches(i).FirstIndex + 1)
                End If
                changeToCmdXZ cutApart(strXZ, CStr(strTemp)), finalStr
            ElseIf QuestionType(strTemp, questionTypeTK) = True Then
                '填空题
                If i + 1 < mMatches.Count Then
                    strTK = Mid(doc, mMatches(i).FirstIndex + 1, mMatches(i + 1).FirstIndex - mMatches(i).FirstIndex)
                Else
                    strTK = Mid(doc, mMatches(i).FirstIndex + 1)
                End If
                'correctTK strTK
                changeToCmdTK cutApart(strTK, CStr(strTemp)), finalStr
            ElseIf QuestionType(strTemp, questionTypeJD) = True Then
                '解答题
                If i + 1 < mMatches.Count Then
                    strJD = Mid(doc, mMatches(i).FirstIndex + 1, mMatches(i + 1).FirstIndex - mMatches(i).FirstIndex)
                Else
                    strJD = Mid(doc, mMatches(i).FirstIndex + 1)
                End If
                changeToCmdJD cutApart(strJD, CStr(strTemp)), finalStr
            End If
            DoEvents
        Next i
    Else
        insertDollerT doc
        correctEnvs doc
        correctFig doc, False
        correctTabular doc, False
        correctMathScriptInList doc
        DelDollerInList doc
        doc = delLeftRight(doc)
        correctLeftRight doc
        finalStr = doc
    End If
    re.Pattern = "\\frontPath/"
    finalStr = re.Replace(finalStr, "")
    '$修正
    'adjustDoller finalStr, "." + "(\n|\r|" + Chr(13) + ")."
    re.Pattern = "\$ *\$"
    finalStr = re.Replace(finalStr, "")
    re.Pattern = "(\n|\r|" + Chr(13) + ")" + "(\\item(X|T|J|F))"
    If re.test(finalStr) Then
        finalStr = re.Replace(finalStr, Chr(13) + Chr(13) + "$2")
    End If
    re.Pattern = "(\n|\r|" + Chr(13) + ")"
    strTestName = re.Replace(strTestName, "")
    finalStr = "\section{" + strTestName + "}" + finalStr + Chr(13) + "\end{myitemize}"
    If readReplaceList = True Then
        replaceList finalStr
    End If
End Function

Function cutApart(ByRef strQuestionAndAnswer As String, strType As String) As String()
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim strTemp As String
    Set re = New RegExp
    re.Global = True
    re.Pattern = "([\n|\r|" + Chr(13) + "])+(\d+．)"
    If re.test(strQuestionAndAnswer) = False Then MsgBox strType + "未发现题目！" + Chr(13) + "([\n|\r| + Chr(13) + ])+\d+．分割无效。"
    If strType = "" Then
        cutApart = Split(re.Replace(strQuestionAndAnswer, Chr(0) + "$1$2"), Chr(0))
    Else
        cutApart = Split(re.Replace(strQuestionAndAnswer, Chr(0)), Chr(0))
    End If
End Function

Function CutAnswer(ByVal strAnswer As String) As String
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim strTemp As String
    Set re = New RegExp
    re.Global = True
'Public AnswerSolutionBoundary As String
'Public SolutionStart As String
'Public SolutionEnd As String
    re.Pattern = "[\n\r" + Chr(13) + "]" + answerSolutionBoundary
    Set mMatches = re.Execute(strAnswer)
    If mMatches.Count = 1 Then
        strTemp = Left(strAnswer, mMatches(0).FirstIndex)
    End If
    re.Pattern = "[\n\r" + Chr(13) + "]" + solutionStart
    Set mMatches = re.Execute(strAnswer)
    If mMatches.Count = 1 Then
        strAnswer = Mid(strAnswer, mMatches(0).FirstIndex + mMatches(0).Length + 1)
    Else
        strTemp = ""
    End If
    re.Pattern = "[\n\r" + Chr(13) + "]" + solutionEnd
    Set mMatches = re.Execute(strAnswer)
    If mMatches.Count = 1 Then
        strAnswer = Left(strAnswer, mMatches(0).FirstIndex)
    End If
    CutAnswer = strTemp + strAnswer
End Function
Function changeToCmdXZ(ByRef strQuestionAndAnswer() As String, ByRef finalStr As String)
    Dim strQuestion() As String
    Dim strCTemp(4) As String
    Dim strAnswer As String
    Dim strTemp As String
    Dim str As String
    Dim item As String
    Dim flag As Boolean 'true 普通选择题 false为图像选择题
    Dim re As Object
    Dim mMatches As Object, cMatches As Object     '匹配字符串集合对象
    
    Set re = New RegExp
    re.Global = True
    '下标可能越界
    If firstQuestionFlag = True Then
        finalStr = finalStr + strQuestionAndAnswer(0) + Chr(13) + "\begin{myitemize}" + Chr(13)
        firstQuestionFlag = False
    Else
        finalStr = finalStr + Chr(13) + "\end{myitemize}" + strQuestionAndAnswer(0) + Chr(13) + "\begin{myitemize}" + Chr(13)
    End If
    For i = 1 To UBound(strQuestionAndAnswer)
        flag = False
        str = strQuestionAndAnswer(i)
        re.Pattern = questionAnswerBoundary '"【解答】"
        If re.test(str) = True Then
            Set mMatches = re.Execute(str)
            strTemp = Mid(str, 1, mMatches(0).FirstIndex)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            strAnswer = CutAnswer(Mid(str, mMatches(0).FirstIndex + mMatches(0).Length + 1))
            
            
            strQuestion = splitXX(strTemp)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''TrimEnter strQuestion(UBound(strQuestion))
            re.Pattern = "\\includegraphics\[width=\\lFigWidth\]\{([\w-/])+/image([0-9])+\.png\}"
            For j = 1 To UBound(strQuestion)
                strCTemp(j) = TrimEnter(strQuestion(j))
                Set mMatches = re.Execute(strCTemp(j))
                If mMatches.Count > 0 Then
                    If mMatches(0).Length <> Len(strCTemp(j)) Then
                        flag = True
                        Exit For
                    End If
                Else
                    flag = True
                    Exit For
                End If
            Next j
            If flag = True Then
                item = "\itemX{"
                For j = 0 To UBound(strQuestion)
                    insertDollerT strQuestion(j)
                    '修正
                    correct strQuestion(j)
                    strQuestion(j) = TrimEnter(strQuestion(j))
                Next j
            Else
                item = "\itemF{"
                insertDollerT strQuestion(0)
                '修正
                correct strQuestion(0)
                strQuestion(0) = TrimEnter(strQuestion(0))
                For j = 1 To UBound(strQuestion)
                    strQuestion(j) = strCTemp(j)
                Next j
            End If
            insertDollerT strAnswer
            re.Pattern = "故选"
            '修正
            correct strAnswer
            strAnswer = re.Replace(strAnswer, Chr(13) + "\hh\color{blue}故选")
            lastReplace strAnswer
            finalStr = finalStr + Chr(13) + item + returnID(questionID) + "}{" + strQuestion(0) + "\xz }"
            For k = 1 To UBound(strQuestion)
                finalStr = finalStr + Chr(13) + "{" + strQuestion(k) + "}"
            Next k
            For k = k To 4
                finalStr = finalStr + Chr(13) + "{\color{red}选项未识别}"
            Next k
            strAnswer = TrimEnter(strAnswer)
            finalStr = finalStr + Chr(13) + "{" + strAnswer + "}"
        Else
            strUnidentified = strUnidentified + str + Chr(13)
            Unidentified = Unidentified + 1
        End If
        DoEvents
    Next i
End Function

Function splitXX(ByVal str As String) As String()
    
    Dim strTemp As String
    Dim strTG As String
    Dim strXX As String
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Set re = New RegExp
    re.Global = True
    re.Pattern = "[\n|\r|" + Chr(13) + "]" + "A[．\.]"
    Set mMatches = re.Execute(str)
    If mMatches.Count = 1 Then
        strTG = Left(str, mMatches(0).FirstIndex)
        strXX = Mid(str, mMatches(0).FirstIndex + 1)
        re.Pattern = "[A-D][\.．]"
        Set mMatches = re.Execute(strXX)
        If mMatches.Count = 4 Then
            strXX = re.Replace(strXX, Chr(0))
            strTemp = strTG + strXX
        Else
            strTemp = str
            MsgBox "选项未能分割为4项！"
        End If
    Else
        strTemp = str
        MsgBox "未找到A选项分界！" + Chr(13) + "[\n|\r| + Chr(13) + ]" + "A[．\.]"
    End If
    splitXX = Split(strTemp, Chr(0))
End Function

Function lastReplace(ByRef str As String)
    Dim re As Object
    Set re = New RegExp
    re.Global = True
    re.Pattern = "\\\\\\FigMinipage"
    str = re.Replace(str, "\hh\FigMinipage")
    re.Pattern = "\\\\\\TabMinipage"
    str = re.Replace(str, "\hh\TabMinipage")
End Function

Function changeToCmdTK(ByRef strQuestionAndAnswer() As String, ByRef finalStr As String)
    Dim strQuestion As String
    Dim strAnswer As String
    Dim strTemp As String
    Dim str As String
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Set re = New RegExp
    re.Global = True
    '下标可能越界
    If firstQuestionFlag = True Then
        finalStr = finalStr + strQuestionAndAnswer(0) + Chr(13) + "\begin{myitemize}" + Chr(13)
        firstQuestionFlag = False
    Else
        finalStr = finalStr + Chr(13) + "\end{myitemize}" + strQuestionAndAnswer(0) + Chr(13) + "\begin{myitemize}" + Chr(13)
    End If

    For i = 1 To UBound(strQuestionAndAnswer)
        str = strQuestionAndAnswer(i)
        re.Pattern = questionAnswerBoundary '"【解答】"
        If re.test(str) = True Then
            Set mMatches = re.Execute(str)
            strQuestion = Mid(str, 1, mMatches(0).FirstIndex)
            strAnswer = CutAnswer(Mid(str, mMatches(0).FirstIndex + mMatches(0).Length + 1))
            
            insertDollerT strQuestion
            '修正
            correct strQuestion
            strQuestion = TrimEnter(strQuestion)
            insertDollerT strAnswer
            '修正
            correct strAnswer
            lastReplace strAnswer
            re.Pattern = "故答案"
            If re.test(strAnswer) Then
                strAnswer = re.Replace(strAnswer, Chr(13) + "\hh\color{blue}故答案")
            End If
            strAnswer = TrimEnter(strAnswer)
            finalStr = finalStr + Chr(13) + "\itemT{" + returnID(questionID) + "}{" + strQuestion + "}" + Chr(13) + "{" + strAnswer + "}"
        Else
            strUnidentified = strUnidentified + str + Chr(13)
            Unidentified = Unidentified + 1
        End If
        DoEvents
    Next
End Function

Function changeToCmdJD(ByRef strQuestionAndAnswer() As String, ByRef finalStr As String)
    Dim strQuestion As String
    Dim strAnswer As String
    Dim strTemp As String
    Dim str As String
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Set re = New RegExp
    re.Global = True
    '下标可能越界
    If firstQuestionFlag = True Then
        finalStr = finalStr + strQuestionAndAnswer(0) + Chr(13) + "\begin{myitemize}" + Chr(13)
        firstQuestionFlag = False
    Else
        finalStr = finalStr + Chr(13) + "\end{myitemize}" + strQuestionAndAnswer(0) + Chr(13) + "\begin{myitemize}" + Chr(13)
    End If

    For i = 1 To UBound(strQuestionAndAnswer)
        str = strQuestionAndAnswer(i)
        re.Pattern = questionAnswerBoundary '"【解答】"
        If re.test(str) = True Then
            Set mMatches = re.Execute(str)
            strQuestion = Mid(str, 1, mMatches(0).FirstIndex)
            strAnswer = CutAnswer(Mid(str, mMatches(0).FirstIndex + mMatches(0).Length + 1))
        
            questionList strQuestion
            
            re.Pattern = "(\n|\r|" + Chr(13) + ")" + "解?：?" + "(\(2\))"
            If re.test(strAnswer) Then
                strAnswer = re.Replace(strAnswer, Chr(13) + "\hh\color{two}$2")
            End If
            re.Pattern = "(\n|\r|" + Chr(13) + ")" + "解?：?" + "(\(3\))"
            If re.test(strAnswer) Then
                strAnswer = re.Replace(strAnswer, Chr(13) + "\hh\color{blue}$2")
            End If
            insertDollerT strAnswer
            '修正
            correct strAnswer
            lastReplace strAnswer
            strAnswer = TrimEnter(strAnswer)
            finalStr = finalStr + Chr(13) + "\itemJ{" + returnID(questionID) + "}{" + strQuestion + "}" + Chr(13) + "{" + strAnswer + "}"
        Else
            strUnidentified = strUnidentified + str + Chr(13)
            Unidentified = Unidentified + 1
        End If
        DoEvents
    Next
End Function

Function questionList(ByRef strQuestion As String)
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim str() As String
    Set re = New RegExp
    re.Global = True
    re.Pattern = "(\n|\r|" + Chr(13) + ")" + "\([1-9]\)"
    If re.test(strQuestion) = True Then
        strQuestion = re.Replace(strQuestion, Chr(0))
        str = Split(strQuestion, Chr(0))
        For i = 0 To UBound(str)
            insertDollerT str(i)
            correct str(i)
            str(i) = TrimEnter(str(i))
        Next
        strQuestion = str(0) + Chr(13) + "\begin{questionList}"    'questionList
        For i = 1 To UBound(str)
            strQuestion = strQuestion + Chr(13) + "\itemQ " + str(i)
        Next
        strQuestion = strQuestion + Chr(13) + "\end{questionList}" 'questionList
    Else
        insertDollerT strQuestion
        correct strQuestion
        strQuestion = TrimEnter(strQuestion)
    End If
End Function


Function insertDollerT(ByRef str As String)
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim eMatches As Object
    Dim strTemp As String
    Dim prev As Long
    
    prev = 1
    Set re = New RegExp
    re.Global = True
    re.Pattern = "([a-zA-Z0-9\^\\_\*\+-=<>\(\)\[\]\{\}\|/ %" + Chr(39) + "])+" '\n\r加数学环境
    
    Set mMatches = re.Execute(str)
    If mMatches.Count > 0 Then
        For Each mMatch In mMatches
            strTemp = strTemp + Mid(str, prev, mMatch.FirstIndex + 1 - prev)
            
            strTemp = strTemp & "$" & Mid(str, mMatch.FirstIndex + 1, mMatch.Length) & "$" 'addDollor(Mid(str, mMatch.FirstIndex + 1, mMatch.Length))
            prev = mMatch.FirstIndex + mMatch.Length + 1
            DoEvents
        Next
        strTemp = strTemp + Mid(str, prev)
        're.Pattern = "\$(\n|\r)+\$"
        'strTemp = re.Replace(strTemp, "")
        str = strTemp
    End If
End Function
'加$  已废弃
Function addDollor(ByVal str As String) As String
    Dim re As Object
    Dim eMatches As Object
    Dim strTemp As String
    Dim seFlag As Boolean, eeFlag As Boolean
    
    seFlag = False
    eeFlag = False
    prev = 1
    Set re = New RegExp
    re.Global = True
    re.Pattern = "(\n|\r|" + Chr(13) + ")+" '加数学环境
    
    'str = Chr(13) + "math" + Chr(13) + Chr(13)
    str = Trim(str)
    Set eMatches = re.Execute(str)
    If eMatches.Count = 0 Then
        '无回车
    ElseIf eMatches.Count = 1 Then
        If eMatches(0).Length <> Len(str) Then
            If eMatches(0).FirstIndex = 0 Then
                '回车开头
                seFlag = True
            ElseIf Len(str) = eMatches(0).FirstIndex + eMatches(0).Length Then
                '回车结尾
                eeFlag = True
            End If
        Else
            '只有回车
            Exit Function
        End If
    Else
        '多个回车
        If eMatches(0).FirstIndex = 0 Then
            '回车开头
            seFlag = True
        End If
        If Len(str) = eMatches(eMatches.Count - 1).FirstIndex + eMatches(eMatches.Count - 1).Length Then
            '回车结尾
            eeFlag = True
        End If
    End If
    If seFlag And eeFlag Then
        '头尾都有回车
        strTemp = Chr(13) + "$" + Mid(str, eMatches(0).Length + 1, eMatches(eMatches.Count - 1).FirstIndex - eMatches(0).Length) + "$" + Chr(13)
    ElseIf seFlag Then
        '头回车
        strTemp = Chr(13) + "$" + Right(str, Len(str) - eMatches(0).Length) + "$"
    ElseIf eeFlag Then
        '尾回车
        strTemp = "$" + Left(str, eMatches(eMatches.Count - 1).FirstIndex) + "$" + Chr(13)
    Else
        '头尾无回车
        strTemp = "$" + str + "$"
    End If
    addDollor = strTemp
End Function

Function TrimEnterA(ByVal str As String, Optional LR As String = "") As String
    Dim re As Object
    Dim eMatches As Object
    Dim strTemp As String
    Dim seFlag As Boolean, eeFlag As Boolean
    
    seFlag = False
    eeFlag = False
    prev = 1
    Set re = New RegExp
    re.Global = True
    re.Pattern = "(\n|\r|" + Chr(13) + "| )+" '加数学环境
    
    'str = Chr(13) + "math" + Chr(13) + Chr(13)
    Set eMatches = re.Execute(str)
    If eMatches.Count = 0 Then
        '无回车
        TrimEnterA = str
        Exit Function
    ElseIf eMatches.Count = 1 Then
        If eMatches(0).Length <> Len(str) Then
            If eMatches(0).FirstIndex = 0 Then
                '回车开头
                seFlag = True
            ElseIf Len(str) = eMatches(0).FirstIndex + eMatches(0).Length Then
                '回车结尾
                eeFlag = True
            End If
        Else
            '只有回车
            Exit Function
        End If
    Else
        '多个回车
        If eMatches(0).FirstIndex = 0 Then
            '回车开头
            seFlag = True
        End If
        If Len(str) = eMatches(eMatches.Count - 1).FirstIndex + eMatches(eMatches.Count - 1).Length Then
            '回车结尾
            eeFlag = True
        End If
    End If
    
    If seFlag And eeFlag Then
        '头尾都有回车
        If LR = "" Then
            strTemp = Mid(str, eMatches(0).Length + 1, eMatches(eMatches.Count - 1).FirstIndex - eMatches(0).Length)
        ElseIf UCase(LR) = "L" Then
            strTemp = Right(str, Len(str) - eMatches(0).Length)
        ElseIf UCase(LR) = "R" Then
            strTemp = Left(str, eMatches(eMatches.Count - 1).FirstIndex)
        End If
    ElseIf seFlag Then
        '头回车
        If LR = "" Or UCase(LR) = "L" Then
            strTemp = Right(str, Len(str) - eMatches(0).Length)
        Else
            strTemp = str
        End If
    ElseIf eeFlag Then
        '尾回车
        If LR = "" Or UCase(LR) = "R" Then
            strTemp = Left(str, eMatches(eMatches.Count - 1).FirstIndex)
        Else
            strTemp = str
        End If
    Else
        strTemp = str
    End If
    TrimEnterA = strTemp
End Function

Function TrimEnter(ByVal str As String, Optional LR As String = "") As String
    Dim c As String
    Dim flag As Boolean
    Dim startStr As Long, endStr As Long
    Dim coordinate As Long, lenStr As Long
    
    lenStr = Len(str)
    coordinate = 1
    flag = True
    
    Do
        If coordinate > lenStr Then
            Exit Function
        End If
            
        c = Mid(str, coordinate, 1)
        If c = " " Or c = Chr(13) Then
            coordinate = coordinate + 1
        Else
            startStr = coordinate
            flag = False
        End If
        'DoEvents
    Loop While flag
    flag = True
    coordinate = lenStr
    Do
        c = Mid(str, coordinate, 1)
        If c = " " Or c = Chr(13) Then
            coordinate = coordinate - 1
        Else
            endStr = coordinate
            flag = False
        End If
        'DoEvents
    Loop While flag
    If LR = "" Then
        TrimEnter = Mid(str, startStr, endStr - startStr + 1)
    ElseIf UCase(LR) = "R" Then
        TrimEnter = Mid(str, 1, endStr)
    ElseIf UCase(LR) = "L" Then
        TrimEnter = Mid(str, startStr)
    Else
        MsgBox "TrimEnter参数 LR 错误!"
    End If
End Function

Function correct(ByRef str As String)
    correctAlign str
    correctArray str
    'correctCases str
    correctEnvs str
    correctFig str
    correctTabular str
    correctMathScriptInList str
    DelDollerInList str
    correctDfracInScript str
    str = delLeftRight(str)
    correctLeftRight str
    'delDoller str, ".\\wendu."
End Function

Function returnID(ByRef ID As Long) As String
    ID = ID + 1
    returnID = CStr(ID)
End Function

Function correctFig(ByRef str As String, Optional Qflag As Boolean = True)
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    Dim prev As Long
    Dim strTemp, tempXiuZheng As String
    Dim includeTemp As String
    
    Set re = New RegExp
    re.Global = True
    prev = 1
    
    
    re.Pattern = ".\\includegraphics\[width=\\lFigWidth\]\{([\w-/])+/image([0-9])+\.png\}."    '

    're.Pattern = ".\\includegraphics\[width=\\lFigWidth\]\{\\frontPath([\w-/])+\.png\}."    '
    Set mMatches = re.Execute(str)
    If mMatches.Count > 0 Then
        For Each mMatch In mMatches
            strTemp = mMatch.Value
            tempXiuZheng = tempXiuZheng + Mid(str, prev, mMatch.FirstIndex + 1 - prev)
            If Qflag = True Then
                includeTemp = "\\\FigMinipage{%" + Chr(13) + Mid(strTemp, 2, mMatch.Length - 2) + "%" + Chr(13) + "}{" + returnID(figID) + "}"
            Else
                includeTemp = Chr(13) + Mid(strTemp, 2, mMatch.Length - 2) + Chr(13)
            End If
            re.Pattern = "\$"
            If re.test(includeTemp) Then
                includeTemp = re.Replace(includeTemp, "")
            End If
            If Left(strTemp, 1) = "$" And Right(strTemp, 1) = "$" Then
                tempXiuZheng = tempXiuZheng + includeTemp
            ElseIf Left(strTemp, 1) = "$" Then
                tempXiuZheng = tempXiuZheng + includeTemp + "$" + Right(strTemp, 1)
            ElseIf Right(strTemp, 1) = "$" Then
                tempXiuZheng = tempXiuZheng + "$" + Left(strTemp, 1) + includeTemp
            Else
                tempXiuZheng = tempXiuZheng + "$" + Left(strTemp, 1) + includeTemp + "$" + Right(strTemp, 1)
            End If
            prev = mMatch.FirstIndex + mMatch.Length + 1
            DoEvents
        Next
        tempXiuZheng = tempXiuZheng + Mid(str, prev)
        str = tempXiuZheng
    End If
End Function

Function correctArray(ByRef str As String)
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    Dim prev As Long
    Dim strTemp As String
    Dim tempXiuZheng As String
    Dim n As Long
    Dim rightCount As Long
    
    tempXiuZheng = ""
    Set re = New RegExp
    re.Global = True
    prev = 1
    
    re.Pattern = "\\left\\\{\\begin\{array\}\{[clr]\}"
    Set mMatches = re.Execute(str)
    re.Pattern = "\\end\{array\}\\right\."
    re.Global = False
    If mMatches.Count > 0 Then
        For Each mMatch In mMatches
            If prev < mMatch.FirstIndex Then
                strTemp = Mid(str, prev, mMatch.FirstIndex + 1 - prev)
                If prev > 1 Then
                    'strTemp = re.Replace(strTemp, "\end{cases}")
                    If matchendarray(strTemp, rightCount) = False Then
                        Exit Function
                    End If
                End If
                tempXiuZheng = tempXiuZheng + strTemp + "\begin{cases}"
                prev = mMatch.FirstIndex + mMatch.Length + 1
            End If
            DoEvents
        Next
        strTemp = Mid(str, prev)
        'strTemp = re.Replace(strTemp, "\end{cases}")
                    If matchendarray(strTemp, rightCount) = False Then
                        Exit Function
                    End If
        str = tempXiuZheng + strTemp
    End If
End Function

Function matchendarray(ByRef str As String, ByRef rightCount As Long) As Boolean
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    Dim prev As Long
    Dim flag As Long
    Dim n As Long
    Dim strTemp As String, tempXiuZheng As String
    
    tempXiuZheng = ""
    flag = 0
    Set re = New RegExp
    re.Global = True
    prev = 1
    
    re.Pattern = "(\\begin\{array\})|(\\end\{array\})"
    Set mMatches = re.Execute(str)
    If mMatches.Count > 0 Then
        For n = 0 To mMatches.Count - 1
            If CStr(mMatches(n).Value) = "\begin{array}" Then
                flag = flag + 1
            ElseIf CStr(mMatches(n).Value) = "\end{array}" Then
                flag = flag - 1
            End If
            If flag < 0 Then Exit For
        Next
        If flag < 0 Then
            strTemp = Left(str, mMatches(n).FirstIndex)
            strTemp = strTemp + "\end{cases}" + Mid(str, mMatches(n).FirstIndex + mMatches(n).Length + 1)
        Else
            matchendarray = False
            Exit Function
        End If
    Else
        matchendarray = False
        Exit Function
    End If
    
    re.Pattern = "(\\left(\.|\(|\[|\\\{|\||\\\||\\[a-zA-Z]+|/|\)|\]|\\\}))|(\\right(\.|\)|\]|\\\}|\||\\\||\\[a-zA-Z]+|\(|\[|\\\{|/))"
    Set mMatches = re.Execute(strTemp)
    flag = 0
    
    If mMatches.Count > 0 Then
        For n = 0 To mMatches.Count - 1
            If Left(CStr(mMatches(n).Value), 2) = "\l" Then
                flag = flag + 1
            ElseIf Left(CStr(mMatches(n).Value), 2) = "\r" Then
                flag = flag - 1
            End If
            If flag < 0 Then Exit For
        Next
        If flag < 0 Then
            tempXiuZheng = Left(strTemp, mMatches(n).FirstIndex)
            tempXiuZheng = tempXiuZheng + Mid(strTemp, mMatches(n).FirstIndex + mMatches(n).Length + 1)
        Else
            rightCount = rightCount + flag + 1
            matchendarray = False
            Exit Function
        End If
    Else
        matchendarray = False
        Exit Function
    End If
    str = tempXiuZheng
    matchendarray = True
End Function

Function correctEnvs(ByRef str As String)
    If UBound(correctEnvironments) > 0 Then
        For Each env In correctEnvironments
            correctEnv str, CStr(env)
        Next
    End If
End Function

Function correctEnv(ByRef str As String, env As String)
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    Dim prev As Long
    Dim strTemp, tempXiuZheng As String
    Dim n As Long
    
    tempXiuZheng = ""
    Set re = New RegExp
    re.Global = True
    prev = 1
    
    re.Pattern = "(\\begin\{" + env + "\})|(\\end\{" + env + "\})"
    Set mMatches = re.Execute(str)
    If mMatches.Count > 0 Then
        For n = 0 To mMatches.Count - 1 Step 2
            tempXiuZheng = tempXiuZheng + Mid(str, prev, mMatches(n).FirstIndex - prev + 1)
            strTemp = Mid(str, mMatches(n).FirstIndex + mMatches(n).Length + 1, mMatches(n + 1).FirstIndex - mMatches(n).FirstIndex - mMatches(n).Length)
            
            re.Pattern = "\$"
            If re.test(strTemp) Then
                strTemp = re.Replace(strTemp, "")
            End If
            
            re.Pattern = "，"
            If re.test(strTemp) Then
                strTemp = re.Replace(strTemp, ",")
            End If
            tempXiuZheng = tempXiuZheng + "\begin{" + env + "}" + strTemp + "\end{" + env + "}"
            prev = mMatches(n + 1).FirstIndex + mMatches(n + 1).Length + 1
            DoEvents
        Next n
        str = tempXiuZheng + Mid(str, prev)
    End If

End Function

Function correctTabular(ByRef str As String, Optional Qflag As Boolean = True)
    Dim larr, lbrr As Long
    Dim gapeFlag As Boolean
    Dim arr() As String
    Dim brr() As String
    
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    Dim prev As Long
    Dim strTemp, tempXiuZheng As String
    Dim n As Long
    Dim beginTabular, endTabular As String
    
    tempXiuZheng = ""
    Set re = New RegExp
    re.Global = True
    prev = 1
    
    re.Pattern = "(.(\n|\r){0,}\\begin\{tabular\})|(\\end\{tabular\}(\n|\r){0,}.)"
    Set mMatches = re.Execute(str)
    If mMatches.Count > 0 Then
        For n = 0 To mMatches.Count - 1 Step 2
            tempXiuZheng = tempXiuZheng + Mid(str, prev, mMatches(n).FirstIndex - prev + 1)
            strTemp = Mid(str, mMatches(n).FirstIndex + mMatches(n).Length + 1, mMatches(n + 1).FirstIndex - mMatches(n).FirstIndex - mMatches(n).Length)
            re.Pattern = "\$|\n|\r|" + Chr(13)
            tabularTemp = re.Replace(strTemp, "")
            re.Pattern = "\\par"
            tabularTemp = re.Replace(tabularTemp, "")
            
            lc = Left(CStr(mMatches(n).Value), 1)
            rc = Right(CStr(mMatches(n + 1).Value), 1)
            
            beginTabular = mMatches(n).Value
            endTabular = mMatches(n + 1).Value
            
            beginTabular = Right(beginTabular, 15)
            endTabular = Left(endTabular, 13)
            
            
            arr = Split(tabularTemp, "\hline")
            larr = UBound(arr)
            For i = 1 To larr
                If InStr(1, arr(i), "\dfrac") > 0 Then
                    gapeFlag = True
                Else
                    gapeFlag = False
                End If
                brr = Split(arr(i), "&")
                lbrr = UBound(brr)
                For j = 0 To lbrr
                    brr(j) = Trim(brr(j))
                    If InStr(str, "\multi") = 0 Then
                        insertDollerT brr(j)
                    End If
                    If gapeFlag = True Then
                        If InStr(1, brr(j), "\dfrac") > 0 Then
                            brr(j) = "\gape{" + brr(j) + "}"
                            gapeFlag = False
                        End If
                    End If
                Next j
                arr(i) = Join(brr, " & ")
                DoEvents
            Next i
            tabularTemp = Join(arr, Chr(13) + "\hline" + Chr(13))
            re.Pattern = "\\\\\$"
            tabularTemp = re.Replace(tabularTemp, "$\\")
            re.Pattern = "\$ +\$"
            tabularTemp = re.Replace(tabularTemp, "")
            
            If Qflag = True Then
                tabularTemp = "\\\TabMinipage{" + beginTabular + tabularTemp + endTabular + "}{" + returnID(tabID) + "}\\" + Chr(13)
            Else
                tabularTemp = Chr(13) + beginTabular + tabularTemp + endTabular + Chr(13)
            End If
            If lc = "$" And rc = "$" Then
                tempXiuZheng = tempXiuZheng + tabularTemp
            ElseIf lc = "$" Then
                tempXiuZheng = tempXiuZheng + tabularTemp + "$" + rc
            ElseIf rc = "$" Then
                tempXiuZheng = tempXiuZheng + "$" + lc + tabularTemp
            Else
                tempXiuZheng = tempXiuZheng + lc + "$" + tabularTemp + "$" + rc
            End If
            prev = mMatches(n + 1).FirstIndex + mMatches(n + 1).Length + 1
            DoEvents
        Next n
        str = tempXiuZheng + Mid(str, prev)
    End If
End Function

Function correctMathScriptInList(ByRef str As String)
    Dim strPram() As String
    If UBound(mathScriptList) >= 0 Then
        For Each env In mathScriptList
            strPram = Split(CStr(env), ";")
            If UBound(strPram) = 0 Then
                correctMathScript str, strPram(0)
            ElseIf UBound(strPram) = 1 Then
                correctMathScript str, strPram(0), CInt(strPram(1))
            End If
        Next
    End If
End Function

Function correctMathScript(ByRef s As String, strPattern As String, Optional braceNum As Integer = 1) As String
    Dim re As Object
    Dim mMatches, mMatchesD As Object     '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    Dim str As String, strTemp As String, strMathScript As String
    Dim prev As Long
    Dim prevD As Long
    Dim i, lFlag As Integer
    
    prev = 1
    str = ""
    strTemp = ""
    
    Set re = New RegExp
    re.Global = True
    re.Pattern = strPattern
    
    Set mMatches = re.Execute(s)
    If mMatches.Count > 0 Then
        For Each mMatch In mMatches
            If mMatch.FirstIndex < prev Then
                GoTo nextmMatch
            Else
                str = str + Mid(s, prev, mMatch.FirstIndex + 1 - prev)
            End If
            prev = nextRightBrace(mMatch.FirstIndex + mMatch.Length + 1, s, braceNum)
            strMathScript = Mid(s, mMatch.FirstIndex + 1, prev - mMatch.FirstIndex) '截取括号内代码
            
            re.Pattern = "\$"
            If re.test(strMathScript) Then
                tempXiuZheng = ""
                prevD = 1
                Set mMatchesD = re.Execute(strMathScript)
                For n = 0 To mMatchesD.Count - 1 Step 2
                    tempXiuZheng = tempXiuZheng + Mid(strMathScript, prevD, mMatchesD(n).FirstIndex - prevD + 1)
                    strTemp = "\text{" + Mid(strMathScript, mMatchesD(n).FirstIndex + mMatchesD(n).Length + 1, _
                                                         mMatchesD(n + 1).FirstIndex - mMatchesD(n).FirstIndex - mMatchesD(n).Length) + "}"
                    tempXiuZheng = tempXiuZheng + strTemp
                    prevD = mMatchesD(n + 1).FirstIndex + mMatchesD(n + 1).Length + 1
                Next n
                tempXiuZheng = tempXiuZheng + Mid(strMathScript, prevD)
            Else
                tempXiuZheng = strMathScript
            End If
            prev = prev + 1
            str = str + tempXiuZheng
nextmMatch:
            DoEvents
        Next
        s = str + Mid(s, prev)
    End If
End Function

Function correctDfracInScript(ByRef s As String)
    Dim re As Object
    Dim mMatches, mMatchesD As Object     '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    Dim str As String, strMathScript As String
    Dim prev As Long
    
    prev = 1
    str = ""
    
    Set re = New RegExp
    re.Global = True
    re.Pattern = "[\^_]\{"
    
    Set mMatches = re.Execute(s)
    If mMatches.Count > 0 Then
        For Each mMatch In mMatches
            If mMatch.FirstIndex < prev Then
                GoTo nextmMatch
            Else
                str = str + Mid(s, prev, mMatch.FirstIndex + 1 - prev)
            End If
            prev = nextRightBrace(mMatch.FirstIndex + mMatch.Length + 1, s)
            strMathScript = Mid(s, mMatch.FirstIndex + 1, prev - mMatch.FirstIndex) '截取括号内代码
            strMathScript = Replace(strMathScript, "\dfrac{", "\frac{")
            str = str + strMathScript
            prev = prev + 1
nextmMatch:
            DoEvents
        Next
        s = str + Mid(s, prev)
    End If
End Function

Public Function nextLBrace(ByRef coordinate As Long, ByVal str As String) As Boolean
    Dim c As String
    Dim l As Long
    Dim index As Long
    index = coordinate
    l = Len(str)
    Do
        c = Mid(str, index, 1)
        If c = "{" Then
            nextLBrace = True
            coordinate = index
            Exit Function
        ElseIf index <= l Then
            index = index + 1
        Else
            nextLBrace = False
            Exit Function
        End If
    Loop While True
End Function

Function DelDollerInList(ByRef str As String)
    If UBound(delDollerList) >= 0 Then
        For Each env In delDollerList
            delDoller str, "." & CStr(env) & "."
        Next
    End If
End Function

Function delDoller(ByRef str As String, strPattern As String)
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    Dim prev As Long
    Dim strTemp, tempXiuZheng As String
    
    Set re = New RegExp
    re.Global = True
    prev = 1
    
    re.Pattern = strPattern

    Set mMatches = re.Execute(str)
    If mMatches.Count > 0 Then
        For Each mMatch In mMatches
            strTemp = mMatch.Value
            tempXiuZheng = tempXiuZheng + Mid(str, prev, mMatch.FirstIndex - prev + 1)
            If Left(strTemp, 1) = "$" And Right(strTemp, 1) = "$" Then
                tempXiuZheng = tempXiuZheng + Mid(strTemp, 2, mMatch.Length - 2)
            ElseIf Left(strTemp, 1) = "$" Then
                tempXiuZheng = tempXiuZheng + Mid(strTemp, 2, mMatch.Length - 2) + "$" + Right(strTemp, 1)
            ElseIf Right(strTemp, 1) = "$" Then
                tempXiuZheng = tempXiuZheng + Left(strTemp, 1) + "$" + Mid(strTemp, 2, mMatch.Length - 2)
            Else
                tempXiuZheng = tempXiuZheng + Left(strTemp, 1) + "$" + Mid(strTemp, 2, mMatch.Length - 2) + "$" + Right(strTemp, 1)
            End If
            prev = mMatch.FirstIndex + mMatch.Length + 1
        Next
        tempXiuZheng = tempXiuZheng + Mid(str, prev)
        str = tempXiuZheng
    End If
End Function
'调整$位置， 已废弃
Function adjustDoller(ByRef str As String, strPattern As String)
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    Dim prev As Long
    Dim strTemp, tempXiuZheng As String
    
    Set re = New RegExp
    re.Global = True
    prev = 1
    
    re.Pattern = "(\n|\r|" + Chr(13) + ")" + "{2,}"
    str = re.Replace(str, Chr(13))
    
    re.Pattern = strPattern

    re.Pattern = "\$ +\$"
    str = re.Replace(str, "")
    re.Pattern = "\$\$"
    str = re.Replace(str, "")
    're.Pattern = "\$" + Chr(13) + "\$"
    'str = re.Replace(str, Chr(13))
    re.Pattern = Chr(13) + "\$\\\\\\FigMinipage"
    str = re.Replace(str, "$" + Chr(13) + "\\\FigMinipage")
    re.Pattern = Chr(13) + "\$\\item"
    str = re.Replace(str, "$" + Chr(13) + "\item")
    re.Pattern = "\\hline\$" + "(\n|\r|" + Chr(13) + ")"
    str = re.Replace(str, "\hline" + Chr(13) + "$")
    re.Pattern = "\\end\{cases\}" + "(\n|\r|" + Chr(13) + ") *\$"
    str = re.Replace(str, "\end{cases}$" + Chr(13))
End Function

Function nextLeftBrace(ByVal coordinate As Long, ByVal str As String, Optional num As Integer = 1) As Long
    Dim c As String
    Dim flag As Boolean
    Dim stack As Integer
    Dim index, lenStr As Long
    
    lenStr = Len(str)
    index = coordinate
    stack = 0
    flag = True
    
    Do
        If coordinate > lenStr Then
            nextLeftBrace = 0
            Exit Function
        End If
            
        c = Mid(str, coordinate, 1)
         If c = "}" Then
            stack = stack + 1
            coordinate = coordinate + 1
        ElseIf c = "{" Then
            stack = stack - 1
            If stack < 0 Then
                If num = 1 Then
                    flag = False
                Else
                    num = num - 1
                    coordinate = coordinate + 1
                End If
            Else
                coordinate = coordinate + 1
            End If
        ElseIf c = "\" Then
            If Mid(str, coordinate + 1, 4) = "left" Then
                coordinate = coordinate + 5
                If Mid(str, coordinate, 1) = "\" Then
                    coordinate = coordinate + 2
                Else
                    coordinate = coordinate + 1
                End If
            ElseIf Mid(str, coordinate + 1, 5) = "right" Then
                coordinate = coordinate + 6
                If Mid(str, coordinate, 1) = "\" Then
                    coordinate = coordinate + 2
                Else
                    coordinate = coordinate + 1
                End If
            Else
                coordinate = coordinate + 2
            End If
        Else
            coordinate = coordinate + 1
        End If
        DoEvents
    Loop While flag

    nextLeftBrace = coordinate
End Function

Function nextRightBrace(ByVal coordinate As Long, ByVal str As String, Optional num As Integer = 1) As Long
    Dim c As String
    Dim flag As Boolean
    Dim stack As Integer
    Dim index, lenStr As Long
    
    lenStr = Len(str)
    index = coordinate
    stack = 0
    flag = True
    
    Do
        If coordinate > lenStr Then
            nextRightBrace = index
            Exit Function
        End If
            
        c = Mid(str, coordinate, 1)
         If c = "{" Then
            stack = stack + 1
            coordinate = coordinate + 1
        ElseIf c = "}" Then
            stack = stack - 1
            If stack < 0 Then
                If num = 1 Then
                    flag = False
                Else
                    num = num - 1
                    coordinate = coordinate + 1
                End If
            Else
                coordinate = coordinate + 1
            End If
        ElseIf c = "\" Then
            If Mid(str, coordinate + 1, 4) = "left" Then
                coordinate = coordinate + 5
                If Mid(str, coordinate, 1) = "\" Then
                    coordinate = coordinate + 2
                Else
                    coordinate = coordinate + 1
                End If
            ElseIf Mid(str, coordinate + 1, 5) = "right" Then
                coordinate = coordinate + 6
                If Mid(str, coordinate, 1) = "\" Then
                    coordinate = coordinate + 2
                Else
                    coordinate = coordinate + 1
                End If
            Else
                coordinate = coordinate + 2
            End If
        Else
            coordinate = coordinate + 1
        End If
        DoEvents
    Loop While flag

    nextRightBrace = coordinate
End Function


Function nextRightSquareBrackets(ByRef coordinate As Long, ByVal str As String, Optional num As Integer = 1) As Boolean
    Dim c As String
    Dim flag As Boolean
    Dim stack As Integer
    Dim index, lenStr As Long
    
    lenStr = Len(str)
    index = coordinate
    stack = 0
    flag = True
    
    Do
        If coordinate > lenStr Then
            coordinate = index
            nextRightSquareBrackets = False
            Exit Function
        End If
            
        c = Mid(str, coordinate, 1)
         If c = "[" Then
            stack = stack + 1
            coordinate = coordinate + 1
        ElseIf c = "]" Then
            stack = stack - 1
            If stack < 0 Then
                If num = 1 Then
                    flag = False
                Else
                    num = num - 1
                    coordinate = coordinate + 1
                End If
            Else
                coordinate = coordinate + 1
            End If
        ElseIf c = "\" Then
            If Mid(str, coordinate + 1, 4) = "left" Then
                coordinate = coordinate + 5
                If Mid(str, coordinate, 1) = "\" Then
                    coordinate = coordinate + 2
                Else
                    coordinate = coordinate + 1
                End If
            ElseIf Mid(str, coordinate + 1, 5) = "right" Then
                coordinate = coordinate + 6
                If Mid(str, coordinate, 1) = "\" Then
                    coordinate = coordinate + 2
                Else
                    coordinate = coordinate + 1
                End If
            Else
                coordinate = coordinate + 2
            End If
        Else
            coordinate = coordinate + 1
        End If
        DoEvents
    Loop While flag
    nextRightSquareBrackets = True
End Function

Function readINI() As Boolean
'    questionID = 题目ID号
'    figID = 图片ID号
'    tabID = 表格ID号
    readINI = True
    Counter = GetIniLong("Counter", "No.")
    If Counter = -1 Then
        readINI = False
        Exit Function
    End If
    
    questionID = GetIniLong("ID" + CStr(Counter), "questionID")
    If questionID = -1 Then
        readINI = False
        Exit Function
    End If
    
    figID = GetIniLong("ID" + CStr(Counter), "FigID")
    If figID = -1 Then
        readINI = False
        Exit Function
    End If
    
    tabID = GetIniLong("ID" + CStr(Counter), "TabID")
    If tabID = -1 Then
        readINI = False
        Exit Function
    End If
End Function

Function writeINI(ByVal strFileName As String)
    Counter = Counter + 1
    WriteIniLong "Counter", "No.", CStr(Counter)
    WriteIniLong "ID" + CStr(Counter), "questionID", CStr(questionID)
    WriteIniLong "ID" + CStr(Counter), "FigID", CStr(figID)
    WriteIniLong "ID" + CStr(Counter), "TabID", CStr(tabID)
    WriteIniLong "ID" + CStr(Counter), "Time", CStr(Date) + "  " + CStr(Time) + "  " + strFileName
End Function

Function writeTex(str As String, strPathName As String)
Dim adostream As New ADODB.Stream
With adostream
.Type = adTypeText
.Mode = adModeReadWrite
.Charset = "utf-8"
.Open
.Position = 0
.WriteText str
.SaveToFile strPathName, adSaveCreateOverWrite
.Close
End With
Set adostream = Nothing
End Function

Function readUTF8(ByVal texFile As String, ByRef str As String)

'2.VB读取utf-8文本文件
    
    Dim adostream As New ADODB.Stream
    With adostream
    .Type = adTypeText
    .Mode = adModeReadWrite
    .Charset = "utf-8"
    .Open
    .LoadFromFile texFile
    str = .ReadText
    .Close
    End With
    Set adostream = Nothing
    
End Function


Public Function delLeftRight(ByVal str As String) As String
    Dim mylink As DBLink
    Dim i As Long
    Dim strTemp As String
    Set mylink = New DBLink
    'str = Text1.Text '"5．已知向量$\vec{a},\vec{b}$满足$\left| \vec{a}\right| =3,\left| \frac{\left(a\right)}{b}\right| =2\sqrt{3}$，且$\vec{a}\bot \left(\vec{a}+\vec{b}\right)$，则$\vec{b}$在$\vec{a}$方向上的投影为（ ）"
    If splitNode(mylink, str, 0) = False Then Exit Function
    
    i = mylink.getFirstIndex
    Do
        mylink.finished = True
        Do While i <> 0
            If mylink.getNodeFlag(i) = False Then
                If splitNode(mylink, mylink.getNode(i), i) = False Then Exit Function
            End If
            i = mylink.getNextIndex(i)
        Loop
    Loop While mylink.finished = False
    delLeftRight = mylink.printDBLink
End Function

Private Function splitNode(ByRef mylink As DBLink, ByVal str As String, ByVal index As Long) As Boolean
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    'Dim mMatch As Object        '匹配字符串
    'Dim str As String
    Dim prev As Long
    Dim first As Long
    Dim last As Long
    Dim c As Long
    Dim strTemp As String
    Dim textStr As String
    Dim i As Long, j As Long, k As Long

    splitNode = True
    Set re = New RegExp
    re.Pattern = "(\\left(\.|\(|\[|\\\{|\||\\\||\\[a-zA-Z]+|/|\)|\]|\\\}))|(\\right(\.|\)|\]|\\\}|\||\\\||\\[a-zA-Z]+|\(|\[|\\\{|/))"
    re.Global = True
    're.IgnoreCase = True
    prev = 1
    
    'str = "5．已知向量$\vec{a},\vec{b}$满足$\left| \vec{a}\right| =3,\left| \frac{a}{b}\right| =2\sqrt{3}$，且$\vec{a}\bot \left(\vec{a}+\vec{b}\right)$，则$\vec{b}$在$\vec{a}$方向上的投影为（ ）"
    If re.test(str) = True Then
        mylink.finished = False
    Else
        If index = 0 Then
            k = mylink.creatNode(str, True)
            mylink.replaceNode index, k
        Else
            mylink.setNodeFlag index, True
        End If
        Exit Function
    End If
    Set mMatches = re.Execute(str)
    first = mylink.creatNode("", True)
    last = first
    For i = 0 To mMatches.Count - 1
        If mMatches(i).FirstIndex + 1 > prev Then
            strTemp = Mid(str, prev, mMatches(i).FirstIndex + 1 - prev)
            k = mylink.creatNode(strTemp)
            mylink.linkTwo last, k
        End If
        
        If Left(CStr(mMatches(i).Value), 5) = "\left" Then
            j = i + 1
            If j < mMatches.Count Then
                stack = 0
                For j = j To mMatches.Count - 1
                    If Left(CStr(mMatches(j).Value), 5) = "\left" Then
                        stack = stack + 1
                    Else
                        stack = stack - 1
                    End If
                    If stack < 0 Then Exit For
                Next
                If j >= mMatches.Count Then
                    MsgBox "have no right" & Chr(13) & Mid(str, mMatches(i).FirstIndex, 50)
                    splitNode = False
                    Exit Function
                Else
                    textStr = Mid(str, mMatches(i).FirstIndex + mMatches(i).Length + 1, mMatches(j).FirstIndex - mMatches(i).FirstIndex - mMatches(i).Length)
                    leftstr = CStr(mMatches(i).Value)
                    rightstr = CStr(mMatches(j).Value)
                    If ifneedLR(textStr) = True Then
                        k = mylink.creatNode(leftstr, True)
                        mylink.linkTwo last, k
                        k = mylink.creatNode(textStr)
                        mylink.linkTwo last, k
                        k = mylink.creatNode(rightstr, True)
                        mylink.linkTwo last, k
                    Else
                        leftstr = Mid(leftstr, 6)
                        rightstr = Mid(rightstr, 7)
                        If leftstr = "." Then
                            leftstr = ""
                        End If
                        If rightstr = "." Then
                            rightstr = ""
                        End If
                        k = mylink.creatNode(leftstr + textStr + rightstr)
                        mylink.linkTwo last, k
                    End If
                    i = j
                    prev = mMatches(j).FirstIndex + mMatches(j).Length + 1
                End If
            Else
                MsgBox "have no right here" & Chr(13) & Mid(str, mMatches(i).FirstIndex, 50)
                splitNode = False
                Exit Function
            End If
        Else
            MsgBox "first not left"
            splitNode = False
            Exit Function
        End If
    Next
    strTemp = Mid(str, prev)
    k = mylink.creatNode(strTemp, True)
    mylink.linkTwo last, k
    mylink.replaceNode index, first, last
End Function

Public Function correctLeftRight(ByRef str As String) As Boolean
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim nMatches As Object      '匹配字符串集合对象
    Dim prev As Long
    Dim strTemp As String
    Dim textStr As String
    Dim i As Long, j As Long, k As Long, l As Long

    correctLeftRight = True
    Set re = New RegExp
    Set reN = New RegExp
    re.Pattern = "(\\left(\.|\(|\[|\\\{|\||\\\||\\[a-zA-Z]+|/|\)|\]|\\\}))|(\\right(\.|\)|\]|\\\}|\||\\\||\\[a-zA-Z]+|\(|\[|\\\{|/))"
    re.Global = True
    prev = 1
    
    Set mMatches = re.Execute(str)
    For i = 0 To mMatches.Count - 1
        If mMatches(i).FirstIndex + 1 > prev Then
            strTemp = strTemp + Mid(str, prev, mMatches(i).FirstIndex + 1 - prev)
        End If
        If Left(CStr(mMatches(i).Value), 5) = "\left" Then
            j = i + 1
            If j < mMatches.Count Then
                stack = 0
                For j = j To mMatches.Count - 1
                    If Left(CStr(mMatches(j).Value), 5) = "\left" Then
                        stack = stack + 1
                    Else
                        stack = stack - 1
                    End If
                    If stack < 0 Then Exit For
                Next
                If j >= mMatches.Count Then
                    MsgBox "have no right"
                    correctLeftRight = False
                    Exit Function
                Else
                    textStr = Mid(str, mMatches(i).FirstIndex + 1, mMatches(j).FirstIndex - mMatches(i).FirstIndex)
                    insertTextCmd textStr
                    strTemp = strTemp + textStr
                    i = j
                    prev = mMatches(j).FirstIndex + 1
                End If
            Else
                MsgBox "have no right here"
                correctLeftRight = False
                Exit Function
            End If
        Else
            MsgBox "first not left"
            correctLeftRight = False
            Exit Function
        End If
    Next
    str = strTemp + Mid(str, prev)
End Function

Function insertTextCmd(ByRef str As String)
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    Dim prev As Long
    Dim strTemp, tempXiuZheng As String
    Dim n As Long
    
    tempXiuZheng = ""
    Set re = New RegExp
    re.Global = True
    prev = 1
    
    re.Pattern = "\$"
    Set mMatches = re.Execute(str)
    If mMatches.Count > 0 Then
        For n = 0 To mMatches.Count - 1 Step 2
            tempXiuZheng = tempXiuZheng + Mid(str, prev, mMatches(n).FirstIndex - prev + 1)
            strTemp = Mid(str, mMatches(n).FirstIndex + mMatches(n).Length + 1, mMatches(n + 1).FirstIndex - mMatches(n).FirstIndex - mMatches(n).Length)
            
            tempXiuZheng = tempXiuZheng + "\text{" + strTemp + "}"
            prev = mMatches(n + 1).FirstIndex + mMatches(n + 1).Length + 1
            DoEvents
        Next n
        str = tempXiuZheng + Mid(str, prev)
    End If

End Function

Function ifneedLR(str As String) As Boolean

    Dim re As Object
    
    Set re = New RegExp
    re.Pattern = needLeftRightList
    If re.test(str) = True Then
        ifneedLR = True
    Else
        ifneedLR = False
    End If
End Function

Public Function replaceSymbol(ByRef str As String, ByVal strName As String)
    'Dim str As String
    Dim fr() As String
    Dim re As Object
    Dim mMatches As Object
    
    Set re = New RegExp
    re.Global = True
    replaceSymbolList
'替换列表非空
    If UBound(strReplaceSymbolList) >= 0 Then
        For i = 0 To UBound(strReplaceSymbolList)
            If Trim(strReplaceSymbolList(i)) <> "" Then '非空行，用“;”进行分割
                fr = Split(strReplaceSymbolList(i), ";")
                If UBound(fr) = 1 Then
                    If Left(fr(0), 1) = "'" Then
                    Else
                        re.Pattern = fr(0)
                        str = re.Replace(str, fr(1))
                        'str = Replace(str, fr(0), fr(1))
                    End If
                Else
                    If Left(fr(0), 1) = "'" Then
                    Else
                        MsgBox "第" & i + 1 & "行  " & strReplaceSymbolList(i)
                        Exit Function
                    End If
                End If
            End If
        Next
    End If
End Function

Private Function replaceSymbolList() As Boolean
    Dim TextLine As String
    Dim strTemp As String
    'Dim replaceSymbolListFile As String
    'replaceSymbolListFile = App.Path & "\replaceSymbolList.txt"
    On Error GoTo err1
    readUTF8 replaceSymbolListFile, strTemp
    strTemp = Replace(strTemp, Chr(10), "")
    strReplaceSymbolList = Split(strTemp, Chr(13))
    replaceSymbolList = True
    Exit Function
err1:
    If Err.Number = 53 Then
        MsgBox replaceSymbolListFile & " 未找到！"
    Else
        MsgBox Err.Number & Chr(13) & "请查看" & replaceSymbolListFile & "路径是否正确"
    End If
    strReplaceSymbolList = Split("", " ")
    replaceSymbolList = False
End Function

Public Function readReplaceList() As Boolean
    Dim TextLine() As String
    Dim fr() As String
    Dim strTemp As String
    Dim i As Long, j As Long
    Dim re As Object
    
    Set re = New RegExp
    re.Global = True
    
    On Error GoTo err1
    readUTF8 replaceListFile, strTemp
    strTemp = Replace(strTemp, Chr(10), "")
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    TextLine = Split(strTemp, Chr(13))
    ReDim strReplaceList(UBound(TextLine), 1)
    j = -1
    For i = 0 To UBound(TextLine)
        fr = Split(TextLine(i), ";")
        If UBound(fr) = 1 Then
            j = j + 1
            strReplaceList(j, 0) = fr(0)
            strReplaceList(j, 1) = fr(1)
        Else
            readReplaceList = False
            Exit Function
        End If
    Next
    ReDim Preserve strReplaceList(j, 1)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    readReplaceList = True
    Exit Function
err1:
    If Err.Number = 53 Then
        MsgBox replaceListFile & " 未找到！"
    Else
        MsgBox Err.Number & Chr(13) & "请查看" & replaceSymbolListFile & "路径是否正确"
    End If
    readReplaceList = False
End Function

Public Function replaceMacros()
    Dim str As String
    Dim re As Object
    Dim mMatches As Object
    
    Set re = New RegExp
    re.Global = True
    
    Dim docxFileName
    For Each docxFileName In strFullName
        readUTF8 docxFileName, str
        replaceList str
        writeTex str, CStr(docxFileName)
    Next
End Function

Function replaceList(ByRef str As String)
    Dim re As Object
    Dim mMatches As Object
    
    Set re = New RegExp
    re.Global = True
    
    For i = 0 To UBound(strReplaceList)
        If Left(strReplaceList(i, 0), 1) = "#" Then
            re.Pattern = Replace(Mid(strReplaceList(i, 0), 2), "\n", Chr(13))
            str = re.Replace(str, Replace(strReplaceList(i, 1), "\n", Chr(13)))
        Else
            str = Replace(str, strReplaceList(i, 0), strReplaceList(i, 1))
        End If
    Next
    
    If correctLeftRightFlag = True Then
        str = delLeftRight(str)
        correctLeftRight str
    End If
End Function

Function correctAlign(ByRef str As String)
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    Dim prev As Long
    Dim strTemp, tempXiuZheng As String
    Dim n As Long
    
    tempXiuZheng = ""
    Set re = New RegExp
    re.Global = True
    prev = 1
    
    re.Pattern = "(\\begin\{align\*\})|(\\end\{align\*\})"
    Set mMatches = re.Execute(str)
    If mMatches.Count > 0 Then
        For n = 0 To mMatches.Count - 1 Step 2
            tempXiuZheng = tempXiuZheng + Mid(str, prev, mMatches(n).FirstIndex - prev + 1)
            strTemp = Mid(str, mMatches(n).FirstIndex + mMatches(n).Length + 1, mMatches(n + 1).FirstIndex - mMatches(n).FirstIndex - mMatches(n).Length)
            
            re.Pattern = "\$"
            If re.test(strTemp) Then
                strTemp = re.Replace(strTemp, "")
            End If
            re.Pattern = "\&"
            If re.test(strTemp) Then
                strTemp = re.Replace(strTemp, "")
            End If
            tempXiuZheng = tempXiuZheng + TrimEnter(strTemp)
            prev = mMatches(n + 1).FirstIndex + mMatches(n + 1).Length + 1
            DoEvents
        Next n
        str = tempXiuZheng + Mid(str, prev)
    End If

End Function

Function DelTensor(ByRef s As String)
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    Dim str As String, strTemp As String
    Dim prev As Long, index As Long
    
    prev = 1
    Set re = New RegExp
    re.Pattern = "\\tensor[\*]?\["
    re.Global = True
    
    Set mMatches = re.Execute(s)
    If mMatches.Count > 0 Then
        For Each mMatch In mMatches
            If prev <= mMatch.FirstIndex + 1 Then
                str = str + Mid(s, prev, mMatch.FirstIndex + 1 - prev)
                prev = mMatch.FirstIndex + mMatch.Length + 1
                If nextRightSquareBrackets(prev, s) = True Then                     '截取方括号内参数
                    strTemp = Mid(s, mMatch.FirstIndex + mMatch.Length + 1, prev - mMatch.FirstIndex - mMatch.Length - 1)
                End If
                prev = nextLeftBrace(prev + 1, s)
                index = nextRightBrace(prev + 1, s)
                strTemp = strTemp + Mid(s, prev + 1, index - prev - 1)              '截取第一花括号内参数
                prev = index + 1
                prev = nextLeftBrace(prev, s)
                index = nextRightBrace(prev + 1, s)
                strTemp = strTemp + Mid(s, prev + 1, index - prev - 1)              '截取第二花括号内参数
                strTemp = Replace(strTemp, "^{}", "")
                strTemp = Replace(strTemp, "_{}", "")
                str = str + strTemp
                prev = index + 1
            End If
            DoEvents
        Next
        str = str + Mid(s, prev)
        Debug.Print str
        s = str
    End If
End Function
