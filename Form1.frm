VERSION 5.00
Object = "{0DF5D14C-08DD-4806-8BE2-B59CB924CFC9}#1.7#0"; "VBCCR16.OCX"
Begin VB.Form form_tex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "docx转tex及修正"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9555
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   9555
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command15 
      Caption         =   "Join"
      Height          =   495
      Left            =   8880
      TabIndex        =   43
      Top             =   2040
      Width           =   615
   End
   Begin VB.Frame Frame8 
      Caption         =   "首次修正字符列表"
      Height          =   855
      Left            =   4680
      TabIndex        =   40
      Top             =   1200
      Width           =   4815
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Text            =   "Form1.frx":0000
         Top             =   240
         Width           =   3735
      End
      Begin VB.CommandButton Command14 
         Caption         =   "选择"
         Height          =   495
         Left            =   4080
         TabIndex        =   41
         Top             =   240
         Width           =   615
      End
   End
   Begin VBCCR16.RichTextBox RichTextBox1 
      Height          =   5295
      Left            =   4680
      TabIndex        =   39
      Top             =   2640
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiLine       =   -1  'True
      ScrollBars      =   2
      TextRTF         =   "Form1.frx":0007
   End
   Begin VB.CommandButton Command13 
      Caption         =   "重置ID"
      Height          =   495
      Left            =   3600
      TabIndex        =   37
      Top             =   5280
      Width           =   735
   End
   Begin VB.Frame Frame6 
      Caption         =   "docx2tex路径设置"
      Height          =   855
      Left            =   4680
      TabIndex        =   33
      Top             =   240
      Width           =   4815
      Begin VB.CommandButton Command12 
         Caption         =   "选择"
         Height          =   495
         Left            =   4080
         TabIndex        =   35
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   34
         Text            =   "Form1.frx":016D
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "清空"
      Height          =   495
      Left            =   7800
      TabIndex        =   32
      Top             =   2040
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "加$"
      Height          =   180
      Left            =   3960
      TabIndex        =   31
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "加数学环境"
      Height          =   495
      Left            =   4920
      TabIndex        =   30
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "复制"
      Height          =   495
      Left            =   6480
      TabIndex        =   29
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "..."
      Height          =   495
      Left            =   3840
      TabIndex        =   24
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "遍历替换tex"
      Height          =   495
      Left            =   240
      TabIndex        =   19
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "Form1.frx":0173
      Top             =   4080
      Width           =   3615
   End
   Begin VB.Frame Frame3 
      Caption         =   "ID设置"
      Height          =   2895
      Left            =   2040
      TabIndex        =   8
      Top             =   840
      Width           =   2415
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1000
         TabIndex        =   17
         Text            =   "Text4"
         Top             =   2200
         Width           =   1000
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1000
         TabIndex        =   16
         Text            =   "Text3"
         Top             =   1700
         Width           =   1000
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1000
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   1200
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1000
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   700
         Width           =   1000
      End
      Begin VB.CheckBox Check1 
         Caption         =   "使用临时ID"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label Label5 
         Caption         =   "figID"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   13
         Top             =   2300
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "tabID"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   12
         Top             =   1750
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "queID"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   11
         Top             =   1250
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   10
         Top             =   800
         Width           =   600
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "转换+修正"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   200
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "未识别题目数"
      Height          =   800
      Left            =   200
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
      Begin VB.Label numLabel 
         Height          =   350
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "tex修正"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   200
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "docx转换tex"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "状态"
      Height          =   1155
      Left            =   200
      TabIndex        =   4
      Top             =   840
      Width           =   1575
      Begin VB.Label texFlagLabel 
         Height          =   600
         Left            =   120
         TabIndex        =   5
         Top             =   285
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "tex文件自定义替换"
      Height          =   2895
      Left            =   120
      TabIndex        =   20
      Top             =   5040
      Width           =   4335
      Begin VB.CommandButton Command8 
         Caption         =   "修正ID"
         Height          =   495
         Left            =   2760
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H80000016&
         Height          =   615
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   1680
         Width           =   3855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "list"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   2330
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "选择替换tex"
         Height          =   495
         Left            =   1440
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   1080
         TabIndex        =   21
         Text            =   "双击选择文件夹"
         Top             =   840
         Width           =   3015
      End
      Begin VB.Frame Frame5 
         Caption         =   "替换内容列表文件"
         Height          =   1335
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   4095
         Begin VB.CheckBox Check3 
            Caption         =   "修正\left\right"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   38
            Top             =   870
            Width           =   2640
         End
      End
      Begin VB.Label label6 
         Caption         =   "遍  历文件夹"
         Height          =   385
         Left            =   360
         TabIndex        =   25
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "SourceDB.ini 路径"
      Height          =   1095
      Left            =   120
      TabIndex        =   36
      Top             =   3840
      Width           =   4335
   End
   Begin VB.Label Label1 
      DataMember      =   "240"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "form_tex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Ctrl
Private Sub Check1_Click()
    If Check1.Value = 0 Then
        ifReadINI = True
        setID
        textUnEnabled
    ElseIf Check1.Value = 1 Then
        ifReadINI = False
        textEnabled
    End If
    IDFresh
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        addDollorFlag = True
        form_tex.Width = 9800
    ElseIf Check2.Value = 0 Then
        addDollorFlag = False
        form_tex.Width = 4845
    End If
End Sub


Private Sub Check3_Click()
    If Check3.Value = 1 Then
        correctLeftRightFlag = True
    ElseIf Check3.Value = 0 Then
        correctLeftRightFlag = False
    End If
End Sub

Private Sub Command1_Click()
    Dim tt As String
    Unidentified = 0
    numLabel.Caption = CStr(Unidentified)
    
    texFlagLabel.Caption = "处理中……"
    
    If setID Then
        tt = Main
        IDFresh
        If Unidentified <> 0 Then
            numLabel.Caption = CStr(Unidentified)
        End If
        texFlagLabel.Caption = "完成！" + Chr(13) + tt
    Else
        MsgBox "ID设置错误！"
        texFlagLabel.Caption = "ID设置错误！"
    End If
End Sub

Private Sub Command10_Click()
    Dim str As String
    str = RichTextBox1.Text
    replaceSymbol str, ""
    insertDollerT str
    If correctLeftRightFlag = True Then
        str = delLeftRight(str)
        correctLeftRight str
    End If
        correctEnvs str
        correctMathScript str, "_\{"
        correctMathScript str, "\\dfrac\{", 2
        str = delLeftRight(str)
        correctLeftRight str
    readReplaceList
    replaceList str
    RichTextBox1.Text = str
End Sub

Private Sub Command11_Click()
    RichTextBox1.Text = ""
End Sub

Private Sub Command12_Click()
    Dim getdir As String

    getdir = BrowseForFolder(Me, "选择工作文件夹", Text9.Text)

    If Len(getdir) = 0 Then Exit Sub
    Text9.Text = IIf(Right$(getdir, 1) = "\", getdir, getdir & "\")
    docxToTexPath = Text9.Text
End Sub

Private Sub Command13_Click()
    redistributeAll
End Sub

Private Sub Command14_Click()
    strFullName = selectFile(3)
    If strFullName(0) <> "" Then
        replaceSymbolListFile = strFullName(0)
        Text10.Text = replaceSymbolListFile
    End If
End Sub

Private Sub Command15_Click()
    texFlagLabel.Caption = "处理中……"
    Ctrl = GetKeyState(vbKeyControl)
    If Ctrl < 0 Then
        'MsgBox "Ctrl+单击"
        RichTextBox1.Text = JoinTest(True)
    Else
        'MsgBox "单击"
        RichTextBox1.Text = JoinTest
    End If
    texFlagLabel.Caption = "完成！"
End Sub

Private Sub Command2_Click()
    Dim tt As String
    texFlagLabel.Caption = "处理中……"
    If setID Then
        tt = changeToTex
    Else
        MsgBox "ID设置错误！"
    End If
    texFlagLabel.Caption = "完成！" + Chr(13) + tt
End Sub

Private Sub Command3_Click()
    texFlagLabel.Caption = "处理中……"
    If setID Then
        convertToTex
        IDFresh
    Else
        MsgBox "ID设置错误！"
    End If
    fileSelect = False
    texFlagLabel.Caption = "完成！"
End Sub

Private Sub Command4_Click()
    texFlagLabel.Caption = "处理中……"

    If findTexFile = False Then
        texFlagLabel.Caption = "遍历文件夹未设置！"
        Exit Sub
    End If
    If readReplaceList = False Then
        texFlagLabel.Caption = "退出"
        Exit Sub
    End If
    If strFullName(0) <> "" Then
        replaceMacros
        MsgBox "共 " & UBound(strFullName) & " 个文件"
    Else
        MsgBox "none"
    End If
    texFlagLabel.Caption = "完成"
End Sub
Function findTexFile() As Boolean
    Dim str1() As String
    Dim a As Long, b As Long, c As Long
    Dim strg As String, fz As String
    Dim strTemp As String
    
    findTexFile = False
    If Text6 = "" Then Exit Function
        fz = Text6 '文件夹"d:\tl\" '
        On Error GoTo XH1
XH1:
        Resume Next
        strg = Dir(fz, 63)
XH:
    If strg <> "" Then
        If (GetAttr(fz & strg) <> 9238) And (GetAttr(fz & strg) And 16) = 16 And Right(strg, 1) <> "." Then
            a = a + 1            'Label3.Caption = a
        ReDim Preserve str1(a)   '保存上次运行搜索出的目录
            str1(a) = fz & strg & "\"    '保存上次运行搜索出的目录
          'List1.AddItem str1(a)   '显示搜索的目录
        ElseIf Right(strg, 1) <> "." Then '
          If Right(strg, 4) = ".tex" Then '
            c = c + 1            'Label4.Caption = c
            strTemp = strTemp + "?" + fz & strg
          End If
          'List2.AddItem fz & strg  '显示出搜索的文件
        End If
        strg = Dir           '继续，然后跳回XH:
        'Label3.Caption = a   '目录数
        'Label4.Caption = c   '文件数
        DoEvents             '允许其他事件发生。
        GoTo XH                  '跳到XH:检测是否为空跳转ElseIf
    ElseIf strg = "" And a <> 0 Then  'strg=""说明以上代码完成，a<>0是记录目录不为空，下面代码继续。
        b = b + 1            '以上代码遍历完成，查看是否有新搜索出的目录层，+1逐一搜索，下面b>a检测。
        If b > a Then
            If c > 0 Then
                strFullName = Split(Right(strTemp, Len(strTemp) - 1), "?")
                findTexFile = True
            End If
            Exit Function      '记录的文件夹超出就证明全局遍历完成。
        End If
        fz = str1(b)         '赋值待搜索目录。
        GoTo XH1                 '跳到XH1:标记处执行待搜索，赋值，。
    End If
    If c > 0 Then
        strFullName = Split(Right(strTemp, Len(strTemp) - 1), "?")
        findTexFile = True
    End If
End Function

Private Sub Command5_Click()
    texFlagLabel.Caption = "处理中……"
    If readReplaceList = False Then Exit Sub
    strFullName = selectFile(1)
    If strFullName(0) <> "" Then
        replaceMacros
    End If
    texFlagLabel.Caption = "完成！"
End Sub

Private Sub Command6_Click()
    strFullName = selectFile(3)
    If strFullName(0) <> "" Then
        replaceListFile = strFullName(0)
        Text7.Text = replaceListFile
    End If
End Sub

Private Sub Command7_Click()
    'Dim getdir As String
    Dim str() As String
    'getdir = BrowseForFolder(Me, "选择工作文件夹", Text5.Text)

    'If Len(getdir) = 0 Then Exit Sub
    'Text5.Text = IIf(Right$(getdir, 1) = "\", getdir, getdir & "\")
    strFullName = selectFile(4)
    If strFullName(0) <> "" Then
        str = Split(strFullName(0), ".")
        If UBound(str) > 0 And UCase(str(UBound(str))) = "INI" Then
            bookPath = strFullName(0)
            Text5.Text = bookPath
            Check1.Value = 0
            ifReadINI = True
            setID
            textUnEnabled
        Else
            MsgBox "INI设置错误！"
        End If
    End If
End Sub

Private Sub Command8_Click()
    texFlagLabel.Caption = "开始！"
    Redistribution
    IDFresh
    texFlagLabel.Caption = "完成！"
End Sub

Private Sub Command9_Click()
Clipboard.Clear
Clipboard.SetText (RichTextBox1.Text)
End Sub

Private Sub Form_DblClick()
    texFlagLabel.Caption = "开始！"
    'onlyToTex
    'beforeChange
    'correcttest
    'MsgBox "questionID:" & questionID & " FigID:" & figID & " TabID:" & tabID
    MsgBox nextRightBrace(1, "$M=\left\{x| \dfrac{x}{x-1}\leqslant 0\right\},N=\{.x| x^{2}-2x<0\}$，")
    MsgBox "$" & TrimEnter(Chr(13) + Chr(13) + "start" + Chr(13), "R") & "$"
    texFlagLabel.Caption = "完成！"
End Sub

Function redistributeAll()
    texFlagLabel.Caption = "处理中……"

    If findTexFile = False Then
        texFlagLabel.Caption = "遍历文件夹未设置或未找到tex文件！"
        Exit Function
    End If
    If strFullName(0) <> "" Then
        fileSelect = True
        Redistribution
        MsgBox "共 " & UBound(strFullName) & " 个文件"
        IDFresh
        texFlagLabel.Caption = "完成"
    Else
        MsgBox "请设置遍历文件夹！"
    End If
End Function
Sub test()
    Dim re As Object
    Dim str As String
    Dim s As String
    Dim prev As Long
    Dim mMatches As Object      '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    Set re = New RegExp
    re.Global = True
    re.Pattern = "\\item[XTJFM]\{"
    prev = 1
    s = "dfs\itemX{1}{在等腰$\bigtriangleup ABC$，$AB=AC$，$BC=6$，向量$\vec{AD}=\vec{DC}$，则$\vec{DC}\cdot \vec{BC}$的值为\xz }{$9$}{$18$}{$27$}{$36$}{$A$}"
    Set mMatches = re.Execute(s)
    For Each mMatch In mMatches
        If prev < mMatch.FirstIndex Then
            str = str + Mid(s, prev, mMatch.FirstIndex + 1 + mMatch.Length - prev)
            prev = nextRightBrace(mMatch.FirstIndex + mMatch.Length + 1, s)
            prev = prev + 1
            nextLBrace prev, s
            prev = nextRightBrace(prev + 1, s)
        End If
    Next
    MsgBox str
End Sub
Private Sub Form_Initialize()
    fileSelect = False
    ifReadINI = True
    form_tex.Width = 4695
    readAppIni
    setID
    textUnEnabled
    addDollorFlag = False
    correctLeftRightFlag = False
    Text5.Text = bookPath
    Text9.Text = docxToTexPath
    Text10.Text = replaceSymbolListFile
    Text7.Text = replaceListFile
    'Text7.Enabled = False
End Sub

Function readAppIni()
    Dim str As String
    str = GetAppINI("fullFileName", "bookPath")
    If str <> "" Then
        bookPath = str
    Else
        bookPath = "D:\test\book1\SourceDB.ini"
    End If
    
    str = GetAppINI("fullFileName", "docxToTexPath")
    If str <> "" Then
        docxToTexPath = str
    Else
        docxToTexPath = "D:\docx2tex\"
    End If
    
    str = GetAppINI("fullFileName", "replaceSymbolListFile")    '初次替换列表
    If str <> "" Then
        replaceSymbolListFile = str
    Else
        replaceSymbolListFile = App.Path & "\replaceSymbolList.txt"
    End If
    
    str = GetAppINI("fullFileName", "replaceListFile")          '二次修正列表
    If str <> "" Then
        replaceListFile = str
    Else
        replaceListFile = App.Path & "\replaceList.txt"
    End If
    
    str = GetAppINI("fullFileName", "needLeftRightList")
    If str <> "" Then
        needLeftRightList = str
    Else
        needLeftRightList = "\\dfrac"
    End If
    
    str = GetAppINI("fullFileName", "braceCMDList")
    If str <> "" Then
        braceCMDList = str
    Else
        braceCMDList = "" '"\\textbf\{|\\text\{|\\textit\{|\\mathrm\{|\\boldsymbol\{|\\textcolor\{color-[0-9]\}\{|\\underline\{"
    End If
    
    str = GetAppINI("fullFileName", "questionTypeXZ")
    str = Replace(str, Chr(0), "")
    If str <> "" Then
        questionTypeXZ = str
    Else
        questionTypeXZ = "选择题"
    End If
    
    str = GetAppINI("fullFileName", "questionTypeTK")
    str = Replace(str, Chr(0), "")
    If str <> "" Then
        questionTypeTK = str
    Else
        questionTypeTK = "填空题"
    End If
    
    str = GetAppINI("fullFileName", "questionTypeJD")
    str = Replace(str, Chr(0), "")
    If str <> "" Then
        questionTypeJD = str
    Else
        questionTypeJD = "解答题"
    End If
    
    str = GetAppINI("fullFileName", "questionAnswerBoundary")
    str = Replace(str, Chr(0), "")
    If str <> "" Then
        questionAnswerBoundary = str
    Else
        questionAnswerBoundary = "【解析】"
    End If
    
    str = GetAppINI("fullFileName", "answerBoundary")
    str = Replace(str, Chr(0), "")
    If str <> "" Then
        answerBoundary = str
    Else
        answerBoundary = "参考答案"
    End If
    
    str = GetAppINI("fullFileName", "correctEnvironments")
    If str <> "" Then
        correctEnvironments = Split(str, ",")
    Else
        correctEnvironments = Split("", ",")
    End If
    
    str = GetAppINI("fullFileName", "delDoller")
    If str <> "" Then
        delDollerList = Split(str, ",")
    Else
        delDollerList = Split("", ",")
    End If
    
    str = GetAppINI("fullFileName", "correctMathScript")
    If str <> "" Then
        mathScriptList = Split(str, ",")
    Else
        mathScriptList = Split("", ",")
    End If
'Public AnswerSolutionBoundary As String
'Public SolutionStart As String
'Public SolutionEnd As String
    '答案中简答和详解提取
    str = GetAppINI("answer", "AnswerSolutionBoundary")
    str = Replace(str, Chr(0), "")
    If str <> "" Then
        answerSolutionBoundary = str
    Else
        answerSolutionBoundary = "[【]?解析[】]?"
    End If
    str = GetAppINI("answer", "SolutionStart")
    str = Replace(str, Chr(0), "")
    If str <> "" Then
        solutionStart = str
    Else
        solutionStart = "[【]?详解[】]?"
    End If
    str = GetAppINI("answer", "SolutionEnd")
    str = Replace(str, Chr(0), "")
    If str <> "" Then
        solutionEnd = str
    Else
        solutionEnd = "[【]?点睛[】]?"
    End If
End Function

Function textEnabled()
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    
    Text2.Text = "0"
    Text3.Text = "0"
    Text4.Text = "0"
    
    Text2.BackColor = &H80000005
    Text3.BackColor = &H80000005
    Text4.BackColor = &H80000005
End Function

Function textUnEnabled()
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    
    IDFresh

    Text2.BackColor = &H80000016
    Text3.BackColor = &H80000016
    Text4.BackColor = &H80000016
End Function

Public Function setID() As Boolean
    setID = True
    If ifReadINI = True Then
        If readINI = False Then
            MsgBox "配置文件读取错误！"
            setID = False
            Exit Function
        End If
    Else
        Text1.Text = CStr(Counter)
        questionID = CLng(Text2.Text)
        tabID = CLng(Text3.Text)
        figID = CLng(Text4.Text)
    End If
End Function

Function IDFresh()
    Text1.Text = CStr(Counter)
    Text2.Text = CStr(questionID)
    Text3.Text = CStr(tabID)
    Text4.Text = CStr(figID)
End Function

Private Sub Form_Resize()
    If addDollorFlag = True Then
        form_tex.Width = 9800
        form_tex.Height = 8475
    Else
        form_tex.Width = 4750
        form_tex.Height = 8475
    End If
End Sub

Private Sub Form_Terminate()
    'MsgBox "quite"
    WriteAppINI "fullFileName", "bookPath", bookPath
    WriteAppINI "fullFileName", "docxToTexPath", docxToTexPath
    WriteAppINI "fullFileName", "replaceSymbolListFile", replaceSymbolListFile
    WriteAppINI "fullFileName", "replaceListFile", replaceListFile
End Sub

Private Sub Text10_DblClick()
    Text10.SelStart = Text10.MaxLength
    Shell "notepad " + replaceSymbolListFile, vbNormalFocus
End Sub

Private Sub Text2_Change()
    If Text2.Text = "" Then
        questionID = 0
    Else
        questionID = CLng(Text2.Text)
    End If
End Sub

Private Sub Text3_Change()
    If Text3.Text = "" Then
        tabID = 0
    Else
        tabID = CLng(Text3.Text)
    End If
End Sub

Private Sub Text4_Change()
    If Text3.Text = "" Then
        figID = 0
    Else
        figID = CLng(Text4.Text)
    End If
End Sub

Private Sub Text5_DblClick()
    'beforeChange
    Text5.SelStart = Text5.MaxLength
    Shell "notepad " + bookPath, vbNormalFocus
End Sub

Private Sub Text5_Change()
    bookPath = Text5.Text
End Sub

Private Sub Text6_DblClick()
    Dim getdir As String

    getdir = BrowseForFolder(Me, "选择工作文件夹", Text6.Text)

    If Len(getdir) = 0 Then Exit Sub
    Text6.Text = IIf(Right$(getdir, 1) = "\", getdir, getdir & "\")

End Sub

Private Sub Text7_DblClick()
    Text7.SelStart = Text7.MaxLength
    Shell "notepad " + replaceListFile, vbNormalFocus

End Sub
'''''''''''''隐藏属性'''''''''''''''''
Function correcttest()
    Dim str As String
    Dim fr() As String
    Dim re As Object
    Dim mMatches As Object
    Dim docxFileName
    
    texFlagLabel.Caption = "处理中……"
    strFullName = selectFile(1)
    If strFullName(0) <> "" Then
'''''''''''''''''''''''''''''''''''''''''''
    
    Set re = New RegExp
    re.Global = True
    
    For Each docxFileName In strFullName
        readUTF8 docxFileName, str
        correctQlist str
        writeTex str, CStr(docxFileName)
    Next
''''''''''''''''''''''''''''''''''''''
    End If
    texFlagLabel.Caption = "完成！"
End Function
Function bl()
    Dim str As String
    Dim fr() As String
    Dim re As Object
    Dim mMatches As Object
    Dim docxFileName
    
    texFlagLabel.Caption = "处理中……"

    If findTexFile = False Then
        texFlagLabel.Caption = "遍历文件夹未设置！"
        Exit Function
    End If
    If strFullName(0) <> "" Then
'''''''''''''''''''''''''''''''''''''''''''
        Set re = New RegExp
        re.Global = True
        
        For Each docxFileName In strFullName
            readUTF8 docxFileName, str
            correctQlist str
            writeTex str, CStr(docxFileName)
        Next
''''''''''''''''''''''''''''''''''''''
        MsgBox "共 " & UBound(strFullName) & " 个文件"
    Else
        MsgBox "none"
    End If
    texFlagLabel.Caption = "完成"
End Function
Function correctQlist(ByRef s As String) As String
    Dim re As Object
    Dim mMatches As Object      '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    Dim str As String
    Dim strTemp As String
    Dim prev As Long
    
    prev = 1
    Set re = New RegExp
    re.Pattern = "\\item[XTJFM]\{"
    
    re.Global = True
    Set mMatches = re.Execute(s)
    For Each mMatch In mMatches
        If prev < mMatch.FirstIndex Then
            str = str + Mid(s, prev, mMatch.FirstIndex + 1 + mMatch.Length - prev)
            prev = nextRightBrace(mMatch.FirstIndex + mMatch.Length + 1, s)
            prev = nextRightBrace(prev + 2, s)
            strTemp = Mid(s, mMatch.FirstIndex + mMatch.Length + 1, prev - mMatch.FirstIndex - mMatch.Length - 1)
            strTemp = Replace(strTemp, "myitemize", "questionList")
            str = str + strTemp
            'prev = prev + 1
        End If
        DoEvents
    Next
    s = str + Mid(s, prev)
End Function
