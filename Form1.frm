VERSION 5.00
Object = "{0DF5D14C-08DD-4806-8BE2-B59CB924CFC9}#1.7#0"; "VBCCR16.OCX"
Begin VB.Form form_tex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "docxתtex������"
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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command15 
      Caption         =   "Join"
      Height          =   495
      Left            =   8880
      TabIndex        =   43
      Top             =   2040
      Width           =   615
   End
   Begin VB.Frame Frame8 
      Caption         =   "�״������ַ��б�"
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
         Caption         =   "ѡ��"
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
         Name            =   "����"
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
      Caption         =   "����ID"
      Height          =   495
      Left            =   3600
      TabIndex        =   37
      Top             =   5280
      Width           =   735
   End
   Begin VB.Frame Frame6 
      Caption         =   "docx2tex·������"
      Height          =   855
      Left            =   4680
      TabIndex        =   33
      Top             =   240
      Width           =   4815
      Begin VB.CommandButton Command12 
         Caption         =   "ѡ��"
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
      Caption         =   "���"
      Height          =   495
      Left            =   7800
      TabIndex        =   32
      Top             =   2040
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "��$"
      Height          =   180
      Left            =   3960
      TabIndex        =   31
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "����ѧ����"
      Height          =   495
      Left            =   4920
      TabIndex        =   30
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "����"
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
      Caption         =   "�����滻tex"
      Height          =   495
      Left            =   240
      TabIndex        =   19
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ID����"
      Height          =   2895
      Left            =   2040
      TabIndex        =   8
      Top             =   840
      Width           =   2415
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "ʹ����ʱID"
         BeginProperty Font 
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
      Caption         =   "ת��+����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "δʶ����Ŀ��"
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
      Caption         =   "tex����"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "docxת��tex"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "״̬"
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
      Caption         =   "tex�ļ��Զ����滻"
      Height          =   2895
      Left            =   120
      TabIndex        =   20
      Top             =   5040
      Width           =   4335
      Begin VB.CommandButton Command8 
         Caption         =   "����ID"
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
         Caption         =   "ѡ���滻tex"
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
         Text            =   "˫��ѡ���ļ���"
         Top             =   840
         Width           =   3015
      End
      Begin VB.Frame Frame5 
         Caption         =   "�滻�����б��ļ�"
         Height          =   1335
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   4095
         Begin VB.CheckBox Check3 
            Caption         =   "����\left\right"
            BeginProperty Font 
               Name            =   "����"
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
         Caption         =   "��  ���ļ���"
         Height          =   385
         Left            =   360
         TabIndex        =   25
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "SourceDB.ini ·��"
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
    
    texFlagLabel.Caption = "�����С���"
    
    If setID Then
        tt = Main
        IDFresh
        If Unidentified <> 0 Then
            numLabel.Caption = CStr(Unidentified)
        End If
        texFlagLabel.Caption = "��ɣ�" + Chr(13) + tt
    Else
        MsgBox "ID���ô���"
        texFlagLabel.Caption = "ID���ô���"
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

    getdir = BrowseForFolder(Me, "ѡ�����ļ���", Text9.Text)

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
    texFlagLabel.Caption = "�����С���"
    Ctrl = GetKeyState(vbKeyControl)
    If Ctrl < 0 Then
        'MsgBox "Ctrl+����"
        RichTextBox1.Text = JoinTest(True)
    Else
        'MsgBox "����"
        RichTextBox1.Text = JoinTest
    End If
    texFlagLabel.Caption = "��ɣ�"
End Sub

Private Sub Command2_Click()
    Dim tt As String
    texFlagLabel.Caption = "�����С���"
    If setID Then
        tt = changeToTex
    Else
        MsgBox "ID���ô���"
    End If
    texFlagLabel.Caption = "��ɣ�" + Chr(13) + tt
End Sub

Private Sub Command3_Click()
    texFlagLabel.Caption = "�����С���"
    If setID Then
        convertToTex
        IDFresh
    Else
        MsgBox "ID���ô���"
    End If
    fileSelect = False
    texFlagLabel.Caption = "��ɣ�"
End Sub

Private Sub Command4_Click()
    texFlagLabel.Caption = "�����С���"

    If findTexFile = False Then
        texFlagLabel.Caption = "�����ļ���δ���ã�"
        Exit Sub
    End If
    If readReplaceList = False Then
        texFlagLabel.Caption = "�˳�"
        Exit Sub
    End If
    If strFullName(0) <> "" Then
        replaceMacros
        MsgBox "�� " & UBound(strFullName) & " ���ļ�"
    Else
        MsgBox "none"
    End If
    texFlagLabel.Caption = "���"
End Sub
Function findTexFile() As Boolean
    Dim str1() As String
    Dim a As Long, b As Long, c As Long
    Dim strg As String, fz As String
    Dim strTemp As String
    
    findTexFile = False
    If Text6 = "" Then Exit Function
        fz = Text6 '�ļ���"d:\tl\" '
        On Error GoTo XH1
XH1:
        Resume Next
        strg = Dir(fz, 63)
XH:
    If strg <> "" Then
        If (GetAttr(fz & strg) <> 9238) And (GetAttr(fz & strg) And 16) = 16 And Right(strg, 1) <> "." Then
            a = a + 1            'Label3.Caption = a
        ReDim Preserve str1(a)   '�����ϴ�������������Ŀ¼
            str1(a) = fz & strg & "\"    '�����ϴ�������������Ŀ¼
          'List1.AddItem str1(a)   '��ʾ������Ŀ¼
        ElseIf Right(strg, 1) <> "." Then '
          If Right(strg, 4) = ".tex" Then '
            c = c + 1            'Label4.Caption = c
            strTemp = strTemp + "?" + fz & strg
          End If
          'List2.AddItem fz & strg  '��ʾ���������ļ�
        End If
        strg = Dir           '������Ȼ������XH:
        'Label3.Caption = a   'Ŀ¼��
        'Label4.Caption = c   '�ļ���
        DoEvents             '���������¼�������
        GoTo XH                  '����XH:����Ƿ�Ϊ����תElseIf
    ElseIf strg = "" And a <> 0 Then  'strg=""˵�����ϴ�����ɣ�a<>0�Ǽ�¼Ŀ¼��Ϊ�գ�������������
        b = b + 1            '���ϴ��������ɣ��鿴�Ƿ�������������Ŀ¼�㣬+1��һ����������b>a��⡣
        If b > a Then
            If c > 0 Then
                strFullName = Split(Right(strTemp, Len(strTemp) - 1), "?")
                findTexFile = True
            End If
            Exit Function      '��¼���ļ��г�����֤��ȫ�ֱ�����ɡ�
        End If
        fz = str1(b)         '��ֵ������Ŀ¼��
        GoTo XH1                 '����XH1:��Ǵ�ִ�д���������ֵ����
    End If
    If c > 0 Then
        strFullName = Split(Right(strTemp, Len(strTemp) - 1), "?")
        findTexFile = True
    End If
End Function

Private Sub Command5_Click()
    texFlagLabel.Caption = "�����С���"
    If readReplaceList = False Then Exit Sub
    strFullName = selectFile(1)
    If strFullName(0) <> "" Then
        replaceMacros
    End If
    texFlagLabel.Caption = "��ɣ�"
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
    'getdir = BrowseForFolder(Me, "ѡ�����ļ���", Text5.Text)

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
            MsgBox "INI���ô���"
        End If
    End If
End Sub

Private Sub Command8_Click()
    texFlagLabel.Caption = "��ʼ��"
    Redistribution
    IDFresh
    texFlagLabel.Caption = "��ɣ�"
End Sub

Private Sub Command9_Click()
Clipboard.Clear
Clipboard.SetText (RichTextBox1.Text)
End Sub

Private Sub Form_DblClick()
    texFlagLabel.Caption = "��ʼ��"
    'onlyToTex
    'beforeChange
    'correcttest
    'MsgBox "questionID:" & questionID & " FigID:" & figID & " TabID:" & tabID
    MsgBox nextRightBrace(1, "$M=\left\{x| \dfrac{x}{x-1}\leqslant 0\right\},N=\{.x| x^{2}-2x<0\}$��")
    MsgBox "$" & TrimEnter(Chr(13) + Chr(13) + "start" + Chr(13), "R") & "$"
    texFlagLabel.Caption = "��ɣ�"
End Sub

Function redistributeAll()
    texFlagLabel.Caption = "�����С���"

    If findTexFile = False Then
        texFlagLabel.Caption = "�����ļ���δ���û�δ�ҵ�tex�ļ���"
        Exit Function
    End If
    If strFullName(0) <> "" Then
        fileSelect = True
        Redistribution
        MsgBox "�� " & UBound(strFullName) & " ���ļ�"
        IDFresh
        texFlagLabel.Caption = "���"
    Else
        MsgBox "�����ñ����ļ��У�"
    End If
End Function
Sub test()
    Dim re As Object
    Dim str As String
    Dim s As String
    Dim prev As Long
    Dim mMatches As Object      'ƥ���ַ������϶���
    Dim mMatch As Object        'ƥ���ַ���
    Set re = New RegExp
    re.Global = True
    re.Pattern = "\\item[XTJFM]\{"
    prev = 1
    s = "dfs\itemX{1}{�ڵ���$\bigtriangleup ABC$��$AB=AC$��$BC=6$������$\vec{AD}=\vec{DC}$����$\vec{DC}\cdot \vec{BC}$��ֵΪ\xz }{$9$}{$18$}{$27$}{$36$}{$A$}"
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
    
    str = GetAppINI("fullFileName", "replaceSymbolListFile")    '�����滻�б�
    If str <> "" Then
        replaceSymbolListFile = str
    Else
        replaceSymbolListFile = App.Path & "\replaceSymbolList.txt"
    End If
    
    str = GetAppINI("fullFileName", "replaceListFile")          '���������б�
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
        questionTypeXZ = "ѡ����"
    End If
    
    str = GetAppINI("fullFileName", "questionTypeTK")
    str = Replace(str, Chr(0), "")
    If str <> "" Then
        questionTypeTK = str
    Else
        questionTypeTK = "�����"
    End If
    
    str = GetAppINI("fullFileName", "questionTypeJD")
    str = Replace(str, Chr(0), "")
    If str <> "" Then
        questionTypeJD = str
    Else
        questionTypeJD = "�����"
    End If
    
    str = GetAppINI("fullFileName", "questionAnswerBoundary")
    str = Replace(str, Chr(0), "")
    If str <> "" Then
        questionAnswerBoundary = str
    Else
        questionAnswerBoundary = "��������"
    End If
    
    str = GetAppINI("fullFileName", "answerBoundary")
    str = Replace(str, Chr(0), "")
    If str <> "" Then
        answerBoundary = str
    Else
        answerBoundary = "�ο���"
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
    '���м��������ȡ
    str = GetAppINI("answer", "AnswerSolutionBoundary")
    str = Replace(str, Chr(0), "")
    If str <> "" Then
        answerSolutionBoundary = str
    Else
        answerSolutionBoundary = "[��]?����[��]?"
    End If
    str = GetAppINI("answer", "SolutionStart")
    str = Replace(str, Chr(0), "")
    If str <> "" Then
        solutionStart = str
    Else
        solutionStart = "[��]?���[��]?"
    End If
    str = GetAppINI("answer", "SolutionEnd")
    str = Replace(str, Chr(0), "")
    If str <> "" Then
        solutionEnd = str
    Else
        solutionEnd = "[��]?�㾦[��]?"
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
            MsgBox "�����ļ���ȡ����"
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

    getdir = BrowseForFolder(Me, "ѡ�����ļ���", Text6.Text)

    If Len(getdir) = 0 Then Exit Sub
    Text6.Text = IIf(Right$(getdir, 1) = "\", getdir, getdir & "\")

End Sub

Private Sub Text7_DblClick()
    Text7.SelStart = Text7.MaxLength
    Shell "notepad " + replaceListFile, vbNormalFocus

End Sub
'''''''''''''��������'''''''''''''''''
Function correcttest()
    Dim str As String
    Dim fr() As String
    Dim re As Object
    Dim mMatches As Object
    Dim docxFileName
    
    texFlagLabel.Caption = "�����С���"
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
    texFlagLabel.Caption = "��ɣ�"
End Function
Function bl()
    Dim str As String
    Dim fr() As String
    Dim re As Object
    Dim mMatches As Object
    Dim docxFileName
    
    texFlagLabel.Caption = "�����С���"

    If findTexFile = False Then
        texFlagLabel.Caption = "�����ļ���δ���ã�"
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
        MsgBox "�� " & UBound(strFullName) & " ���ļ�"
    Else
        MsgBox "none"
    End If
    texFlagLabel.Caption = "���"
End Function
Function correctQlist(ByRef s As String) As String
    Dim re As Object
    Dim mMatches As Object      'ƥ���ַ������϶���
    Dim mMatch As Object        'ƥ���ַ���
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
