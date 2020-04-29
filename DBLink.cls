VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Node"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private freeList As Collection '局部复制
Private doubleLink As Collection '局部复制
Private first As Long
Private last As Long
Public finished As Boolean

Public Function creatNode(ByVal str As String, Optional ByVal flag As Boolean = False) As Long
    Dim i As Long
    Dim aNode As Node
    If freeList.Count = 0 Then
        Set aNode = New Node
        aNode.flag = flag
        doubleLink.add aNode
        i = doubleLink.Count
        doubleLink.Item(i).str = str
        creatNode = i
    Else
        i = freeList.Item(1)
        doubleLink.Item(i).str = str
        freeList.Remove (1)
        creatNode = i
    End If
End Function
Public Function printDBLink() As String
    Dim s As String
    Dim i As Long
    i = getFirstIndex()
    Do While i <> 0
        s = s + getNode(i)
        i = getNextIndex(i)
    Loop
    printDBLink = s
End Function

Public Function getFirstIndex() As Long
    getFirstIndex = first
End Function

Public Function getNextIndex(ByVal index As Long) As Long
    getNextIndex = doubleLink.Item(index).nextIndex
End Function

Public Function test(ByVal index As Long)
    freeList.add index
End Function
Public Function linkTwo(ByRef i As Long, ByVal j As Long)
    doubleLink.Item(i).nextIndex = j
    doubleLink.Item(j).prevIndex = i
    i = j
End Function
Public Function getNode(ByVal index As Long) As String
    getNode = doubleLink.Item(index).str
End Function

Public Function getNodeFlag(ByVal index As Long) As Boolean
    getNodeFlag = doubleLink.Item(index).flag
End Function

Public Function addNode(ByVal index As Long)
    If first = 0 Then
        first = index
        last = index
    Else
        doubleLink.Item(last).nextIndex = index
        doubleLink.Item(index).prevIndex = last
        last = index
    End If
End Function
Public Function setNodeFlag(ByVal index As Long, ByVal flag As Boolean)
    doubleLink.Item(index).flag = flag
End Function

Public Function replaceNode(ByVal index As Long, ByVal insertSIndex As Long, Optional ByVal insertEIndex As Long = 0)
    If index = 0 Then
        first = insertSIndex
        last = insertEIndex
        Exit Function
    End If
    If insertEIndex = 0 Then insertEIndex = insertSIndex
    If doubleLink.Item(index).prevIndex = 0 Then        '替换首个
        doubleLink.Item(insertEIndex).nextIndex = doubleLink.Item(index).nextIndex
        doubleLink.Item(doubleLink.Item(index).nextIndex).prevIndex = insertEIndex
    ElseIf doubleLink.Item(index).nextIndex = 0 Then    '替换末尾
        doubleLink.Item(insertSIndex).prevIndex = doubleLink.Item(index).prevIndex
        doubleLink.Item(doubleLink.Item(index).prevIndex).nextIndex = insertSIndex
    Else
        doubleLink.Item(insertEIndex).nextIndex = doubleLink.Item(index).nextIndex
        doubleLink.Item(doubleLink.Item(index).nextIndex).prevIndex = insertEIndex
        
        doubleLink.Item(insertSIndex).prevIndex = doubleLink.Item(index).prevIndex
        doubleLink.Item(doubleLink.Item(index).prevIndex).nextIndex = insertSIndex
    End If
    freeList.add index
    If doubleLink.Item(insertEIndex).nextIndex = 0 Then last = insertEIndex
    If doubleLink.Item(insertSIndex).prevIndex = 0 Then first = insertSIndex
End Function

Private Sub Class_Initialize()
    '当创建 DBLink 类时，创建 mNode 对象
    Set doubleLink = New Collection
    Set freeList = New Collection
    first = 0
    last = 0
    finished = False
End Sub
