VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 部门管理 
   Caption         =   "部门管理"
   ClientHeight    =   6770
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9350.001
   OleObjectBlob   =   "部门管理.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "部门管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
部门增加.show
End Sub

Private Sub CommandButton2_Click()
   Dim itm As ListItem
Dim deptID As Long
Dim found As Boolean

found = False
For Each itm In ListView1.ListItems
    If itm.Selected Then
        deptID = CLng(itm.text)
        found = True
        Exit For
    End If
Next itm

If Not found Then
    MsgBox "未选择部门", , "提示"
    Exit Sub
End If
    ' 调用部门修改窗体的初始化过程，填充对应数据
    部门修改.InitializeData deptID
    部门修改.show
End Sub
Private Sub CommandButton3_Click()
    Dim sh As Worksheet
    Dim itm As ListItem
    Dim sel As String
    Dim s As Variant
    Dim anySel As Boolean
    
    Set sh = Sheets("部门档案")
    anySel = False
    
    ' 遍历所有 ListItems，找第一个 Selected 的项
    For Each itm In ListView1.ListItems
        If itm.Selected Then
            anySel = True
            sel = itm.text
            Exit For
        End If
    Next itm
    
    If Not anySel Then
        ' 未选中：清空并允许手动输入
        With 部门删除
            .TextBox1.Enabled = True
            .TextBox1.BackColor = vbWhite
            .TextBox1.Value = ""     ' 清空
            .show
        End With
    Else
        ' 已选中：自动填充并禁止修改
        s = Application.Match(CLng(sel), sh.Range("A:A"), 0)
        If IsError(s) Then
            MsgBox "找不到对应部门编号！", , "提示"
            Exit Sub
        End If
        
        With 部门删除
            .TextBox1.Enabled = False
            .TextBox1.BackColor = &H8000000B
            .TextBox1.Value = sel
            .show
        End With
    End If
End Sub


Private Sub CommandButton4_Click()
Unload Me
系统平台.show
    
End Sub

Private Sub UserForm_Activate()
Call Flesh
End Sub

Public Sub Flesh()

ListView1.ColumnHeaders.Clear:
ListView1.ListItems.Clear '清空列表
Dim N, i As Integer
Set sh = Sheets("部门档案") '显示部门档案列表
Columnnumber = sh.Range("IV1").End(xlToLeft).Column
Rownumber = sh.Range("a65535").End(xlUp).Row
For i = 1 To 3
ListView1.ColumnHeaders.Add i, , sh.Cells(1, i), ListView1.Width / 3  '可设宽度
Next
ListView1.View = lvwReport: ListView1.FullRowSelect = True: ListView1.Gridlines = True
For i = 2 To Rownumber
Set itm = ListView1.ListItems.Add
itm.text = sh.Cells(i, 1)
For j = 2 To 3
itm.SubItems(j - 1) = sh.Cells(i, j)
Next
Next

If ListView1.ListItems.count > 0 Then ListView1.ListItems(1).Selected = False
End Sub

Public Sub Flesh2()

TreeView1.Nodes.Clear

Dim Nodx As Node
Dim N As Integer
Dim i, j As Integer

Set bm = Sheets("部门档案")


Set Nodx = TreeView1.Nodes.Add(, , "总公司", "总公司部门结构") '总公司是顶级节点的索引值,

    N = bm.Range("a65500").End(xlUp).Row
    For i = 2 To N
    If Len(bm.Cells(i, 1)) = 1 Then
        Set Nodx = TreeView1.Nodes.Add("总公司", tvwChild, , "(" & bm.Cells(i, 1) & ")" & bm.Cells(i, 2))
    ElseIf Len(bm.Cells(i, 1)) = 3 Then
        For j = 2 To N
            If Left(Trim(bm.Cells(i, 1)), 1) = Trim(bm.Cells(j, 1)) Then
                Set Nodx = TreeView1.Nodes.Add(FindNodeByText("(" & bm.Cells(j, 1) & ")" & bm.Cells(j, 2)), tvwChild, , "(" & bm.Cells(i, 1) & ")" & bm.Cells(i, 2))
            End If
         Next j
    End If

   Next i


For i = 1 To TreeView1.Nodes.count
     TreeView1.Nodes(i).Expanded = True
Next
End Sub
Private Sub UserForm_Initialize()
Dim Nodx As Node
Dim N As Integer
Dim i, j As Integer

Set bm = Sheets("部门档案")


Set Nodx = TreeView1.Nodes.Add(, , "总公司", "总公司部门结构") '总公司是顶级节点的索引值,

    N = bm.Range("a65500").End(xlUp).Row
    For i = 2 To N
    If Len(bm.Cells(i, 1)) = 1 Then
        Set Nodx = TreeView1.Nodes.Add("总公司", tvwChild, , "(" & bm.Cells(i, 1) & ")" & bm.Cells(i, 2))
    ElseIf Len(bm.Cells(i, 1)) = 3 Then
        For j = 2 To N
            If Left(Trim(bm.Cells(i, 1)), 1) = Trim(bm.Cells(j, 1)) Then
                Set Nodx = TreeView1.Nodes.Add(FindNodeByText("(" & bm.Cells(j, 1) & ")" & bm.Cells(j, 2)), tvwChild, , "(" & bm.Cells(i, 1) & ")" & bm.Cells(i, 2))
            End If
         Next j
    End If

   Next i


For i = 1 To TreeView1.Nodes.count
     TreeView1.Nodes(i).Expanded = True
Next


End Sub
Private Function FindNodeByText(text As String) As Node
    Dim Node As Node
    
    For Each Node In TreeView1.Nodes
        If Node.text = text Then
            Set FindNodeByText = Node
            Exit Function
        End If
    Next
    
    Set FindNodeByText = Nothing
End Function

