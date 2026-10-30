VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 人员管理 
   Caption         =   "人员管理"
   ClientHeight    =   7050
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   11560
   OleObjectBlob   =   "人员管理.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "人员管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 声明模块级变量保存当前选中人员编号
Private currentSelectedID As String

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
    Cancel = True ' 禁止直接编辑
End Sub

Private Sub UserForm_Initialize()
    InitTreeView
    InitListView
    FillListView
    currentSelectedID = ""  ' 一定要清空
End Sub

Private Sub UserForm_Activate()
    InitTreeView
    InitListView
    FillListView
    currentSelectedID = ""  ' 一定要清空
End Sub

' —— 初始化 TreeView ——
Public Sub InitTreeView()
    Dim bm As Worksheet
    Dim nod As MSComctlLib.Node
    Dim i As Long, j As Long
    Dim lastRow As Long
    
    Set bm = ThisWorkbook.Sheets("部门档案")
    lastRow = bm.Cells(bm.Rows.count, "A").End(xlUp).Row
    
    With Me.TreeView1
        .Nodes.Clear
        Set nod = .Nodes.Add(, , "root", "总公司人事结构")
        For i = 2 To lastRow
            If Len(Trim(bm.Cells(i, "A").Value)) = 1 Then
                .Nodes.Add "root", tvwChild, "dept" & i, bm.Cells(i, "B").Value
            ElseIf Len(Trim(bm.Cells(i, "A").Value)) = 3 Then
                For j = 2 To lastRow
                    If Left(Trim(bm.Cells(i, "A").Value), 1) = Trim(bm.Cells(j, "A").Value) Then
                        Set nod = FindNodeByText(bm.Cells(j, "B").Value)
                        If Not nod Is Nothing Then
                            .Nodes.Add nod.key, tvwChild, "dept" & i, bm.Cells(i, "B").Value
                        End If
                        Exit For
                    End If
                Next j
            End If
        Next i
        
        Dim k As Long
        For k = 1 To .Nodes.count
            .Nodes(k).Expanded = True
        Next k
    End With
End Sub

' —— 初始化 ListView 列头 ——
Private Sub InitListView()
    With Me.ListView1
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .ColumnHeaders.Clear
        .ListItems.Clear
        .ColumnHeaders.Add 1, "c1", "人员编号", .Width / 7
        .ColumnHeaders.Add 2, "c2", "人员姓名", .Width / 7, lvwColumnCenter
        .ColumnHeaders.Add 3, "c3", "性别", .Width / 7, lvwColumnCenter
        .ColumnHeaders.Add 4, "c4", "部门总称", .Width / 7, lvwColumnCenter
        .ColumnHeaders.Add 5, "c5", "行政部门", .Width / 7, lvwColumnCenter
        .ColumnHeaders.Add 6, "c6", "雇佣状态", .Width / 7, lvwColumnCenter
        .ColumnHeaders.Add 7, "c7", "是否业务员", .Width / 7, lvwColumnCenter
    End With
End Sub

' —— 填充 ListView 内容 ——
Public Sub FillListView()
    Dim rs As Worksheet
    Dim li As ListItem
    Dim i As Long
    Dim lastRow As Long
    Dim deptSel As String

    Set rs = ThisWorkbook.Sheets("人员档案")
    lastRow = rs.Cells(rs.Rows.count, "A").End(xlUp).Row

    ListView1.ListItems.Clear

    If Not TreeView1.SelectedItem Is Nothing Then
        deptSel = TreeView1.SelectedItem.text
    Else
        deptSel = ""
    End If

    For i = 2 To lastRow
        If deptSel = "" _
        Or rs.Cells(i, "D").Value = deptSel _
        Or rs.Cells(i, "E").Value = deptSel _
        Or deptSel = "总公司人事结构" Then

            Set li = ListView1.ListItems.Add(, , rs.Cells(i, "A").Value)
            li.SubItems(1) = rs.Cells(i, "B").Value
            li.SubItems(2) = rs.Cells(i, "C").Value
            li.SubItems(3) = rs.Cells(i, "D").Value
            li.SubItems(4) = rs.Cells(i, "E").Value
            li.SubItems(5) = rs.Cells(i, "F").Value
            li.SubItems(6) = rs.Cells(i, "G").Value
        End If
    Next i

    currentSelectedID = "" ' 清空选中
End Sub

' —— 点击树节点刷新 ListView ——
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    InitListView
    FillListView
    currentSelectedID = ""
End Sub

' —— ListView 点击时更新选中 ID ——
Private Sub ListView1_Click()
    Dim itm As ListItem
    currentSelectedID = ""
    For Each itm In ListView1.ListItems
        If itm.Selected Then
            currentSelectedID = itm.text
            Exit For
        End If
    Next
End Sub

' —— “增加人员” 按钮 ——
Private Sub CommandButton1_Click()
    人员增加.show
End Sub

' —— “修改人员” 按钮 ——
Private Sub CommandButton2_Click()
    If currentSelectedID = "" Then
        MsgBox "请选择一个人员进行修改！", vbExclamation, "提示"
        Exit Sub
    End If
    
    人员修改.InitializeData currentSelectedID
    人员修改.show
End Sub

' —— “删除人员” 按钮 ——
Private Sub CommandButton3_Click()
    If currentSelectedID = "" Then
        MsgBox "请选择一个人员进行删除！", vbExclamation, "提示"
        Exit Sub
    End If
    
    With 人员删除
        .TextBox1.Enabled = False
        .TextBox1.BackColor = &H8000000B
        .TextBox1.Value = currentSelectedID
        .show
    End With
End Sub

' —— “关闭” 按钮 ——
Private Sub CommandButton4_Click()
    Me.Hide
    系统平台.show
End Sub

' —— 查找 TreeView 节点 的辅助函数 ——
Private Function FindNodeByText(text As String) As MSComctlLib.Node
    Dim nd As MSComctlLib.Node
    For Each nd In Me.TreeView1.Nodes
        If nd.text = text Then
            Set FindNodeByText = nd
            Exit Function
        End If
    Next
    Set FindNodeByText = Nothing
End Function


