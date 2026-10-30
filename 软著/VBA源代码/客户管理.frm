VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 客户管理 
   Caption         =   "客户管理"
   ClientHeight    =   6540
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   10490
   OleObjectBlob   =   "客户管理.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "客户管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 在窗体顶部声明模块级变量
Private currentSelectedID As String

' —— 新增：窗体初始化时，确保 ListView 单选，并无任何默认选中 ——
Private Sub UserForm_Initialize()
    With Me.ListView1
        .MultiSelect = False
    End With
End Sub

' —— 增加客户按钮 ——
Private Sub CommandButton1_Click()
    客户增加.show
End Sub

' —— 修改客户按钮 ——
Private Sub CommandButton2_Click()
    If currentSelectedID = "" Then
        MsgBox "请选择一个客户进行修改！", vbExclamation, "提示"
        Exit Sub
    End If
    客户修改.InitializeData currentSelectedID
    客户修改.show
End Sub

' —— 删除客户按钮 ——
Private Sub CommandButton3_Click()
    If currentSelectedID = "" Then
        MsgBox "请选择一个客户进行删除！", vbExclamation, "提示"
        Exit Sub
    End If
    With 客户删除
        .TextBox1.Value = currentSelectedID
        .TextBox1.Enabled = False
        .show
    End With
End Sub

' —— 返回系统平台 ——
Private Sub CommandButton4_Click()
    Me.Hide
    系统平台.show
End Sub

' —— 禁止在列表中直接编辑 ——
Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

' —— 用户点击 ListView 时，更新 currentSelectedID ——
Private Sub ListView1_Click()
    Dim itm As ListItem
    currentSelectedID = ""
    For Each itm In Me.ListView1.ListItems
        If itm.Selected Then
            currentSelectedID = itm.text
            Exit For
        End If
    Next
End Sub

' —— 窗体激活时刷新列表 ——
Private Sub UserForm_Activate()
    RefreshListView
End Sub

' —— 刷新 ListView1 内容 ——
Public Sub RefreshListView()
    Dim ws As Worksheet
    Dim r As Long, i As Long
    Dim itm As ListItem
    
    Set ws = ThisWorkbook.Sheets("客户档案")
    r = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    With Me.ListView1
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .MultiSelect = False          ' 确保单选
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        ' 添加列头
        .ColumnHeaders.Add , , "客户编号", 60
        .ColumnHeaders.Add , , "客户名称", 100
        .ColumnHeaders.Add , , "客户简称", 60
        .ColumnHeaders.Add , , "分管部门", 60
        .ColumnHeaders.Add , , "分管业务员", 80
        .ColumnHeaders.Add , , "联系方式", 80
        
        ' 填充数据
        For i = 2 To r
            If ws.Cells(i, 1).Value <> "" Then
                Set itm = .ListItems.Add(, , ws.Cells(i, 1).Value)
                itm.SubItems(1) = ws.Cells(i, 2).Value
                itm.SubItems(2) = ws.Cells(i, 3).Value
                itm.SubItems(3) = ws.Cells(i, 4).Value
                itm.SubItems(4) = ws.Cells(i, 5).Value
                itm.SubItems(5) = ws.Cells(i, 6).Value
            End If
        Next i
        
        ' —— 清除任何默认选中 ——
        For Each itm In .ListItems
            itm.Selected = False
        Next itm
    End With
    
    currentSelectedID = ""  ' 刷新后清空选中编号
End Sub

