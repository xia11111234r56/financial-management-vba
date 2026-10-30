VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 供应商管理 
   Caption         =   "供应商管理"
   ClientHeight    =   7070
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   10660
   OleObjectBlob   =   "供应商管理.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "供应商管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' —— 窗体初始化时设置 ListView 属性 ——
Private Sub UserForm_Initialize()
    With Me.ListView1
        .MultiSelect = False    ' 单选
        .HideSelection = False
    End With
End Sub

' —— 窗体激活时刷新列表并清除选中 ——
Private Sub UserForm_Activate()
    RefreshListView
End Sub

' —— 窗体显示后也清除选中 ——
Private Sub UserForm_Show()
    ClearSelection
End Sub

' —— 刷新 ListView1 内容 ——
Public Sub RefreshListView()
    Dim ws As Worksheet
    Dim r As Long, i As Long
    Dim itm As ListItem
    
    Set ws = ThisWorkbook.Sheets("供应商档案")
    r = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    With Me.ListView1
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .MultiSelect = False
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        .ColumnHeaders.Add , , "编号", 60
        .ColumnHeaders.Add , , "供应商名称", 120
        .ColumnHeaders.Add , , "供应物品", 100
        .ColumnHeaders.Add , , "分管业务员", 100
        
        For i = 2 To r
            If ws.Cells(i, 1).Value <> "" Then
                Set itm = .ListItems.Add(, , ws.Cells(i, 1).Value)
                itm.SubItems(1) = ws.Cells(i, 2).Value
                itm.SubItems(2) = ws.Cells(i, 3).Value
                itm.SubItems(3) = ws.Cells(i, 4).Value
            End If
        Next i
        
        ' 清除所有选中状态
        ClearSelection
    End With
    
    ' 将焦点移到“增加供应商”按钮，防止 ListView 自动选中
    Me.CommandButton1.SetFocus
End Sub

' —— 辅助：清除所有选中 ——
Private Sub ClearSelection()
    Dim itm As ListItem
    On Error Resume Next
    For Each itm In Me.ListView1.ListItems
        itm.Selected = False
    Next itm
    On Error GoTo 0
End Sub

' —— 增加供应商 ——
Private Sub CommandButton1_Click()
    供应商增加.show
End Sub

' —— 修改供应商 ——
Private Sub CommandButton2_Click()
    Dim itm As ListItem
    Dim id As Long
    Dim found As Boolean

    found = False
    ' 遍历 ListItems，只有真正选中时才赋值
    For Each itm In Me.ListView1.ListItems
        If itm.Selected Then
            id = CLng(itm.text)
            found = True
            Exit For
        End If
    Next itm
    
    If Not found Then
        MsgBox "请选择一个供应商进行修改！", vbExclamation, "提示"
        Exit Sub
    End If
    
    供应商修改.InitializeData id
    供应商修改.show
End Sub

' —— 删除供应商 ——
Private Sub CommandButton3_Click()
    Dim itm As ListItem, id As Long
    For Each itm In Me.ListView1.ListItems
        If itm.Selected Then
            id = CLng(itm.text)
            With 供应商删除
                .TextBox1.Value = id
                .TextBox1.Enabled = False
                .show
            End With
            Exit Sub
        End If
    Next
    MsgBox "请选择一个供应商进行删除！", vbExclamation, "提示"
End Sub

' —— 返回/关闭 ——
Private Sub CommandButton4_Click()
    Me.Hide
    系统平台.show
End Sub

