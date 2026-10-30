VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 供应商增加 
   Caption         =   "增加供应商"
   ClientHeight    =   5880
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5490
   OleObjectBlob   =   "供应商增加.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "供应商增加"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 每次窗体显示前清空上次数据
Private Sub UserForm_Activate()
    Me.TextBox1.Value = ""
    Me.TextBox2.Value = ""
    Me.TextBox3.Value = ""
    Me.ComboBox1.Value = ""
End Sub

' 窗体初始化：加载分管业务员列表（首次载入）
Private Sub UserForm_Initialize()
    Dim wsP As Worksheet
    Dim lastRow As Long, i As Long
    Set wsP = ThisWorkbook.Sheets("人员档案")
    lastRow = wsP.Cells(wsP.Rows.count, "A").End(xlUp).Row
    
    Me.ComboBox1.Clear
    For i = 2 To lastRow
        If Trim(wsP.Cells(i, "G").Value) = "是" Then
            Me.ComboBox1.AddItem wsP.Cells(i, "B").Value
        End If
    Next i
End Sub

' 确认添加
Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim no As Long, na As String, Item As String, salesman As String
    
    Set ws = ThisWorkbook.Sheets("供应商档案")
    no = Val(Me.TextBox1.Value)
    na = Trim(Me.TextBox2.Value)
    Item = Trim(Me.TextBox3.Value)
    salesman = Trim(Me.ComboBox1.Value)
    
    ' 校验
    If no = 0 Or na = "" Or Item = "" Then
        MsgBox "编号、名称、物品不能为空！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 查重
    If Not IsError(Application.Match(no, ws.Range("A:A"), 0)) Then
        MsgBox "该编号已存在！", vbExclamation, "提示"
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row + 1
    ws.Cells(lastRow, 1).Value = no
    ws.Cells(lastRow, 2).Value = na
    ws.Cells(lastRow, 3).Value = Item
    ws.Cells(lastRow, 4).Value = salesman
    
    MsgBox "添加成功！", vbInformation, "提示"
    Me.Hide
    供应商管理.RefreshListView
End Sub

' 取消
Private Sub CommandButton2_Click()
    Me.Hide
End Sub

