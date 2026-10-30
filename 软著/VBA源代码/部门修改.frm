VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 部门修改 
   Caption         =   "修改部门"
   ClientHeight    =   5130
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4940
   OleObjectBlob   =   "部门修改.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "部门修改"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 初始化并加载部门数据
Public Sub InitializeData(ByVal deptID As Long)
    Dim sh As Worksheet
    Dim rowIndex As Variant
    
    Set sh = ThisWorkbook.Sheets("部门档案")
    
    rowIndex = Application.Match(deptID, sh.Range("A:A"), 0)
    
    If IsError(rowIndex) Then
        MsgBox "部门编号不存在！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 填充文本框
    TextBox1.Value = sh.Cells(rowIndex, 1).Value   ' 部门编号
    TextBox2.Value = sh.Cells(rowIndex, 2).Value   ' 部门名称
    TextBox3.Value = sh.Cells(rowIndex, 3).Value   ' 备注
    
    ' 编号不可修改
    TextBox1.Enabled = False
    TextBox1.BackColor = &H8000000B ' 灰色背景
End Sub

' 保存按钮点击事件，更新部门信息
Private Sub CommandButton1_Click()
    Dim no As String
    Dim na As String
    Dim ps As String
    Dim result As Variant
    Dim s As Long
    Dim sh As Worksheet

    Set sh = ThisWorkbook.Sheets("部门档案")

    no = Trim(TextBox1.text)
    na = Trim(TextBox2.text)
    ps = Trim(TextBox3.text)

    If na = "" Then
        MsgBox "未输入部门名称", , "提示"
        Exit Sub
    End If

    result = Application.Match(CLng(no), sh.Range("A:A"), 0)

    If IsError(result) Then
        MsgBox "部门编号不存在", , "提示"
        Exit Sub
    Else
        s = result
        sh.Cells(s, 2).Value = na
        sh.Cells(s, 3).Value = ps

        MsgBox "修改成功", , "提示"
        Unload Me
        Call 部门管理.Flesh
        Call 部门管理.Flesh2
    End If
End Sub

' 关闭按钮点击事件
Private Sub CommandButton2_Click()
    Unload Me
End Sub
