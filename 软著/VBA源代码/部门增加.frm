VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 部门增加 
   Caption         =   "增加部门"
   ClientHeight    =   5360
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4920
   OleObjectBlob   =   "部门增加.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "部门增加"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
 Set sh = Sheets("部门档案")
 
no = TextBox1.text
na = TextBox2.text
ps = TextBox3.text
 If Len(TextBox1.text) <> 1 And Len(TextBox1.text) <> 3 Then
        MsgBox "编码方式不合法！"
        Exit Sub
End If
If na = "" Or no = "" Then MsgBox "部门编号和部门名称为必填项", , "提示": Exit Sub

 ' 检查部门编号是否已存在
s = IsError(Application.Match(Val(no), sh.[A:A], 0))

If s = False Then MsgBox "部门已存在，请重新填写", , "提示": Exit Sub

' 检查部门名称是否已存在
    sna = Application.Match(na, sh.Range("B:B"), 0)
    If Not IsError(sna) Then
        MsgBox "部门名称已存在，请重新填写", vbExclamation, "提示"
        Exit Sub
    End If
 ' 找到最后一行
N = sh.[A65536].End(xlUp).Row
' 新增部门
sh.Cells(N + 1, 1) = no
sh.Cells(N + 1, 2) = na
sh.Cells(N + 1, 3) = ps
' 显示成功并关闭表单
MsgBox "新建部门成功", , "提示":
Unload Me

 ' 刷新部门管理表单
 Call 部门管理.Flesh
 Call 部门管理.Flesh2
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

