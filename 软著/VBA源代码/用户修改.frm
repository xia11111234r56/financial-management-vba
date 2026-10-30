VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 用户修改 
   Caption         =   "修改用户"
   ClientHeight    =   6230
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5870
   OleObjectBlob   =   "用户修改.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "用户修改"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Private Sub CommandButton1_Click()
 Dim sig As Integer
    Dim N As Integer
    Dim i As Integer
    Dim num As String
    sig = 0 '列表中是否已有账套主管
    N = Sheets("用户档案").Range("a65535").End(xlUp).Row
    num = TextBox2.Value
     For i = 3 To N
        If Sheets("用户档案").Range("a" & i) = num Then
            a = i
        End If
        If Sheets("用户档案").Range("c" & i) = "账套主管" Then
        MsgBox "已存在账套主管，请重新输入！", , "提示"
        End If
    Next i
    If Len(TextBox3.Value) <> 0 Then
        Sheets("用户档案").Range("b" & a) = TextBox3.Value
    End If

    If Len(TextBox4.Value) <> 0 Then
        Sheets("用户档案").Range("c" & a) = TextBox4.Value
    End If
   
    If Len(ComboBox1.Value) <> 0 Then
        Sheets("用户档案").Range("d" & a) = ComboBox1.Value
    End If
    MsgBox "修改成功", , "提示"
    Unload Me
    用户管理.Hide
    用户管理.show
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub


Private Sub UserForm_Activate()
    ComboBox1.AddItem "账套主管"
    ComboBox1.AddItem "会计人员"
    ComboBox1.AddItem "出纳人员"
    ComboBox1.AddItem "普通用户"
End Sub


