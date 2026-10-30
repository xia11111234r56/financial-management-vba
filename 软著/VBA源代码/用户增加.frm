VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 用户增加 
   Caption         =   "增加用户"
   ClientHeight    =   5580
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5330
   OleObjectBlob   =   "用户增加.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "用户增加"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim N As Integer
    N = Sheets("用户档案").Range("a65535").End(xlUp).Row
    Dim i As Integer
    Dim no As Integer
    Dim name As String
    Dim pass As String
    Dim role As String
    If Len(TextBox1.Value) = 0 Or Len(TextBox2.Value) = 0 Or Len(TextBox3.Value) = 0 Or Len(ComboBox1.Value) = 0 Then
        MsgBox "信息不完整！", , "提示"
        Exit Sub
    End If
    no = TextBox1.Value
    name = TextBox2.Value
    pass = TextBox3.Value
    role = ComboBox1.Value
    For i = 2 To N
        If ComboBox1.Value = "账套主管" Then
            MsgBox "已有账套主管，请重新选择类型！", , "提示"
            Exit Sub
        End If
        If Sheets("用户档案").Range("a" & i) = no Then
            MsgBox "该编号已存在，请重新输入！", , "提示"
            Exit Sub
        End If
    Next i
    Sheets("用户档案").Range("a" & N + 1) = no
    Sheets("用户档案").Range("b" & N + 1) = name
    Sheets("用户档案").Range("c" & N + 1) = pass
    Sheets("用户档案").Range("d" & N + 1) = role
    
      ' 排序数据
    With Sheets("用户档案").Sort
        .SortFields.Clear
        .SortFields.Add key:=Range("A3:A" & N + 1), Order:=xlAscending
        .SetRange Range("A3:D" & N + 1)
        .Apply
    End With
    MsgBox "增加成功！", , "提示"
    Unload Me
    用户管理.Hide
    用户管理.show
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub


Private Sub UserForm_Activate()

    ComboBox1.AddItem "管理员"
    ComboBox1.AddItem "账套主管"
    ComboBox1.AddItem "会计人员"
    ComboBox1.AddItem "出纳人员"
    ComboBox1.AddItem "普通用户"
End Sub

