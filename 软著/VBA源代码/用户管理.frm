VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 用户管理 
   Caption         =   "用户管理"
   ClientHeight    =   4990
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9010.001
   OleObjectBlob   =   "用户管理.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "用户管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub TextBox1_Change()
    Dim sh As Worksheet
    Dim userNumber As String
    Dim userRow As Long
    Dim N As Long

    ' 设置工作表
    Set sh = Sheets("用户档案")

    ' 获取 TextBox1 中的人员编号
    userNumber = TextBox2.Value

    ' 如果人员编号为空，则清空其他文本框并退出
    If Len(userNumber) = 0 Then
        TextBox3.Value = ""
        TextBox4.Value = ""
        ComboBox1.Value = ""
        Exit Sub
    End If

    ' 查找人员编号在工作表中的位置
    On Error Resume Next
    userRow = WorksheetFunction.Match(userNumber, sh.Range("A:A"), 0)
    On Error GoTo 0

    ' 如果找到匹配项，则填充对应的文本框和组合框
    If Not IsError(userRow) Then
        用户修改.TextBox3.Value = sh.Cells(userRow, 2).Value
        用户修改.TextBox4.Value = sh.Cells(userRow, 3).Value
        用户修改.ComboBox1.Value = sh.Cells(userRow, 4).Value
    Else
        用户修改.TextBox3.Value = ""
        用户修改.TextBox4.Value = ""
        用户修改.ComboBox1.Value = ""
    End If
End Sub
Private Sub CommandButton1_Click()
用户增加.show
End Sub

Private Sub CommandButton2_Click()
 If (ListBox1.ListIndex = -1) Then 用户删除.show: Exit Sub
    For i = 0 To ListBox1.ListCount - 1   '遍历列表框选项
        If ListBox1.Selected(i) Then    '该选项是否被选择
        X = ListBox1.List(i)     '选中则拼接到字符串上
        End If
    Next
    用户删除.TextBox1.Enabled = False: 用户删除.TextBox1.BackColor = &H8000000B: 用户删除.TextBox1.text = X
    用户删除.show
End Sub

Private Sub CommandButton3_Click()
 N = Sheets("用户档案").Range("a65535").End(xlUp).Row
    If ListBox1.ListIndex = -1 Then
        用户修改.show
        Exit Sub
    Else
        Dim selectedUser As String
        selectedUser = ListBox1.Value

        ' 在用户信息中查找选中用户的信息
        For i = 3 To N
            If Sheets("用户档案").Range("a" & i).Value = selectedUser Then
                ' 填充选中用户的信息到用户修改界面
                用户修改.TextBox2.Enabled = False
                用户修改.TextBox2.BackColor = &H8000000B
                用户修改.TextBox2.text = selectedUser
                用户修改.TextBox3.text = Sheets("用户档案").Range("b" & i).Value
                用户修改.TextBox4.text = Sheets("用户档案").Range("c" & i).Value
                用户修改.ComboBox1.text = Sheets("用户档案").Range("d" & i).Value
                用户修改.show
                Exit Sub
            End If
        Next i
    End If
End Sub

Private Sub CommandButton4_Click()
 Unload Me
 系统平台.show
 
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub UserForm_Activate()
    Dim r As Long
    r = Sheets("用户档案").Range("A65536").End(xlUp).Row  ' 最后一行数据行号

    With ListBox1
        .ColumnCount = 4
        .ColumnHeads = True
        .RowSource = "用户档案!A2:D" & r  ' 表头在 A2:D2，这样 A2 就会作为列标题
    End With
End Sub

