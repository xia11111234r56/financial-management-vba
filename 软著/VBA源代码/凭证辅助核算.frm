VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 凭证辅助核算 
   Caption         =   "辅助核算"
   ClientHeight    =   6330
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6460
   OleObjectBlob   =   "凭证辅助核算.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "凭证辅助核算"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
If title.Caption = "客户往来" Then
凭证填制.kh.Caption = ComboBox1.Value
End If

If title.Caption = "个人往来" Then
凭证填制.gr.Caption = ComboBox1.Value
End If

If title.Caption = "部门核算" Then
凭证填制.bm.Caption = ComboBox1.Value
End If

'If title.Caption = "数量核算" Then
'凭证填制.Label9.Caption = ComboBox1.Value
'End If


If title.Caption = "项目核算" Then
凭证填制.xm.Caption = ComboBox1.Value
End If

End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub xg_Click()
Set sh = Sheets("用户档案")
na = TextBox1.text
ps = TextBox2.text
ty = ComboBox1.Value '取得登录窗口中的姓名与密码
If na = "" Or ps = "" Or ty = "" Then MsgBox "存在未输入数据", , "提示": Exit Sub

s = WorksheetFunction.Match(na, sh.[A:A], 0)

sh.Cells(s, 2) = ps: sh.Cells(s, 3) = ty
'MsgBox "修改信息成功", , "提示": Unload 用户信息: Call 用户管理.Reflush

End Sub



Private Sub qd_Click()

If title.Caption = "客户往来" Then
凭证填制.ywy.Caption = TextBox1.text
End If

Unload Me

End Sub

Private Sub UserForm_Activate()

If title.Caption = "客户往来" Then
TextBox1.Visible = True
ywy.Visible = True
kh.Visible = True
Set sh = Sheets("客户档案")
lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
For i = 2 To lastRow
      ComboBox1.AddItem sh.Cells(i, 2)
   Next i

ElseIf title.Caption = "部门核算" Then
bm.Visible = True
Set sh = Sheets("部门档案")
lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
For i = 2 To lastRow
      ComboBox1.AddItem sh.Cells(i, 2)
   Next i
   
ElseIf title.Caption = "个人往来" Then
gr.Visible = True
Set sh = Sheets("人员档案")
lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
For i = 2 To lastRow
      ComboBox1.AddItem sh.Cells(i, 2)
   Next i
   
ElseIf title.Caption = "项目核算" Then
gys.Visible = True
ywy.Visible = True
TextBox1.Visible = True
Set sh = Sheets("供应商档案")
lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
For i = 2 To lastRow
      ComboBox1.AddItem sh.Cells(i, 2)
   Next i
   
End If
End Sub


