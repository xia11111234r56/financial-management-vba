VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 凭证处理 
   Caption         =   "凭证处理"
   ClientHeight    =   8570.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   16030
   OleObjectBlob   =   "凭证处理.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "凭证处理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()

End Sub

Private Sub ComboBox1_Change()
Dim itm As ListItem
TextBox1.text = ""
ComboBox2.Value = ""
ListView1.ListItems.Clear

Set sh = Sheets("凭证汇总")

lastColumn = sh.Cells(1, sh.Columns.count).End(xlToLeft).Column
lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row

For i = 2 To lastRow
    If sh.Cells(i - 1, 1) <> sh.Cells(i, 1) Then
    If (Sheets("凭证汇总").Cells(i, 3) = ComboBox1.text And Sheets("凭证汇总").Cells(i, 12) = ComboBox2.text) Or (Sheets("凭证汇总").Cells(i, 3) = ComboBox1.text And ComboBox2.text = "") Or (ComboBox1.text = "" And Sheets("凭证汇总").Cells(i, 12) = ComboBox2.text) Then
       If sh.Cells(i, 12) <> "已作废" Then
        Set itm = ListView1.ListItems.Add
        itm.text = sh.Cells(i, 1)
        k = 1
        For j = 2 To lastColumn
        If j <= 4 Or j >= 8 Then
        itm.SubItems(k) = sh.Cells(i, j)
        k = k + 1
        End If
        Next
        End If
    ElseIf ComboBox1.text = "" And ComboBox2.text = "" Then
        If sh.Cells(i, 12) <> "已作废" Then
        Set itm = ListView1.ListItems.Add
        itm.text = sh.Cells(i, 1)
        X = 1
        For j = 2 To lastColumn
        If j <= 4 Or j >= 8 Then
        Set temp = itm.ListSubItems.Add
        temp.text = sh.Cells(i, j)
        End If
        Next j
        End If
    End If
    End If
Next i
End Sub

Private Sub ComboBox2_Change()
Dim itm As ListItem
TextBox1.text = ""
ListView1.ListItems.Clear

Set sh = Sheets("凭证汇总")

lastColumn = sh.Cells(1, sh.Columns.count).End(xlToLeft).Column
lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row

For i = 2 To lastRow
    If sh.Cells(i - 1, 1) <> sh.Cells(i, 1) Then
    If (Sheets("凭证汇总").Cells(i, 3) = ComboBox1.text And Sheets("凭证汇总").Cells(i, 12) = ComboBox2.text) Or (Sheets("凭证汇总").Cells(i, 3) = ComboBox1.text And ComboBox2.text = "") Or (ComboBox1.text = "" And Sheets("凭证汇总").Cells(i, 12) = ComboBox2.text) Then
       If sh.Cells(i, 12) <> "已作废" Then
        Set itm = ListView1.ListItems.Add
        itm.text = sh.Cells(i, 1)
        k = 1
        For j = 2 To lastColumn
        If j <= 4 Or j >= 8 Then
        itm.SubItems(k) = sh.Cells(i, j)
        k = k + 1
        End If
        Next
        End If
    ElseIf ComboBox1.text = "" And ComboBox2.text = "" Then
        If sh.Cells(i, 12) <> "已作废" Then
        Set itm = ListView1.ListItems.Add
        itm.text = sh.Cells(i, 1)
        X = 1
        For j = 2 To lastColumn
        If j <= 4 Or j >= 8 Then
        Set temp = itm.ListSubItems.Add
        temp.text = sh.Cells(i, j)
        End If
        Next j
        End If
    End If
    End If
Next i
End Sub

Private Sub ComboBox3_Change()
Dim itm As ListItem
TextBox1.text = ""
ListView1.ListItems.Clear

Set sh = Sheets("会计科目表")

lastColumn = sh.Cells(1, sh.Columns.count).End(xlToLeft).Column
lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row

For i = 2 To lastRow
    If (Sheets("会计科目表").Cells(i, 5) = ComboBox1.text And Sheets("会计科目表").Cells(i, 7) = ComboBox2.text And Sheets("会计科目表").Cells(i, 8) = ComboBox3.text) Or (Sheets("会计科目表").Cells(i, 5) = ComboBox1.text And ComboBox2.text = "" And ComboBox3.text = "") Or (Sheets("会计科目表").Cells(i, 5) = ComboBox1.text And Sheets("会计科目表").Cells(i, 7) = ComboBox2.text And ComboBox3.text = "") Or (Sheets("会计科目表").Cells(i, 5) = ComboBox1.text And ComboBox2.text = "" And Sheets("会计科目表").Cells(i, 8) = ComboBox3.text) Or (ComboBox1.text = "" And Sheets("会计科目表").Cells(i, 7) = ComboBox2.text And ComboBox3.text = "") Or (ComboBox1.text = "" And Sheets("会计科目表").Cells(i, 7) = ComboBox2.text And Sheets("会计科目表").Cells(i, 8) = ComboBox3.text) Or (ComboBox1.text = "" And ComboBox2.text = "" And Sheets("会计科目表").Cells(i, 8) = ComboBox3.text) Then
        Set itm = ListView1.ListItems.Add
        itm.text = sh.Cells(i, 1)
        
        For j = 4 To lastColumn
        
        itm.SubItems(j - 3) = sh.Cells(i, j)
        
        Next j
    ElseIf ComboBox1.text = "" And ComboBox2.text = "" And ComboBox3.text = "" Then
        Set itm = ListView1.ListItems.Add
        itm.text = sh.Cells(i, 1)
        For j = 4 To lastColumn
        Set temp = List.ListSubItems.Add
        temp.text = sh.Cells(i, j)
        
        Next j
        
    End If
Next i
End Sub

Private Sub CommandButton1_Click()
    Dim lastrow1 As Long
    Dim i       As Long
    Dim del     As String
    Dim sh      As Worksheet
    Dim s       As Long

    ' 1. 检查本月是否已结账
    lastrow1 = Sheets("结账登记表").Cells(Sheets("结账登记表").Rows.count, "A").End(xlUp).Row
    For i = 2 To lastrow1
        If Sheets("结账登记表").Cells(i, 1).Value = Month(Date) Then
            MsgBox "本月已结账，不能再进行修改", , "提示"
            Exit Sub
        End If
    Next i

    ' 2. 必须先选中一个凭证
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "未选择凭证", , "提示"
        Exit Sub
    End If

    ' 3. 取出选中凭证的“凭证号”
    del = CStr(ListView1.ListItems(ListView1.SelectedItem.Index).text)
    Set sh = Sheets("凭证汇总")
    s = WorksheetFunction.Match(Val(del), sh.[A:A], 0)

    ' 4. 只有“制单人”本人才能修改
    If sh.Cells(s, 8).Value <> Sheets("用户档案").Cells(currentUser, "B").Value Then
        MsgBox "非制单人，无法修改", , "提示"
        Exit Sub
    End If

    ' 5. 只有“待审核”状态才能修改
    If sh.Cells(s, 12).Value <> "待审核" Then
        MsgBox "非待审核凭证，无法修改", , "提示"
        Exit Sub
    End If

    ' 6. 打开“凭证信息”窗体，显示修改界面
    凭证信息.xg.Visible = True
    Call 函数.生成凭证(Val(del))
    凭证信息.show
End Sub


Private Sub CommandButton2_Click()
    Dim lastrow1 As Long
    Dim i As Long
    Dim del As String
    Dim sh As Worksheet
    Dim s As Long
    Dim obj As Object

    ' 1. 判断是否已结账
    lastrow1 = Sheets("结账登记表").Cells(Sheets("结账登记表").Rows.count, "A").End(xlUp).Row
    For i = 2 To lastrow1
        If Sheets("结账登记表").Cells(i, 1).Value = Month(Date) Then
            MsgBox "本月已结账，不能再进行修改", , "提示"
            Exit Sub
        End If
    Next i

    ' 2. 检查是否选中凭证
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "未选择凭证", , "提示"
        Exit Sub
    End If

    ' 3. 获取凭证编号
    del = CStr(ListView1.ListItems(ListView1.SelectedItem.Index).text)
    Set sh = Sheets("凭证汇总")

    ' 4. 查找凭证所在行
    s = WorksheetFunction.Match(Val(del), sh.[A:A], 0)

    ' 5. 检查是否为制单人
    If sh.Cells(s, 8).Value <> Sheets("用户档案").Cells(currentUser, "B").Value Then
        MsgBox "非制单人，无法作废", , "提示"
        Exit Sub
    End If

    ' 6. 仅“待审核”或“已作废”状态允许作废
    If sh.Cells(s, 12).Value <> "待审核" And sh.Cells(s, 12).Value <> "已作废" Then
        MsgBox "非待审核/已作废凭证，无法作废/取消作废", , "提示"
        Exit Sub
    End If

    ' 7. 禁用控件并显示作废按钮
    For Each obj In 凭证信息.Controls
        obj.Enabled = False
    Next

    凭证信息.CommandButton4.Enabled = True
    凭证信息.zf.Enabled = True
    凭证信息.zf.Visible = True
    凭证信息.zlpz.Enabled = True
    凭证信息.zlpz.Visible = True

    ' 8. 打开凭证信息窗体
    Call 函数.生成凭证(Val(del))
    凭证信息.show
End Sub


Private Sub CommandButton3_Click()
    Dim del As String
    Dim s As Long
    Dim sh As Worksheet
    Dim obj As Object

    ' 1. 必须先选中一个凭证
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "未选择凭证", , "提示"
        Exit Sub
    End If

    ' 2. 取出选中项的凭证号
    del = CStr(ListView1.ListItems(ListView1.SelectedItem.Index).text)
    Set sh = Sheets("凭证汇总")    ' 凭证汇总表

    s = WorksheetFunction.Match(Val(del), sh.[A:A], 0)

    ' 3. 判断当前用户是否为账套主管
    If Sheets("用户档案").Cells(currentUser, "D").Value <> "账套主管" Then
        MsgBox "非账套主管，不可审核", , "提示"
        Exit Sub
    End If

    ' 4. 只有“待审核”或“已审核”状态，才能进入审核/取消审核界面
    If sh.Cells(s, 12).Value <> "待审核" And sh.Cells(s, 12).Value <> "已审核" Then
        MsgBox "非待审核/已审核凭证，无法审核/取消审核", , "提示"
        Exit Sub
    End If

    ' 5. 禁用凭证信息窗体上的所有控件，只保留审核相关按钮可用
    For Each obj In 凭证信息.Controls
        obj.Enabled = False
    Next
    凭证信息.CommandButton4.Enabled = True   ' 例如返回按钮
    凭证信息.shpz.Enabled = True              ' 签批按钮
    凭证信息.shpz.Visible = True              ' 确保签批按钮可见

    ' 6. 调用生成凭证子程序并显示凭证信息窗体
    Call 函数.生成凭证(Val(del))
    凭证信息.show
End Sub



Private Sub CommandButton4_Click()
Unload Me

End Sub

Private Sub CommandButton5_Click()
If (ListView1.SelectedItem.Selected = False) Then MsgBox "未选择凭证", , "提示": Exit Sub
del = CStr(ListView1.ListItems(ListView1.SelectedItem.Index).text)
Dim obj As Object
For Each obj In 凭证信息.Controls
obj.Enabled = False
Next
凭证信息.CommandButton4.Enabled = True
Call 函数.生成凭证(Val(del))

凭证信息.show
End Sub

Private Sub CommandButton6_Click()
If (ListView1.SelectedItem.Selected = False) Then MsgBox "未选择凭证", , "提示": Exit Sub
del = CStr(ListView1.ListItems(ListView1.SelectedItem.Index).text)
Set sh = Sheets("凭证汇总") '显示凭证信息列表

s = WorksheetFunction.Match(Val(del), sh.[A:A], 0)
If sh.Cells(s, 12) <> "已签字" And sh.Cells(s, 12) <> "已审核" Then MsgBox "非已审核/已签字凭证，无法签字/取消签字", , "提示": Exit Sub
Dim obj As Object
For Each obj In 凭证填制.Controls
obj.Enabled = False
Next
凭证信息.CommandButton4.Enabled = True
凭证信息.qzpz.Enabled = True
凭证信息.qzpz.Visible = True

Call 函数.生成凭证(Val(del))
凭证信息.show

End Sub





Private Sub TextBox1_Change()
If Len(TextBox1) >= 1 Then
        Call serch
        
    Else
        Call Reflush
    End If
End Sub

Private Sub UserForm_Activate()
  

ComboBox1.AddItem "付款凭证"
ComboBox1.AddItem "收款凭证"
ComboBox1.AddItem "转账凭证"
ComboBox1.AddItem ""

ComboBox2.AddItem "待审核"
ComboBox2.AddItem "已审核"
ComboBox2.AddItem "已签字"
ComboBox2.AddItem "已记账"
ComboBox2.AddItem ""

Call Reflush

End Sub

Public Sub Reflush()

ListView1.ColumnHeaders.Clear:

ListView1.ListItems.Clear '清空列表

zy = TextBox1.text
tp = ComboBox1.text
state = ComboBox2.text
Set sh = Sheets("凭证汇总") '显示凭证信息列表

Dim lastRow, i As Integer

lastColumn = sh.Cells(1, sh.Columns.count).End(xlToLeft).Column
lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row

ListView1.ColumnHeaders.Add 1, , sh.Cells(1, 1)
j = 2
For i = 2 To lastColumn
If i <= 4 Or i >= 8 Then
ListView1.ColumnHeaders.Add j, , sh.Cells(1, i)
j = j + 1

End If
Next

ListView1.View = lvwReport: ListView1.FullRowSelect = True: ListView1.Gridlines = True

For i = 2 To lastRow
If sh.Cells(i - 1, 1) <> sh.Cells(i, 1) Then
If zy = "" Or sh.Cells(i, 4) Like "*" & zy & "*" Then
If tp = "" Or sh.Cells(i, 3) Like tp Then
If state = "" Or sh.Cells(i, 12) Like state Then
Set itm = ListView1.ListItems.Add
itm.text = sh.Cells(i, 1)
N = 1
For j = 2 To lastColumn
If j <= 4 Or j >= 8 Then
itm.SubItems(N) = sh.Cells(i, j)
N = N + 1

End If
Next
End If: End If: End If: End If
Next

If ListView1.ListItems.count > 0 Then ListView1.ListItems(1).Selected = False

End Sub
Sub serch()

    Set sh = Sheets("凭证汇总")
    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    For i = 2 To lastRow
    If sh.Range("d" & i) Like "*" & TextBox1.text & "*" Then
        Call Reflush
    End If
   Next i
End Sub

