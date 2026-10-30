VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 凭证填制 
   Caption         =   "凭证填制"
   ClientHeight    =   11640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   17420
   OleObjectBlob   =   "凭证填制.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "凭证填制"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ignoreChange As Boolean
Dim flag As Integer

Private Sub add1_Click()
If flag = 0 Then
MsgBox "请先新建凭证！"
Else
del1.Visible = True: add1.Visible = False: pzdf3.Visible = True: pzjf3.Visible = True: add2.Visible = True: pzna3.Visible = True
End If
End Sub

Private Sub add2_Click()
del2.Visible = True: add2.Visible = False: pzdf4.Visible = True: pzjf4.Visible = True: add3.Visible = True: pzna4.Visible = True
End Sub

Private Sub add3_Click()
del3.Visible = True: add3.Visible = False: pzdf5.Visible = True: pzjf5.Visible = True: pzna5.Visible = True
End Sub

Private Sub CommandButton4_Click()
Unload Me
系统平台.show

End Sub

Private Sub del1_Click()
del1.Visible = False: add1.Visible = True: pzdf3.Visible = False: pzjf3.Visible = False: pzna3.Visible = False
End Sub

Private Sub del2_Click()
del2.Visible = False: add2.Visible = True: pzdf4.Visible = False: pzjf4.Visible = False: pzna4.Visible = False
End Sub

Private Sub del3_Click()
del3.Visible = False: add3.Visible = True: pzdf5.Visible = False: pzjf5.Visible = False: pzna5.Visible = False
End Sub

Private Sub pzdf1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' 检查输入的字符是否为数字
    If KeyAscii < 48 Or KeyAscii > 57 Then
        ' 如果不是数字，则阻止该字符输入
        KeyAscii = 0
        MsgBox ("请输入数字！"), , "提示"
    End If
End Sub
Private Sub pzdf2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' 检查输入的字符是否为数字
    If KeyAscii < 48 Or KeyAscii > 57 Then
        ' 如果不是数字，则阻止该字符输入
        KeyAscii = 0
        MsgBox ("请输入数字！"), , "提示"
    End If
End Sub
Private Sub pzdf3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' 检查输入的字符是否为数字
    If KeyAscii < 48 Or KeyAscii > 57 Then
        ' 如果不是数字，则阻止该字符输入
        KeyAscii = 0
        MsgBox ("请输入数字！"), , "提示"
    End If
End Sub
Private Sub pzdf4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' 检查输入的字符是否为数字
    If KeyAscii < 48 Or KeyAscii > 57 Then
        ' 如果不是数字，则阻止该字符输入
        KeyAscii = 0
        MsgBox ("请输入数字！"), , "提示"
    End If
End Sub
Private Sub pzdf5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' 检查输入的字符是否为数字
    If KeyAscii < 48 Or KeyAscii > 57 Then
        ' 如果不是数字，则阻止该字符输入
        KeyAscii = 0
        MsgBox ("请输入数字！"), , "提示"
    End If
End Sub

Private Sub pzdf1_Change()
If ignoreChange = False Then
If flag = 0 Then
MsgBox "请先新建凭证！"
ignoreChange = True
pzdf1.text = "0"
ignoreChange = False
Else
Call loanreflush
End If
End If
End Sub

Private Sub pzdf2_Change()
If ignoreChange = False Then
If flag = 0 Then
MsgBox "请先新建凭证！"
ignoreChange = True
pzdf2.text = "0"
ignoreChange = False
Else
Call loanreflush
End If
End If
End Sub

Private Sub pzhj_Click()
If flag = 0 Then
MsgBox "请先新建凭证！"
End If
End Sub

Private Sub pzdf3_Change()
Call loanreflush
End Sub

Private Sub pzdf4_Change()
Call loanreflush
End Sub

Private Sub pzdf5_Change()
Call loanreflush
End Sub
Private Sub pzjf1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' 检查输入的字符是否为数字
    If KeyAscii < 48 Or KeyAscii > 57 Then
        ' 如果不是数字，则阻止该字符输入
        KeyAscii = 0
        MsgBox ("请输入数字！"), , "提示"
    End If
End Sub
Private Sub pzjf2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' 检查输入的字符是否为数字
    If KeyAscii < 48 Or KeyAscii > 57 Then
        ' 如果不是数字，则阻止该字符输入
        KeyAscii = 0
        MsgBox ("请输入数字！"), , "提示"
    End If
End Sub
Private Sub pzjf3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' 检查输入的字符是否为数字
    If KeyAscii < 48 Or KeyAscii > 57 Then
        ' 如果不是数字，则阻止该字符输入
        KeyAscii = 0
        MsgBox ("请输入数字！"), , "提示"
    End If
End Sub
Private Sub pzjf4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' 检查输入的字符是否为数字
    If KeyAscii < 48 Or KeyAscii > 57 Then
        ' 如果不是数字，则阻止该字符输入
        KeyAscii = 0
        MsgBox ("请输入数字！"), , "提示"
    End If
End Sub
Private Sub pzjf5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' 检查输入的字符是否为数字
    If KeyAscii < 48 Or KeyAscii > 57 Then
        ' 如果不是数字，则阻止该字符输入
        KeyAscii = 0
        MsgBox ("请输入数字！"), , "提示"
    End If
End Sub
Private Sub pzjf1_Change()
If ignoreChange = False Then
If flag = 0 Then
MsgBox "请先新建凭证！"
ignoreChange = True
pzjf1.text = "0"
ignoreChange = False
Else
Call borrowreflush
End If
End If
End Sub

Private Sub pzjf2_Change()
If ignoreChange = False Then
If flag = 0 Then
MsgBox "请先新建凭证！"
ignoreChange = True
pzjf2.text = "0"
ignoreChange = False
Else
Call borrowreflush
End If
End If
End Sub

Private Sub pzjf3_Change()
Call borrowreflush
End Sub

Private Sub pzjf4_Change()
Call borrowreflush
End Sub

Private Sub pzjf5_Change()
Call borrowreflush
End Sub


Private Sub pzli1_Click()
pzna1.text = pzli1.Value '将值赋给输入框
Dim fzhs As String
Set sh = Sheets("会计科目表")
    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If pzna1.text = sh.Cells(i, 4) Then '查询到指定名字的科目
        If sh.Cells(i, 1) = sh.Cells(i, 2) Then '如果为一级科目
            If sh.Cells(i, 2) = sh.Cells(i + 1, 2) Then '如果还存在二级子科目，只能选择末科目
                MsgBox "非末级科目，请重新选择"
                pzna1.text = ""
                Exit Sub
            End If
        End If
        If sh.Cells(i, 5) <> "" Then
           ' fzhs = sh.Cells(i, 5)
            
            凭证辅助核算.title.Caption = sh.Cells(i, 5)
            凭证辅助核算.show
        End If
        End If
   Next i



pzli1.Clear
pzli1.Visible = False

End Sub

Private Sub pzli2_Click()
pzna2.text = pzli2.Value '将值赋给输入框
Dim fzhs As String
Set sh = Sheets("会计科目表")
    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If pzna2.text = sh.Cells(i, 4) Then '查询到指定名字的科目
        If sh.Cells(i, 1) = sh.Cells(i, 2) Then '如果为一级科目
            If sh.Cells(i, 2) = sh.Cells(i + 1, 2) Then '如果还存在二级子科目，只能选择末科目
                MsgBox "非末级科目，请重新选择"
                pzna2.text = ""
                Exit Sub
            End If
        End If
        If sh.Cells(i, 5) <> "" Then
           ' fzhs = sh.Cells(i, 5)
            
            凭证辅助核算.title.Caption = sh.Cells(i, 5)
            凭证辅助核算.show
        End If
        End If
   Next i



pzli2.Clear
pzli2.Visible = False

End Sub

Private Sub pzli3_Click()
pzna3.text = pzli3.Value '将值赋给输入框
Dim fzhs As String
Set sh = Sheets("会计科目表")
    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If pzna3.text = sh.Cells(i, 4) Then '查询到指定名字的科目
        If sh.Cells(i, 1) = sh.Cells(i, 2) Then '如果为一级科目
            If sh.Cells(i, 2) = sh.Cells(i + 1, 2) Then '如果还存在二级子科目，只能选择末科目
                MsgBox "非末级科目，请重新选择"
                pzna3.text = ""
                Exit Sub
            End If
        End If
        If sh.Cells(i, 5) <> "" Then
           ' fzhs = sh.Cells(i, 5)
            
            凭证辅助核算.title.Caption = sh.Cells(i, 5)
            凭证辅助核算.show
        End If
        End If
   Next i



pzli3.Clear
pzli3.Visible = False

End Sub

Private Sub pzli4_Click()
pzna4.text = pzli4.Value '将值赋给输入框
Dim fzhs As String
Set sh = Sheets("会计科目表")
    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If pzna4.text = sh.Cells(i, 4) Then '查询到指定名字的科目
        If sh.Cells(i, 1) = sh.Cells(i, 2) Then '如果为一级科目
            If sh.Cells(i, 2) = sh.Cells(i + 1, 2) Then '如果还存在二级子科目，只能选择末科目
                MsgBox "非末级科目，请重新选择"
                pzna4.text = ""
                Exit Sub
            End If
        End If
        If sh.Cells(i, 5) <> "" Then
           ' fzhs = sh.Cells(i, 5)
            
            凭证辅助核算.title.Caption = sh.Cells(i, 5)
            凭证辅助核算.show
        End If
        End If
   Next i



pzli4.Clear
pzli4.Visible = False

End Sub

Private Sub pzli5_Click()
pzna5.text = pzli5.Value '将值赋给输入框
Dim fzhs As String
Set sh = Sheets("会计科目表")
    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If pzna5.text = sh.Cells(i, 4) Then '查询到指定名字的科目
        If sh.Cells(i, 1) = sh.Cells(i, 2) Then '如果为一级科目
            If sh.Cells(i, 2) = sh.Cells(i + 1, 2) Then '如果还存在二级子科目，只能选择末科目
                MsgBox "非末级科目，请重新选择"
                pzna5.text = ""
                Exit Sub
            End If
        End If
        If sh.Cells(i, 5) <> "" Then
           ' fzhs = sh.Cells(i, 5)
            
            凭证辅助核算.title.Caption = sh.Cells(i, 5)
            凭证辅助核算.show
        End If
        End If
   Next i



pzli5.Clear
pzli5.Visible = False

End Sub

Private Sub pzna1_Change()
If ignoreChange = False Then
If flag = 0 Then
MsgBox "请先新建凭证！"
ignoreChange = True
pzna1.text = ""
ignoreChange = False
ElseIf flag = 1 Then
pzli1.Clear
pzli1.Visible = True
Set sh = Sheets("会计科目表")
    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    For i = 2 To lastRow
    If Len(pzna1.text) <= 4 Then
    
        If sh.Range("d" & i) Like "*" & pzna1.text & "*" Or sh.Range("b" & i) Like pzna1.text & "*" Then
        pzli1.AddItem sh.Cells(i, 4)
        pzli1.AddItem ""
        End If
    ElseIf Len(pzna1.text) > 4 Then
        If sh.Range("d" & i) Like "*" & pzna1.text & "*" Or sh.Range("c" & i) Like pzna1.text & "*" Then
        pzli1.AddItem sh.Cells(i, 4)
        pzli1.AddItem ""
        End If
    End If
   Next i
End If
End If


   
End Sub

Private Sub pzna2_Change()
If ignoreChange = False Then
If flag = 0 Then
MsgBox "请先新建凭证！"
ignoreChange = True
pzna2.text = ""
ignoreChange = False
ElseIf flag = 1 Then

pzli2.Clear
pzli2.Visible = True
Set sh = Sheets("会计科目表")
    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    For i = 2 To lastRow
    If Len(pzna2.text) <= 4 Then
    
        If sh.Range("d" & i) Like "*" & pzna2.text & "*" Or sh.Range("b" & i) Like pzna2.text & "*" Then
        pzli2.AddItem sh.Cells(i, 4)
        pzli2.AddItem ""
        End If
    ElseIf Len(pzna2.text) > 4 Then
        If sh.Range("d" & i) Like "*" & pzna2.text & "*" Or sh.Range("c" & i) Like pzna2.text & "*" Then
        pzli2.AddItem sh.Cells(i, 4)
        pzli2.AddItem ""
        End If
    End If
   Next i
End If
End If
End Sub

Private Sub pzna3_Change()
pzli3.Clear
pzli3.Visible = True
Set sh = Sheets("会计科目表")
    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    For i = 2 To lastRow
    If Len(pzna3.text) <= 4 Then
    
        If sh.Range("d" & i) Like "*" & pzna3.text & "*" Or sh.Range("b" & i) Like pzna3.text & "*" Then
        pzli3.AddItem sh.Cells(i, 4)
        pzli3.AddItem ""
        End If
    ElseIf Len(pzna3.text) > 4 Then
        If sh.Range("d" & i) Like "*" & pzna3.text & "*" Or sh.Range("c" & i) Like pzna3.text & "*" Then
        pzli3.AddItem sh.Cells(i, 4)
        pzli3.AddItem ""
        End If
    End If
   Next i
   
End Sub

Private Sub pzna4_Change()
pzli4.Clear
pzli4.Visible = True
Set sh = Sheets("会计科目表")
    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    For i = 2 To lastRow
    If Len(pzna4.text) <= 4 Then
    
        If sh.Range("d" & i) Like "*" & pzna4.text & "*" Or sh.Range("b" & i) Like pzna4.text & "*" Then
        pzli4.AddItem sh.Cells(i, 4)
        pzli4.AddItem ""
        End If
    ElseIf Len(pzna4.text) > 4 Then
        If sh.Range("d" & i) Like "*" & pzna4.text & "*" Or sh.Range("c" & i) Like pzna4.text & "*" Then
        pzli4.AddItem sh.Cells(i, 4)
        pzli4.AddItem ""
        End If
    End If
   Next i
   
End Sub

Private Sub pzna5_Change()
pzli5.Clear
pzli5.Visible = True
Set sh = Sheets("会计科目表")
    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    For i = 2 To lastRow
    If Len(pzna5.text) <= 4 Then
    
        If sh.Range("d" & i) Like "*" & pzna5.text & "*" Or sh.Range("b" & i) Like pzna5.text & "*" Then
        pzli5.AddItem sh.Cells(i, 4)
        pzli5.AddItem ""
        End If
    ElseIf Len(pzna5.text) > 4 Then
        If sh.Range("d" & i) Like "*" & pzna5.text & "*" Or sh.Range("c" & i) Like pzna5.text & "*" Then
        pzli5.AddItem sh.Cells(i, 4)
        pzli5.AddItem ""
        End If
    End If
   Next i
   
End Sub

Private Sub pztp_Change()
If ignoreChange = False Then
If flag = 0 Then
MsgBox "请先新建凭证！"
ignoreChange = True
pztp.text = ""
ignoreChange = False
Else
If pztp.text = "收款凭证" Then Label2.Caption = "收 款 凭 证"
If pztp.text = "付款凭证" Then Label2.Caption = "付 款 凭 证"
If pztp.text = "转账凭证" Then Label2.Caption = "转 账 凭 证"
If pztp.text = "" Then Label2.Caption = "记 账 凭 证"
End If
End If
End Sub

Private Sub pzzy_Change()
If ignoreChange = False Then
If flag = 0 Then
MsgBox "请先新建凭证！"
ignoreChange = True
pzzy.text = ""
ignoreChange = False
ElseIf flag = 1 Then
End If
End If
End Sub

Private Sub save_Click()

Set sh = Sheets("凭证汇总")
lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
If pztp.text = "" Then MsgBox "未选择凭证类型", , "提示": Exit Sub
If pzzy.text = "" Then MsgBox "未输入摘要", , "提示": Exit Sub
If pzjf.Caption <> pzdf.Caption Then MsgBox "借贷不平衡", , "提示": Exit Sub

If 科目检查() = 1 Then Exit Sub
pzzd.Caption = Sheets("用户档案").Cells(currentUser, "B").Value

If lastRow = 1 Then Call 存储凭证(Val(lastRow), 1, "待审核") Else Call 存储凭证(Val(lastRow), Val(sh.Cells(lastRow, 1)) + 1, "待审核")
MsgBox "新建成功", , "提示"
Unload Me
凭证填制.show

End Sub

Private Sub TextBox1_Change()
If ignoreChange = False Then
If flag = 0 Then
MsgBox "请先新建凭证！"
ignoreChange = True
TextBox1.text = ""
ignoreChange = False
End If
End If

End Sub

Private Sub UserForm_Initialize()

flag = 0
ignoreChange = False
pzdate.Caption = Date

pztp.AddItem "收款凭证"
pztp.AddItem "付款凭证"
pztp.AddItem "转账凭证"
pztp.AddItem ""

End Sub

Public Sub borrowreflush()
Dim sum As Double
Dim jf1 As Double: Dim jf2 As Double: Dim jf3 As Double: Dim jf4 As Double: Dim jf5 As Double
'CDbl将文本类型转换为double
    If pzjf1.text = "" Then jf1 = 0 Else jf1 = CDbl(pzjf1.text)
    If pzjf2.text = "" Then jf2 = 0 Else jf2 = CDbl(pzjf2.text)
    If pzjf3.text = "" Then jf3 = 0 Else jf3 = CDbl(pzjf3.text)
    If pzjf4.text = "" Then jf4 = 0 Else jf4 = CDbl(pzjf4.text)
    If pzjf5.text = "" Then jf5 = 0 Else jf5 = CDbl(pzjf5.text)
sum = jf1 + jf2 + jf3 + jf4 + jf5

pzjf.Caption = sum

End Sub

Public Sub loanreflush()
Dim sum As Double
Dim df1 As Double: Dim df2 As Double: Dim df3 As Double: Dim df4 As Double: Dim df5 As Double
'CDbl将文本类型转换为double
    If pzdf1.text = "" Then df1 = 0 Else df1 = CDbl(pzdf1.text)
    If pzdf2.text = "" Then df2 = 0 Else df2 = CDbl(pzdf2.text)
    If pzdf3.text = "" Then df3 = 0 Else df3 = CDbl(pzdf3.text)
    If pzdf4.text = "" Then df4 = 0 Else df4 = CDbl(pzdf4.text)
    If pzdf5.text = "" Then df5 = 0 Else df5 = CDbl(pzdf5.text)
sum = df1 + df2 + df3 + df4 + df5

pzdf.Caption = sum

End Sub

Public Function 科目检查() As Integer

科目检查 = 0
Set sh = Sheets("会计科目表")
isborrow = False: isloan = False
j = 0
Dim arr(1 To 5)
For i = 1 To 5
tbborrow = "pzjf" & i: tbloan = "pzdf" & i: tbna = "pzna" & i
If Controls(tbna).Visible = True Then
If j = 0 Then
j = j + 1
arr(1) = Controls(tbna).text
Else
For k = 1 To j '相同科目不能重复
If Controls(tbna).text = arr(j) Then MsgBox "科目重复", , "提示": 科目检查 = 1: Exit Function
Next
End If
If Controls(tbna).text = "" Then MsgBox "未输入科目", , "提示": 科目检查 = 1: Exit Function

s = WorksheetFunction.Match(Controls(tbna).text, sh.[D:D], 0) 'Controls根据名字生成对应控件

If sh.Cells(s, 2) = "1001" Or sh.Cells(s, 2) = "1002" Then
If Controls(tbborrow).text > 0 Then isborrow = True Else isloan = True
End If: End If
Next
If pztp.text = "收款凭证" And isborrow = False Then MsgBox "收款凭证借方必有库存现金/银行存款科目", , "提示": 科目检查 = 1: Exit Function
If pztp.text = "付款凭证" And isloan = False Then MsgBox "付款凭证贷方必有库存现金/银行存款科目", , "提示": 科目检查 = 1: Exit Function

End Function

Public Sub 存储凭证(s As Integer, id As Integer, state As String)

Set sh = Sheets("凭证汇总")
For i = 1 To 5
tbborrow = "pzjf" & i: tbloan = "pzdf" & i: tbna = "pzna" & i

If Controls(tbna).Visible = True Then
sh.Cells(s + i, 1) = id
sh.Cells(s + i, 2) = pzdate.Caption
sh.Cells(s + i, 3) = pztp.text
sh.Cells(s + i, 4) = pzzy.text
sh.Cells(s + i, 5) = Controls(tbna).text
sh.Cells(s + i, 6) = Controls(tbborrow).text
sh.Cells(s + i, 7) = Controls(tbloan).text
sh.Cells(s + i, 8) = pzzd.Caption
sh.Cells(s + i, 9) = pzsh.Caption
sh.Cells(s + i, 10) = pzcn.Caption
sh.Cells(s + i, 11) = pzjz.Caption
sh.Cells(s + i, 12) = state
sh.Cells(s + i, 13) = bm.Caption
sh.Cells(s + i, 14) = gr.Caption
sh.Cells(s + i, 15) = kh.Caption
sh.Cells(s + i, 16) = xm.Caption
End If
Next
pzid.Caption = id
Call 凭证处理.Reflush:
End Sub

Private Sub xj_Click()
flag = 1
End Sub

