Attribute VB_Name = "函数"
Sub 隐藏表()

For i = 1 To Worksheets.count
 

If Sheets(i).name <> "主页" Then
 

Sheets(i).Visible = 2 '隐藏表
 

Else
 

Sheets(i).Visible = -1 '只让“主页”显示出来
 

End If
 

Next
 

End Sub


Public Sub 生成凭证(id As Integer)

Set sh = Sheets("凭证汇总")


凭证信息.pzli1.Visible = False: 凭证信息.pzli2.Visible = False

lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
j = 1
For i = 2 To lastRow
If sh.Cells(i, 1) = id Then
tbjf = "pzjf" & j: tbdf = "pzdf" & j: tbk = "pzna" & j: Add = "add" & j: del = "del" & j
凭证信息.pzid.Caption = sh.Cells(i, 1): 凭证信息.pzdate.Caption = sh.Cells(i, 2)
凭证信息.pztp.text = sh.Cells(i, 3): 凭证信息.pzzy.text = sh.Cells(i, 4)
凭证信息.pzzd.Caption = sh.Cells(i, 8): 凭证信息.pzsh.Caption = sh.Cells(i, 9)
凭证信息.pzcn.Caption = sh.Cells(i, 10): 凭证信息.pzjz.Caption = sh.Cells(i, 11)
凭证信息.bm.Caption = sh.Cells(i, 13)
If 凭证信息.Controls(tbk).Visible = False Then 凭证信息.Controls(tbk).Visible = True: 凭证信息.Controls(tbjf).Visible = True: 凭证信息.Controls(tbdf).Visible = True: 凭证信息.Controls(del).Visible = True: 凭证信息.Controls(Add).Visible = False:

凭证信息.Controls(tbk).text = sh.Cells(i, 5): 凭证信息.Controls(tbjf).text = sh.Cells(i, 6): 凭证信息.Controls(tbdf).text = sh.Cells(i, 7): j = j + 1
If sh.Cells(i, 12) = "已作废" Then
凭证信息.zflabel.Visible = True
End If
End If
Next
凭证信息.pzli1.Visible = False
凭证信息.pzli2.Visible = False
End Sub


Public Sub 记账()
Set sh = Sheets("凭证汇总")
lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
Dim k As Integer
k = sh.Cells(2, 1)
For i = 2 To lastRow
If sh.Cells(i, 1) = k Then
    Call 入账(k)
    k = -1
    For j = i To lastRow
    If sh.Cells(i, 1) <> sh.Cells(j, 1) Then
    k = sh.Cells(j, 1): Exit For: End If
    Next
End If
Next
End Sub

Public Sub 入账(id As Integer)
    Dim sh As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim yy As Worksheet
    Dim j As Long
    Dim zz As Worksheet
    Dim N As Long
    Dim km As Worksheet
    Dim currentUserName As String
    
    ' 获取当前登录用户姓名
    currentUserName = Sheets("用户档案").Cells(currentUser, "B").Value
    
    Set sh = Sheets("凭证汇总")
    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If sh.Cells(i, 1).Value = id Then
            ' 标记已记账，并记录记账人
            sh.Cells(i, 12).Value = "已记账"
            sh.Cells(i, 11).Value = currentUserName
            
            Set yy = Sheets("发生额及余额表")
            j = WorksheetFunction.Match(CStr(sh.Cells(i, 5).Value), yy.[D:D], 0)
            
            ' 更新发生额及余额表借贷发生额
            yy.Cells(j, 8).Value = CStr(Val(yy.Cells(j, 8).Value) + Val(sh.Cells(i, 6).Value))
            yy.Cells(j, 9).Value = CStr(Val(yy.Cells(j, 9).Value) + Val(sh.Cells(i, 7).Value))
            Call 更新一级科目("发生额及余额表", 8)
            Call 更新一级科目("发生额及余额表", 9)
            Call 更新期末余额
            
            ' 更新总账
            Set zz = Sheets("总账")
            N = zz.Cells(zz.Rows.count, "A").End(xlUp).Row
            zz.Cells(N + 1, 1).Value = yy.Cells(j, 4).Value
            zz.Cells(N + 1, 2).Value = sh.Cells(i, 2).Value
            zz.Cells(N + 1, 3).Value = sh.Cells(i, 3).Value
            zz.Cells(N + 1, 4).Value = sh.Cells(i, 1).Value
            zz.Cells(N + 1, 5).Value = sh.Cells(i, 4).Value
            zz.Cells(N + 1, 6).Value = sh.Cells(i, 6).Value
            zz.Cells(N + 1, 7).Value = sh.Cells(i, 7).Value
            If yy.Cells(j, 10).Value > yy.Cells(j, 11).Value Then
                zz.Cells(N + 1, 8).Value = "借"
                zz.Cells(N + 1, 9).Value = yy.Cells(j, 10).Value - yy.Cells(j, 11).Value
            Else
                zz.Cells(N + 1, 8).Value = "贷"
                zz.Cells(N + 1, 9).Value = yy.Cells(j, 11).Value - yy.Cells(j, 10).Value
            End If
            
            ' 更新明细账
            Set zz = Sheets("明细账")
            N = zz.Cells(zz.Rows.count, "A").End(xlUp).Row
            zz.Cells(N + 1, 1).Value = sh.Cells(i, 5).Value
            zz.Cells(N + 1, 2).Value = sh.Cells(i, 2).Value
            zz.Cells(N + 1, 3).Value = sh.Cells(i, 3).Value
            zz.Cells(N + 1, 4).Value = sh.Cells(i, 1).Value
            zz.Cells(N + 1, 5).Value = sh.Cells(i, 4).Value
            zz.Cells(N + 1, 6).Value = sh.Cells(i, 6).Value
            zz.Cells(N + 1, 7).Value = sh.Cells(i, 7).Value
            If yy.Cells(j, 10).Value > yy.Cells(j, 11).Value Then
                zz.Cells(N + 1, 8).Value = "借"
                zz.Cells(N + 1, 9).Value = yy.Cells(j, 10).Value - yy.Cells(j, 11).Value
            Else
                zz.Cells(N + 1, 8).Value = "贷"
                zz.Cells(N + 1, 9).Value = yy.Cells(j, 11).Value - yy.Cells(j, 10).Value
            End If
            
            ' 更新现金日记账或银行日记账（如有辅助核算要求）
            Set km = Sheets("会计科目表")
            If km.Cells(j, 7).Value = "Y" And km.Cells(j, 8).Value = "" Then
                Set zz = Sheets("现金日记账")
                N = zz.Cells(zz.Rows.count, "A").End(xlUp).Row
                zz.Cells(N + 1, 1).Value = sh.Cells(i, 5).Value
                zz.Cells(N + 1, 2).Value = sh.Cells(i, 2).Value
                zz.Cells(N + 1, 3).Value = sh.Cells(i, 3).Value
                zz.Cells(N + 1, 4).Value = sh.Cells(i, 1).Value
                zz.Cells(N + 1, 5).Value = sh.Cells(i, 4).Value
                zz.Cells(N + 1, 6).Value = sh.Cells(i, 6).Value
                zz.Cells(N + 1, 7).Value = sh.Cells(i, 7).Value
                If yy.Cells(j, 10).Value > yy.Cells(j, 11).Value Then
                    zz.Cells(N + 1, 8).Value = "借"
                    zz.Cells(N + 1, 9).Value = yy.Cells(j, 10).Value - yy.Cells(j, 11).Value
                Else
                    zz.Cells(N + 1, 8).Value = "贷"
                    zz.Cells(N + 1, 9).Value = yy.Cells(j, 11).Value - yy.Cells(j, 10).Value
                End If
            End If
            
            If km.Cells(j, 7).Value = "Y" And km.Cells(j, 8).Value = "Y" Then
                Set zz = Sheets("银行日记账")
                N = zz.Cells(zz.Rows.count, "A").End(xlUp).Row
                zz.Cells(N + 1, 1).Value = sh.Cells(i, 5).Value
                zz.Cells(N + 1, 2).Value = sh.Cells(i, 2).Value
                zz.Cells(N + 1, 3).Value = sh.Cells(i, 3).Value
                zz.Cells(N + 1, 4).Value = sh.Cells(i, 1).Value
                zz.Cells(N + 1, 5).Value = sh.Cells(i, 4).Value
                zz.Cells(N + 1, 6).Value = sh.Cells(i, 6).Value
                zz.Cells(N + 1, 7).Value = sh.Cells(i, 7).Value
                If yy.Cells(j, 10).Value > yy.Cells(j, 11).Value Then
                    zz.Cells(N + 1, 8).Value = "借"
                    zz.Cells(N + 1, 9).Value = yy.Cells(j, 10).Value - yy.Cells(j, 11).Value
                Else
                    zz.Cells(N + 1, 8).Value = "贷"
                    zz.Cells(N + 1, 9).Value = yy.Cells(j, 11).Value - yy.Cells(j, 10).Value
                End If
            End If
        End If
    Next i
End Sub


Public Sub 更新期末余额()
Set sh = Sheets("发生额及余额表")
lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
For i = 2 To lastRow
fina = Val(sh.Cells(i, 6)) + Val(sh.Cells(i, 8)) - Val(sh.Cells(i, 7)) - Val(sh.Cells(i, 9))
If fina > 0 Then sh.Cells(i, 10) = fina: sh.Cells(i, 11) = 0 Else sh.Cells(i, 11) = 0 - fina: sh.Cells(i, 10) = 0
Next
End Sub

Public Sub 更新一级科目(sna As String, Data As Integer)
Set sh = Sheets(sna)
lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
For i = 2 To lastRow - 1
If sh.Cells(i, 1) = sh.Cells(i, 2) And sh.Cells(i + 1, 2) = sh.Cells(i, 2) Then
sh.Cells(i, Data) = 0
For j = i + 1 To lastRow
If sh.Cells(i, 2) = sh.Cells(j, 2) Then sh.Cells(i, Data) = sh.Cells(i, Data) + sh.Cells(j, Data)
Next
End If
Next
End Sub

Public Function 核算(she As String, ind As Integer, ts As Integer) As Integer

Set sh = Sheets(she)

lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row

jfqc = 0: dfqc = 0

For i = 2 To lastRow
'一级科目
If sh.Cells(i, 1) = sh.Cells(i, 2) Then
jfqc = jfqc + sh.Cells(i, ind)
dfqc = dfqc + sh.Cells(i, ind + 1)
End If

Next
'Chr(13):回车符
If ts = 1 Then MsgBox ("期初余额：" & Chr(13) & "借方总额：" & jfqc & Chr(13) & "贷方总额：" & dfqc), , "提示"

If ts = 2 Then MsgBox ("发生额：" & Chr(13) & "借方总额：" & jfqc & Chr(13) & "贷方总额：" & dfqc), , "提示"

If ts = 3 Then MsgBox ("期末余额：" & Chr(13) & "借方总额：" & jfqc & Chr(13) & "贷方总额：" & dfqc), , "提示"

If jfqc = dfqc Then 核算 = 0 Else 核算 = 1

End Function


Public Sub 利润结转()
    Dim sh As Worksheet
    Dim zz As Worksheet
    Dim wsChk As Worksheet
    Dim lastRow As Long
    Dim lastRowChk As Long
    Dim currentUserName As String
    Dim currentUserRole As String
    Dim i As Long
    Dim thisMonth As Integer
    Dim foundClosed As Boolean
    Dim alreadyPosted As Boolean

    ' 获取当前登录用户姓名和角色（角色为第4列，即D列）
    currentUserName = Sheets("用户档案").Cells(currentUser, "B").Value
    currentUserRole = Sheets("用户档案").Cells(currentUser, "D").Value

    '—— 0. 权限判断 ——
    If currentUserRole <> "账套主管" And currentUserRole <> "会计人员" Then
        MsgBox "您没有权限进行月末损益结转操作！", vbCritical, "权限限制"
        Exit Sub
    End If

    thisMonth = Month(Date)
    
    '—— 1. 检查是否已结账 ——
    Set wsChk = Sheets("结账登记表")
    lastRowChk = wsChk.Cells(wsChk.Rows.count, "A").End(xlUp).Row
    foundClosed = False
    For i = 2 To lastRowChk
        If wsChk.Cells(i, "A").Value = thisMonth Then
            foundClosed = True
            Exit For
        End If
    Next i

    If foundClosed Then
        MsgBox "本月已结账，无法进行结转！", vbExclamation, "提示"
        Exit Sub
    End If

    '—— 2. 检查是否已经结转过 ——
    Set sh = Sheets("凭证汇总")
    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    alreadyPosted = False
    For i = 2 To lastRow
        If Month(sh.Cells(i, 2).Value) = thisMonth And sh.Cells(i, 4).Value = "月末损益结转" Then
            alreadyPosted = True
            Exit For
        End If
    Next i

    If alreadyPosted Then
        MsgBox "本月月末损益已结转，无法再次结转！", vbExclamation, "提示"
        Exit Sub
    End If

    '—— 3. 获取利润表数值进行结转 ——
    Set zz = Sheets("利润表")
    If zz.Cells(13, 2).Value = 0 Then
        MsgBox "月末损益为0，不需结转！", , "提示"
        Exit Sub
    End If

    '—— 4. 写入两行结转凭证 ——
    sh.Cells(lastRow + 1, 1).Value = sh.Cells(lastRow, 1).Value + 1
    sh.Cells(lastRow + 1, 2).Value = CStr(Date)
    sh.Cells(lastRow + 1, 3).Value = "转账凭证"
    sh.Cells(lastRow + 1, 4).Value = "月末损益结转"
    sh.Cells(lastRow + 1, 5).Value = "主营业务收入"
    sh.Cells(lastRow + 1, 6).Value = zz.Cells(13, 2).Value
    sh.Cells(lastRow + 1, 7).Value = 0
    sh.Cells(lastRow + 1, 8).Value = currentUserName
    sh.Cells(lastRow + 1, 9).Value = ""
    sh.Cells(lastRow + 1, 10).Value = ""
    sh.Cells(lastRow + 1, 11).Value = ""
    sh.Cells(lastRow + 1, 12).Value = "待审核"

    sh.Cells(lastRow + 2, 1).Value = sh.Cells(lastRow, 1).Value + 1
    sh.Cells(lastRow + 2, 2).Value = CStr(Date)
    sh.Cells(lastRow + 2, 3).Value = "转账凭证"
    sh.Cells(lastRow + 2, 4).Value = "月末损益结转"
    sh.Cells(lastRow + 2, 5).Value = "本年利润"
    sh.Cells(lastRow + 2, 6).Value = 0
    sh.Cells(lastRow + 2, 7).Value = zz.Cells(13, 2).Value
    sh.Cells(lastRow + 2, 8).Value = currentUserName
    sh.Cells(lastRow + 2, 9).Value = ""
    sh.Cells(lastRow + 2, 10).Value = ""
    sh.Cells(lastRow + 2, 11).Value = ""
    sh.Cells(lastRow + 2, 12).Value = "待审核"

    MsgBox "月末损益结转成功！", , "提示"
End Sub

Public Sub 结账()
    Dim sh As Worksheet
    Dim zz As Worksheet
    Dim yy As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim n1 As Long
    Dim thisMonth As Integer
    Dim foundClosed As Boolean

    thisMonth = Month(Date)
    
    '—— 0. 检查是否已结账 ——
    Set yy = Sheets("结账登记表")
    n1 = yy.Cells(yy.Rows.count, "A").End(xlUp).Row
    foundClosed = False
    For i = 2 To n1
        If yy.Cells(i, "A").Value = thisMonth Then
            foundClosed = True
            Exit For
        End If
    Next i
    
    If foundClosed Then
        MsgBox "本月已结账，无需再次结账！", vbExclamation, "提示"
        Exit Sub
    End If

    '—— 1. 检查凭证是否全部“已记账”或“已作废” ——
    Set sh = Sheets("凭证汇总")
    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If sh.Cells(i, 12).Value <> "已作废" And sh.Cells(i, 12).Value <> "已记账" Then
            MsgBox "有未记账凭证，请处理后再结账！", , "提示"
            Exit Sub
        End If
    Next i

    '—— 2. 检查“发生额及余额表”是否平衡 ——
    If 核算("发生额及余额表", 8, 0) = 1 Then
        MsgBox "核算不平衡，不能结账！", , "提示"
        Exit Sub
    End If

    If 核算("发生额及余额表", 10, 0) = 1 Then
        MsgBox "核算不平衡，不能结账！", , "提示"
        Exit Sub
    End If

    '—— 3. 检查资产负债是否平衡 ——
    Set zz = Sheets("资产负债表")
    If zz.Cells(19, 2).Value <> zz.Cells(19, 5).Value Then
        MsgBox "资产：" & zz.Cells(19, 2).Value & "  负债：" & zz.Cells(19, 5).Value & Chr(13) & "资产负债不平衡，不能结账！", , "提示"
        Exit Sub
    End If

    '—— 4. 写入结账登记表本月记录 ——
    yy.Range("A" & (n1 + 1)).Value = thisMonth

    MsgBox "核算平衡，结账成功！", , "提示"
End Sub
