VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 系统平台 
   Caption         =   "系统"
   ClientHeight    =   9320.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8390.001
   OleObjectBlob   =   "系统平台.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "系统平台"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
系统平台.Hide
    用户管理.show
End Sub



Private Sub CommandButton10_Click()
凭证处理.show
End Sub

Private Sub CommandButton13_Click()
  Dim sh As Worksheet
    Dim xx As Worksheet
    Dim yy As Worksheet
    Dim zz As Worksheet
    Dim qq As Worksheet
    Dim lastRow As Long
    Dim lastrow1 As Long
    Dim N2 As Long
    
    Dim sig As Integer
    Dim i As Long
    Dim flag As Integer
flag = 0
Set sh = Sheets("凭证汇总")

lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
For i = 2 To lastRow
If sh.Cells(i, 12) = "已签字" Or sh.Cells(i, 12) = "已记账" Then
    flag = 1
Else
    flag = 0
    MsgBox "存在未签字凭证\作废凭证未删除", , "提示": Exit Sub

End If
Next
If flag = 1 Then
Call 函数.记账: MsgBox "所有已签字凭证均记账成功", , "提示"

'更新利润表
    Dim N, mny As Double
    Set xx = Sheets("利润表")
    Set yy = Sheets("发生额及余额表")
    xx.Range("b3:b15").Clear
    lastrow1 = yy.Cells(sh.Rows.count, "A").End(xlUp).Row
    '主营业务收入
    mny = 0
    For i = 2 To lastrow1
        If yy.Range("a" & i) Like "6001*" Or yy.Range("a" & i) = "6051" Or yy.Range("a" & i) = "4103" Then
            mny = mny + yy.Range("I" & i) - yy.Range("H" & i)
        End If
    Next i
    xx.Range("b2") = mny
    '营业成本
    mny = 0
    For i = 2 To lastrow1
        If yy.Range("a" & i) Like "6401*" Or yy.Range("a" & i) = "6052" Then
            mny = mny + yy.Range("I" & i) - yy.Range("H" & i)
        End If
    Next i
    xx.Range("b3") = mny
    '税金及附加
    mny = 0
    For i = 2 To lastrow1
        If yy.Range("a" & i) Like "6403*" Then
            mny = mny + yy.Range("I" & i) - yy.Range("H" & i)
        End If
    Next i
    xx.Range("b4") = mny
    '销售费用
    mny = 0
    For i = 2 To lastrow1
        If yy.Range("a" & i) Like "6601" Then
            mny = mny + yy.Range("I" & i) - yy.Range("H" & i)
        End If
    Next i
    xx.Range("b5") = mny
    '管理费用
    mny = 0
    For i = 2 To lastrow1
        If yy.Range("a" & i) Like "6602" Then
            mny = mny + yy.Range("I" & i) - yy.Range("H" & i)
        End If
    Next i
    xx.Range("b6") = mny
    '财务费用
    mny = 0
    For i = 2 To lastrow1
        If yy.Range("a" & i) Like "6603" Then
            mny = mny + yy.Range("I" & i) - yy.Range("H" & i)
        End If
    Next i
    xx.Range("b7") = mny

    '营业利润
    xx.Range("b8") = xx.Range("b2") - xx.Range("b3") - xx.Range("b4") - xx.Range("b5") - xx.Range("b6") - xx.Range("b7")
    '营业外收入
    mny = 0
    For i = 2 To lastrow1
        If yy.Range("a" & i) Like "6301" Then
           mny = mny + yy.Range("I" & i) - yy.Range("H" & i)
        End If
    Next i
    xx.Range("b9") = mny
    '营业外支出
    mny = 0
    For i = 2 To lastrow1
        If yy.Range("a" & i) Like "6711" Then
            mny = mny + yy.Range("I" & i) - yy.Range("H" & i)
        End If
    Next i
    xx.Range("b10") = mny
    '利润总额
    xx.Range("b11") = xx.Range("b8") + xx.Range("b9") - xx.Range("b10")
    '所得税费用
    mny = 0
    For i = 2 To lastrow1
        If yy.Range("a" & i) Like "6801" Then
            mny = mny + yy.Range("I" & i) - yy.Range("H" & i)
        End If
    Next i
    xx.Range("b12") = mny
    '净利润
    xx.Range("b13") = xx.Range("b11") - xx.Range("b12")
    
    
    
'更新资产负债表
    Set zz = Sheets("资产负债表")
    Set qq = Sheets("期初余额表")
    N2 = qq.Range("a65535").End(xlUp).Row
    '货币资金
    zz.Range("b3:b20").Clear
    zz.Range("d3:d20").Clear
    mny = 0
    sig = 0
    For i = 3 To lastrow1 - 1
        If yy.Range("a" & i) Like "1001*" Or yy.Range("a" & i) Like "1002*" Then
            sig = 1
            mny = mny + yy.Range("j" & i) - yy.Range("k" & i)
        End If
    Next i
    If sig = 0 Then
        For i = 4 To N2
            If qq.Range("a" & i) Like "1001*" Or Sheet6.Range("a" & i) Like "1002*" Then
                mny = mny + qq.Range("f" & i) - qq.Range("g" & i)
            End If
        Next i
    End If
    zz.Range("b3") = mny
    '应收票据
    mny = 0
    sig = 0
    For i = 3 To lastrow1 - 1
        If yy.Range("a" & i) Like "1121" Then
            sig = 1
            mny = mny + yy.Range("j" & i) - yy.Range("k" & i)
        End If
    Next i
    If sig = 0 Then
        For i = 4 To N2
            If qq.Range("a" & i) Like "1121" Then
                mny = mny + qq.Range("f" & i) - qq.Range("g" & i)
            End If
        Next i
    End If
    zz.Range("b4") = mny
    '应收账款
    mny = 0
    sig = 0
    For i = 3 To lastrow1 - 1
        If yy.Range("a" & i) Like "1122*" Or yy.Range("a" & i) Like "1221*" Then
            sig = 1
            mny = mny + yy.Range("j" & i) - yy.Range("k" & i)
        End If
    Next i
    If sig = 0 Then
        For i = 4 To N2
            If qq.Range("a" & i) Like "1122" Then
                mny = mny + qq.Range("f" & i) - qq.Range("g" & i)
            End If
        Next i
    End If
    zz.Range("b5") = mny
    '存货
    mny = 0
    sig = 0
    For i = 3 To lastrow1 - 1
        If yy.Range("a" & i) = "1405" Or yy.Range("a" & i) Like "1403*" Then
            sig = 1
             mny = mny + yy.Range("j" & i) - yy.Range("k" & i)
        End If
    Next i
    If sig = 0 Then
        For i = 4 To N2
            If qq.Range("a" & i) = "1405" Or qq.Range("a" & i) Like "1403*" Then
                mny = mny + qq.Range("f" & i) - qq.Range("g" & i)
            End If
        Next i
    End If
    zz.Range("b6") = mny

    '流动资产合计
    zz.Range("b7") = zz.Range("b3") + zz.Range("b4") + zz.Range("b5") + zz.Range("b6")
    '固定资产
    mny = 0
    sig = 0
    For i = 3 To lastrow1 - 1
        If yy.Range("a" & i) Like "1602*" Or yy.Range("a" & i) Like "1601*" Or yy.Range("a" & i) Like "1801*" Then
            sig = 1
            mny = mny + yy.Range("j" & i) - yy.Range("k" & i)
        End If
    Next i
    If sig = 0 Then
        For i = 4 To N2
            If qq.Range("a" & i) Like "1602*" Or qq.Range("a" & i) Like "1601*" Then ''''''''长期待摊费用其实应该另开一行'''''''
                mny = mny + qq.Range("f" & i) - qq.Range("g" & i)
            End If
        Next i
    End If
    zz.Range("b9") = mny
    '无形资产
    mny = 0
    sig = 0
    For i = 3 To lastrow1 - 1
        If yy.Range("a" & i) Like "1701*" Or yy.Range("a" & i) Like "1702*" Then
            sig = 1
            mny = mny + yy.Range("j" & i) - yy.Range("k" & i)
        End If
    Next i
    If sig = 0 Then
        For i = 4 To N2
            If qq.Range("a" & i) Like "1701*" Or qq.Range("a" & i) Like "1702*" Then
                mny = mny + qq.Range("f" & i) - qq.Range("g" & i)
            End If
        Next i
    End If
    zz.Range("b10") = mny
    '非流动资产合计
    zz.Range("b11") = Sheet13.Range("b10") + Sheet13.Range("b9")
    '短期借款
    mny = 0
    sig = 0
    For i = 3 To lastrow1 - 1
        If yy.Range("a" & i) Like "2001*" Then
            sig = 1
            mny = mny - yy.Range("j" & i) + yy.Range("k" & i)
        End If
    Next i
    If sig = 0 Then
        For i = 4 To N2
            If qq.Range("a" & i) Like "2001*" Then
                mny = mny - qq.Range("f" & i) + qq.Range("g" & i)
            End If
        Next i
    End If
    zz.Range("d3") = mny
    '应付票据
    mny = 0
    sig = 0
    For i = 3 To lastrow1 - 1
        If yy.Range("a" & i) = "2201" Then
            sig = 1
            mny = mny - yy.Range("j" & i) + yy.Range("k" & i)
        End If
    Next i
    If sig = 0 Then
        For i = 4 To N2
            If qq.Range("a" & i) = "2201" Then
                mny = mny - qq.Range("f" & i) + qq.Range("g" & i)
            End If
        Next i
    End If
    zz.Range("d4") = mny
    '应付账款
    mny = 0
    sig = 0
    For i = 3 To N - 1
        If yy.Range("a" & i) = "2202" Then
            sig = 1
            mny = mny - yy.Range("j" & i) + yy.Range("k" & i)
        End If
    Next i
    If sig = 0 Then
        For i = 4 To N2
            If qq.Range("a" & i) = "2202" Then
                mny = mny - qq.Range("f" & i) + qq.Range("g" & i)
            End If
        Next i
    End If
    zz.Range("d5") = mny
    '应付职工薪资
    mny = 0
    sig = 0
    For i = 3 To lastrow1 - 1
        If yy.Range("a" & i) = "2211" Then
            sig = 1
            mny = mny - yy.Range("j" & i) + yy.Range("k" & i)
        End If
    Next i
    If sig = 0 Then
        For i = 4 To N2
            If qq.Range("a" & i) = "2211" Then
                mny = mny - qq.Range("f" & i) + qq.Range("g" & i)
            End If
        Next i
    End If
    zz.Range("d6") = mny
    '应付税费
    mny = 0
    sig = 0
    For i = 3 To lastrow1 - 1
        If yy.Range("a" & i) Like "2221*" Then
            sig = 1
            mny = mny - yy.Range("j" & i) + yy.Range("k" & i)
        End If
    Next i
    If sig = 0 Then
        For i = 4 To N2
            If qq.Range("a" & i) Like "2221*" Then
                mny = mny - qq.Range("f" & i) + qq.Range("g" & i)
            End If
        Next i
    End If
    zz.Range("d7") = mny
    '其他应付款
    mny = 0
    sig = 0
    For i = 3 To lastrow1 - 1
        If yy.Range("a" & i) = "2241" Then
            sig = 1
            mny = mny - yy.Range("j" & i) + yy.Range("k" & i)
        End If
    Next i
    If sig = 0 Then
        For i = 4 To N2
            If qq.Range("a" & i) = "2241" Then
                mny = mny - qq.Range("f" & i) + qq.Range("g" & i)
            End If
        Next i
    End If
    zz.Range("d8") = mny
    '流动负债合计
    zz.Range("d9") = zz.Range("d3") + zz.Range("d4") + zz.Range("d5") + zz.Range("d6") + zz.Range("d7") + zz.Range("d8")
    '长期借款
    mny = 0
    sig = 0
    For i = 3 To lastrow1 - 1
        If yy.Range("a" & i) = "2501" Then
            sig = 1
            mny = mny - yy.Range("j" & i) + yy.Range("k" & i)
        End If
    Next i
    If sig = 0 Then
        For i = 4 To N2
            If qq.Range("a" & i) = "2501" Then
                mny = mny - qq.Range("f" & i) + qq.Range("g" & i)
            End If
        Next i
    End If
    zz.Range("d11") = mny
    zz.Range("d12") = mny
    '负债合计
    zz.Range("d13") = Sheet13.Range("d12") + Sheet13.Range("d9")
    '实收资本
    mny = 0
    sig = 0
    For i = 3 To lastrow1 - 1
        If yy.Range("a" & i) = "4001" Then
            sig = 1
            mny = mny - yy.Range("j" & i) + yy.Range("k" & i)
        End If
    Next i
    If sig = 0 Then
        For i = 4 To N2
            If qq.Range("a" & i) = "4001" Then
                mny = mny - qq.Range("f" & i) + qq.Range("g" & i)
            End If
        Next i
    End If
    zz.Range("d15") = mny
    '资本公积
    mny = 0
    sig = 0
    For i = 3 To lastrow1 - 1
        If yy.Range("a" & i) = "4002" Then
            sig = 1
            mny = mny - yy.Range("j" & i) + yy.Range("k" & i)
        End If
    Next i
    If sig = 0 Then
        For i = 4 To N2
            If qq.Range("a" & i) = "4002" Then
                mny = mny - qq.Range("f" & i) + qq.Range("g" & i)
            End If
        Next i
    End If
    zz.Range("d16") = mny
    '盈余公积
    mny = 0
    sig = 0
    For i = 3 To lastrow1 - 1
        If yy.Range("a" & i) = "4101" Then
            sig = 1
            mny = mny - yy.Range("j" & i) + yy.Range("k" & i)
        End If
    Next i
    If sig = 0 Then
        For i = 4 To N2
            If qq.Range("a" & i) = "4101" Then
                mny = mny - qq.Range("f" & i) + qq.Range("g" & i)
            End If
        Next i
    End If
    zz.Range("d17") = mny
    '未分配利润
    mny = 0
    sig = 0
    For i = 3 To lastrow1 - 1
        If yy.Range("a" & i) Like "4103*" Then
            sig = 1
            mny = mny - yy.Range("j" & i) + yy.Range("k" & i)
        End If
    Next i
    If sig = 0 Then
        For i = 4 To N2
            If qq.Range("a" & i) = "4101" Then
                mny = mny - qq.Range("f" & i) + qq.Range("g" & i)
            End If
        Next i
    End If
    zz.Range("d18") = mny
    '所有者权益合计
    zz.Range("d19") = zz.Range("d15") + zz.Range("d16") + zz.Range("d17") + zz.Range("d18")
    '资产合计
    zz.Range("b20") = zz.Range("b7") + zz.Range("b11")
    '负债和所有者权益总计
    zz.Range("d20") = zz.Range("d19") + zz.Range("d13")

End If

End Sub



Private Sub CommandButton14_Click()
    Dim response      As VbMsgBoxResult
    Dim wsChk         As Worksheet
    Dim wsPz          As Worksheet
    Dim wsPL          As Worksheet
    Dim wsBS          As Worksheet
    Dim lastRowChk    As Long
    Dim lastRowPz     As Long
    Dim lastAssets    As Long
    Dim lastLiab      As Long
    Dim lastRowBS     As Long
    Dim lastRowPL     As Long
    Dim i             As Long
    Dim thisMonth     As Integer
    Dim foundClosed   As Boolean
    Dim hasPosted     As Boolean      ' ?? 新增：用于判断是否存在已记账凭证
    
    '—— 1. 先让用户确认是否要取消记账 ——
    response = MsgBox("确定要取消记账吗？", vbYesNo + vbQuestion, "确认取消记账")
    If response = vbNo Then Exit Sub
    
    '—— 2. 检查“结账登记表”里是否已包含当前月份 ——
    thisMonth = Month(Date)    ' 当前月份数字 (1–12)
    Set wsChk = ThisWorkbook.Sheets("结账登记表")
    
    lastRowChk = wsChk.Cells(wsChk.Rows.count, "A").End(xlUp).Row
    foundClosed = False
    
    For i = 2 To lastRowChk
        If wsChk.Cells(i, "A").Value = thisMonth Then
            foundClosed = True
            Exit For
        End If
    Next i
    
    If foundClosed Then
        MsgBox "本月已结账，无法取消记账！", vbExclamation, "提示"
        Exit Sub
    End If
    
    '—— ?? 2.5 检查是否存在“已记账”的凭证 ——
    Set wsPz = ThisWorkbook.Sheets("凭证汇总")
    lastRowPz = wsPz.Cells(wsPz.Rows.count, "A").End(xlUp).Row
    hasPosted = False
    
    For i = 2 To lastRowPz
        If wsPz.Cells(i, "L").Value = "已记账" Then
            hasPosted = True
            Exit For
        End If
    Next i
    
    If Not hasPosted Then
        MsgBox "当前没有已记账凭证，无法执行取消记账操作。", vbExclamation, "提示"
        Exit Sub
    End If
    
    '—— 3. 清空“凭证汇总”表里已记账标记 ——
    For i = 2 To lastRowPz
        If wsPz.Cells(i, "L").Value = "已记账" Then
            wsPz.Cells(i, "L").Value = "已签字"
        End If
    Next i
    
    '—— 4. 清空“利润表”里“本期金额”——
    Set wsPL = ThisWorkbook.Sheets("利润表")
    lastRowPL = wsPL.Cells(wsPL.Rows.count, "B").End(xlUp).Row
    If lastRowPL >= 2 Then
        wsPL.Range("B2:B" & lastRowPL).ClearContents
    End If
    
    '—— 5. 清空“资产负债表”里“余额”列——
    Set wsBS = ThisWorkbook.Sheets("资产负债表")
    lastAssets = wsBS.Cells(wsBS.Rows.count, "A").End(xlUp).Row
    lastLiab = wsBS.Cells(wsBS.Rows.count, "C").End(xlUp).Row
    lastRowBS = Application.WorksheetFunction.Max(lastAssets, lastLiab)
    
    If lastRowBS >= 2 Then
        wsBS.Range("B2:B" & lastRowBS).ClearContents
        wsBS.Range("D2:D" & lastRowBS).ClearContents
    End If
    
    MsgBox "已成功取消记账！", vbInformation, "提示"
End Sub



Private Sub CommandButton15_Click()
Call 函数.结账
End Sub

Private Sub CommandButton16_Click()
    Dim sh As Worksheet
    Dim lastRow As Long, i As Long
    Dim count As Integer
    Dim wsChk As Worksheet
    Dim lastRowChk As Long
    Dim thisMonth As Integer

    ' 提示用户是否确认取消结账操作
    If MsgBox("是否确认取消结账？此操作将撤销所有'已记账'凭证状态，并清空利润表和资产负债表数据。", vbYesNo + vbQuestion, "确认取消结账") = vbNo Then Exit Sub

    ' 取消凭证记账状态（将“已记账”改回“已签字”）
    Set sh = Sheets("凭证汇总")
    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    count = 0

    For i = 2 To lastRow
        If sh.Cells(i, 12).Value = "已记账" Then
            sh.Cells(i, 12).Value = "已签字" ' 或根据需求改为 "未签字"
            count = count + 1
        End If
    Next i

    MsgBox count & " 条凭证状态已取消记账，恢复为“已签字”。", vbInformation, "凭证状态重置"

    ' 清空利润表数据
    With Sheets("利润表")
        .Range("B2:B13").ClearContents
    End With

    ' 清空资产负债表数据
    With Sheets("资产负债表")
        .Range("B3:B20").ClearContents ' 资产项
        .Range("D3:D20").ClearContents ' 负债与权益项
    End With

    ' 删除“结账登记表”中本月记录
    Set wsChk = Sheets("结账登记表")
    lastRowChk = wsChk.Cells(wsChk.Rows.count, "A").End(xlUp).Row
    thisMonth = Month(Date)

    For i = lastRowChk To 2 Step -1
        If wsChk.Cells(i, "A").Value = thisMonth Then
            wsChk.Rows(i).Delete
        End If
    Next i

    MsgBox "已清空利润表与资产负债表，取消结账成功完成。", vbInformation, "取消结账完成"
End Sub





Private Sub CommandButton2_Click()

    ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "/账套备份" & Format(Now(), "yyyymmddhhmmss") & ".xls"
    Dim filename As String
    MsgBox "账套导出成功！请到根目录查看账套备份", 1 + 64, "导出提示"
End Sub



Private Sub CommandButton21_Click()
查看账表.ymsy.Visible = True
查看账表.show
End Sub


Private Sub CommandButton23_Click()
 Unload Me
    登录页面.show
End Sub

Private Sub CommandButton24_Click()
系统平台.Hide
End Sub

Private Sub CommandButton3_Click()
系统平台.Hide
    部门管理.show
End Sub

Private Sub CommandButton4_Click()
系统平台.Hide
    人员管理.show
End Sub

Private Sub CommandButton5_Click()
 系统平台.Hide
    客户管理.show
End Sub

Private Sub CommandButton6_Click()
系统平台.Hide
    供应商管理.show
End Sub

Private Sub CommandButton7_Click()
系统平台.Hide
    会计科目.show
End Sub

Private Sub CommandButton8_Click()
系统平台.Hide
    期初余额.show
End Sub

Private Sub CommandButton9_Click()
 Dim sh As Worksheet  ' 确保 sh 被声明为 Worksheet 对象
    Dim lastRow As Long
     Dim i As Long
Set sh = Sheets("结账登记表")
lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
For i = 2 To lastRow
If sh.Cells(i, 1) = Month(Date) Then MsgBox "本月已结账，不能再填制凭证", , "提示": Exit Sub
Next
系统平台.Hide
    凭证填制.show
End Sub

Private Sub UserForm_Activate()
    Dim ws As Worksheet
    Dim role As String

    ' 确保 currentUser 已被设置
    If currentUser <= 0 Then
        MsgBox "当前用户未登录或登录失败！", vbExclamation
        Exit Sub
    End If

    Set ws = Sheets("用户档案")
    On Error Resume Next
    role = ws.Cells(currentUser, "D").Value
    On Error GoTo 0

    ' 检查是否能正确获取角色
    If role = "" Then
        MsgBox "未找到角色信息！", vbExclamation
        Exit Sub
    End If

    ' 设置用户信息显示
    Me.TextBox1.text = ws.Cells(currentUser, "B").Value  ' 用户名
    Me.TextBox2.text = role                              ' 角色
    Me.TextBox3.text = Format(Date, "yyyy-mm-dd")        ' 当前日期

    ' 默认禁用所有页面与按钮
    Dim i As Integer
    For i = 0 To MultiPage1.Pages.count - 1
        MultiPage1.Pages(i).Enabled = False
    Next i

    Dim btn As Control
    For Each btn In Me.Controls
        If TypeName(btn) = "CommandButton" Then btn.Enabled = False
    Next btn

    ' 根据角色分配权限
    Select Case role
        Case "管理员"
            MultiPage1.Page1.Enabled = True
            CommandButton1.Enabled = True  ' 用户管理
            CommandButton2.Enabled = True  ' 导出账套

        Case "账套主管"
            MultiPage1.Page2.Enabled = True
            MultiPage1.Page3.Enabled = True
        '    MultiPage1.Page4.Enabled = True
            MultiPage1.Page5.Enabled = True

            ' 基础设置
            CommandButton3.Enabled = True  ' 部门管理
            CommandButton4.Enabled = True  ' 人员管理
            CommandButton5.Enabled = True  ' 客户管理
            CommandButton6.Enabled = True  ' 供应商管理
            CommandButton7.Enabled = True  ' 会计科目
            CommandButton8.Enabled = True  ' 期初余额

            ' 业务处理
            CommandButton9.Enabled = True   ' 填制凭证
            CommandButton10.Enabled = True  ' 处理凭证
            CommandButton13.Enabled = True  ' 记账
            CommandButton14.Enabled = True  ' 取消记账
            CommandButton15.Enabled = True  ' 结账
            CommandButton16.Enabled = True  ' 取消结账

            ' 账表查询
         '   CommandButton17.Enabled = True  ' 总账
         ''   CommandButton18.Enabled = True  ' 明细账
         '   CommandButton19.Enabled = True  ' 日记账
         '   CommandButton20.Enabled = True  ' 发生额及余额表

            ' 报表查询
            CommandButton21.Enabled = True  ' 资产负债表
      '      CommandButton22.Enabled = True  ' 利润表

        Case "会计人员"
            MultiPage1.Page3.Enabled = True
       '     MultiPage1.Page4.Enabled = True
            MultiPage1.Page5.Enabled = True

            CommandButton9.Enabled = True   ' 填制凭证
            CommandButton10.Enabled = True  ' 处理凭证

        '    CommandButton17.Enabled = True  ' 总账
        '    CommandButton18.Enabled = True  ' 明细账
        '    CommandButton19.Enabled = True  ' 日记账
         '   CommandButton20.Enabled = True  ' 发生额及余额表
            CommandButton21.Enabled = True  ' 资产负债表
         '   CommandButton22.Enabled = True  ' 利润表

        Case "出纳人员"
            MultiPage1.Page3.Enabled = True
      '      MultiPage1.Page4.Enabled = True
            MultiPage1.Page5.Enabled = True

            CommandButton10.Enabled = True  ' 处理凭证

          '  CommandButton17.Enabled = True
          '  CommandButton18.Enabled = True
          '  CommandButton19.Enabled = True
         '   CommandButton20.Enabled = True
            CommandButton21.Enabled = True
        '    CommandButton22.Enabled = True

        Case "普通用户"
        '    MultiPage1.Page4.Enabled = True
            MultiPage1.Page5.Enabled = True

         '   CommandButton17.Enabled = True
         '   CommandButton18.Enabled = True
         ''   CommandButton19.Enabled = True
         '   CommandButton20.Enabled = True
            CommandButton21.Enabled = True
    '        CommandButton22.Enabled = True

        Case Else
            MsgBox "角色信息无效或未授权！", vbCritical
    End Select
  
For i = 0 To MultiPage1.Pages.count - 1
    If MultiPage1.Pages(i).Enabled = True Then
        MultiPage1.Value = i
        Exit For
    End If
Next i
' 所有人共用功能启用
    CommandButton23.Enabled = True  ' 重注册
    CommandButton24.Enabled = True  ' 退出
End Sub



