VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 查看账表 
   Caption         =   "查看账表"
   ClientHeight    =   8940.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   16060
   OleObjectBlob   =   "查看账表.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "查看账表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()

Set sh = Sheets("会计科目表")
lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
ComboBox2.Clear

Select Case ComboBox1.Value

Case "总账"

For i = 2 To lastRow
If sh.Cells(i, 1) = sh.Cells(i, 2) Then ComboBox2.AddItem sh.Cells(i, 4)
Next

Case "明细账"
For i = 2 To lastRow - 1
If sh.Cells(i, 1) = sh.Cells(i, 2) And sh.Cells(i, 2) = sh.Cells(i + 1, 2) Then '不记有子科目的一级科目
Else
ComboBox2.AddItem sh.Cells(i, 4)

End If

Next

ComboBox2.AddItem sh.Cells(lastRow, 4)

Case "现金日记账"

For i = 2 To lastRow

If sh.Cells(i, 7) = "Y" And sh.Cells(i, 8) = "" Then ComboBox2.AddItem sh.Cells(i, 4)
Next

Case "银行日记账"

For i = 2 To lastRow

If sh.Cells(i, 7) = "Y" And sh.Cells(i, 8) = "Y" Then ComboBox2.AddItem sh.Cells(i, 4)

Next
End Select
End Sub

Private Sub CommandButton1_Click()

zb = ComboBox1.text: km = ComboBox2.text '取得账表和科目

ListView1.ColumnHeaders.Clear: ListView1.ListItems.Clear '清空列表

If zb = "" Then MsgBox "未选择账表", , "提示": Exit Sub

If zb = "资产负债表" Or zb = "利润表" Then

    Set sh = Sheets(zb)
    
    lastColumn = sh.Cells(1, sh.Columns.count).End(xlToLeft).Column
    
    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    
    For i = 1 To lastColumn
    
    ListView1.ColumnHeaders.Add i, , sh.Cells(1, i), ListView1.Width / 4
    
    Next
    
    ListView1.View = lvwReport: ListView1.FullRowSelect = True: ListView1.Gridlines = True
    
    For i = 2 To lastRow
    
    Set itm = ListView1.ListItems.Add
    
    itm.text = sh.Cells(i, 1)
    
    For j = 2 To lastColumn
    
    itm.SubItems(j - 1) = sh.Cells(i, j)
    
    Next
    
    Next
    
    Exit Sub

End If

If zb <> "发生额及余额表" And km = "" Then MsgBox "未选择科目", , "提示": Exit Sub



If zb = "发生额及余额表" Then

Set sh = Sheets(zb)

lastColumn = sh.Cells(1, sh.Columns.count).End(xlToLeft).Column
    
lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row


ListView1.ColumnHeaders.Add 1, , sh.Cells(1, 1), ListView1.Width / 8

ListView1.ColumnHeaders.Add 2, , sh.Cells(1, 4), ListView1.Width / 8

For i = 6 To lastColumn

    ListView1.ColumnHeaders.Add i - 3, , sh.Cells(1, i), ListView1.Width / 8

Next

ListView1.View = lvwReport: ListView1.FullRowSelect = True: ListView1.Gridlines = True
    '用于存储合计的
    '借/贷期初余额
    sj = 0: sd = 0
    '借/贷发生额
    jf = 0: df = 0
    '借/贷期末余额
    fj = 0: fd = 0

    '用于存储小计的，每小计一次后重置
    xsj = 0: xsd = 0
    
    xjf = 0: xdf = 0
    
    xfj = 0: xfd = 0

biao = 1

For i = 2 To lastRow

If Val(Left(sh.Cells(i, 1), 1)) > biao And xsj + xsd + xjf + xdf + xfj + xfd > 0 Then

Set itm = ListView1.ListItems.Add

Select Case biao

Case "1"

itm.SubItems(1) = "资产小计"

Case "2"

itm.SubItems(1) = "负债小计"

Case "3"

itm.SubItems(1) = "共同小计"

Case "4"

itm.SubItems(1) = "权益小计"

Case "5"

itm.SubItems(1) = "成本小计"

End Select

itm.SubItems(2) = xsj: itm.SubItems(3) = xsd: itm.SubItems(4) = xjf: itm.SubItems(5) = xdf

itm.SubItems(6) = xfj: itm.SubItems(7) = xfd

xsj = 0: xsd = 0: xjf = 0: xdf = 0: xfj = 0: xfd = 0

biao = Left(sh.Cells(i, 1), 1)

End If

If sh.Cells(i, 6) > 0 Or sh.Cells(i, 7) > 0 Or sh.Cells(i, 8) > 0 Or sh.Cells(i, 9) > 0 Then

biao = Left(sh.Cells(i, 1), 1)

If sh.Cells(i, 1) = sh.Cells(i, 2) Then '如果是一级科目，只算一级科目的总和

    sj = sj + sh.Cells(i, 6): sd = sd + sh.Cells(i, 7): jf = jf + sh.Cells(i, 8): df = df + sh.Cells(i, 9)
    
    fj = fj + sh.Cells(i, 10): fd = fd + sh.Cells(i, 11)
    
    xsj = xsj + sh.Cells(i, 6): xsd = xsd + sh.Cells(i, 7): xjf = xjf + sh.Cells(i, 8): xdf = xdf + sh.Cells(i, 9)
    
    xfj = xfj + sh.Cells(i, 10): xfd = xfd + sh.Cells(i, 11)

End If

Set itm = ListView1.ListItems.Add

itm.text = sh.Cells(i, 1): itm.SubItems(1) = sh.Cells(i, 4)

For j = 6 To lastColumn

itm.SubItems(j - 4) = sh.Cells(i, j)

Next

End If

Next


Set itm = ListView1.ListItems.Add
'6开头为损益类
itm.SubItems(1) = "损益小计"

itm.SubItems(2) = xsj: itm.SubItems(3) = xsd: itm.SubItems(4) = xjf: itm.SubItems(5) = xdf

itm.SubItems(6) = xfj: itm.SubItems(7) = xfd

Set itm = ListView1.ListItems.Add

itm.SubItems(1) = "合计"

itm.SubItems(2) = sj: itm.SubItems(3) = sd: itm.SubItems(4) = jf: itm.SubItems(5) = df

itm.SubItems(6) = fj: itm.SubItems(7) = fd

Exit Sub

End If

'总账

If zb = "总账" Then
Set sh = Sheets(zb)
lastColumn = sh.Cells(1, sh.Columns.count).End(xlToLeft).Column
    
lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row


For i = 5 To lastColumn

ListView1.ColumnHeaders.Add i - 4, , sh.Cells(1, i), ListView1.Width / 5 '可设宽度

Next

ListView1.View = lvwReport: ListView1.FullRowSelect = True: ListView1.Gridlines = True

Set zz = Sheets("发生额及余额表")

'找到匹配的科目名字
s = WorksheetFunction.Match(km, zz.[D:D], 0)

    Set itm = ListView1.ListItems.Add

    itm.text = "期初余额": itm.SubItems(1) = zz.Cells(s, 6): itm.SubItems(2) = zz.Cells(s, 7)
    
    If zz.Cells(s, 6) - zz.Cells(s, 7) > 0 Then itm.SubItems(3) = "借": itm.SubItems(4) = zz.Cells(s, 6) - zz.Cells(s, 7)
    
    If zz.Cells(s, 7) - zz.Cells(s, 6) > 0 Then itm.SubItems(3) = "贷": itm.SubItems(4) = zz.Cells(s, 7) - zz.Cells(s, 6)
    
    If zz.Cells(s, 6) = 0 And zz.Cells(s, 7) = 0 Then itm.SubItems(3) = "平": itm.SubItems(4) = 0


Set itm = ListView1.ListItems.Add

    itm.text = "发生额": itm.SubItems(1) = zz.Cells(s, 8): itm.SubItems(2) = zz.Cells(s, 9)
    
    If zz.Cells(s, 8) - zz.Cells(s, 9) > 0 Then itm.SubItems(3) = "借": itm.SubItems(4) = zz.Cells(s, 8) - zz.Cells(s, 9)
    
    If zz.Cells(s, 9) - zz.Cells(s, 8) > 0 Then itm.SubItems(3) = "贷": itm.SubItems(4) = zz.Cells(s, 9) - zz.Cells(s, 8)
    
    If zz.Cells(s, 8) = 0 And zz.Cells(s, 9) = 0 Then itm.SubItems(3) = "平": itm.SubItems(4) = 0

Set itm = ListView1.ListItems.Add

    itm.text = "期末余额": itm.SubItems(1) = zz.Cells(s, 10): itm.SubItems(2) = zz.Cells(s, 11)
    
    If zz.Cells(s, 10) - zz.Cells(s, 11) > 0 Then itm.SubItems(3) = "借": itm.SubItems(4) = zz.Cells(s, 10) - zz.Cells(s, 11)
    
    If zz.Cells(s, 11) - zz.Cells(s, 10) > 0 Then itm.SubItems(3) = "贷": itm.SubItems(4) = zz.Cells(s, 11) - zz.Cells(s, 10)
    
    If zz.Cells(s, 10) = 0 And zz.Cells(s, 11) = 0 Then itm.SubItems(3) = "平": itm.SubItems(4) = 0
Set itm = ListView1.ListItems.Add

    itm.text = "合计"
    sumj = Val(zz.Cells(s, 6)) + Val(zz.Cells(s, 8)) + Val(zz.Cells(s, 10))
    sumd = Val(zz.Cells(s, 7)) + Val(zz.Cells(s, 9)) + Val(zz.Cells(s, 11))
    itm.SubItems(1) = sumj
    itm.SubItems(2) = sumd
    
    If sumj - sumd > 0 Then itm.SubItems(3) = "借": itm.SubItems(4) = sumj - sumd
    
    If sumd - sumj > 0 Then itm.SubItems(3) = "贷": itm.SubItems(4) = sumd - sumj
    
    If sumj - sumd = 0 Then itm.SubItems(3) = "平": itm.SubItems(4) = 0

Exit Sub

End If

'明细账
If zb = "明细账" Or zb = "现金日记账" Or zb = "银行日记账" Then
Set sh = Sheets(zb)
lastColumn = sh.Cells(1, sh.Columns.count).End(xlToLeft).Column
    
lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row

For i = 1 To lastColumn

ListView1.ColumnHeaders.Add i, , sh.Cells(1, i), ListView1.Width / 9 '可设宽度

Next

ListView1.View = lvwReport: ListView1.FullRowSelect = True: ListView1.Gridlines = True

Set zz = Sheets("发生额及余额表")

s = WorksheetFunction.Match(km, zz.[D:D], 0)

Set itm = ListView1.ListItems.Add

    itm.SubItems(4) = "期初余额": itm.SubItems(5) = zz.Cells(s, 6): itm.SubItems(6) = zz.Cells(s, 7)
    
    If zz.Cells(s, 6) - zz.Cells(s, 7) > 0 Then itm.SubItems(7) = "借": itm.SubItems(8) = zz.Cells(s, 6) - zz.Cells(s, 7)
    
    If zz.Cells(s, 7) - zz.Cells(s, 6) > 0 Then itm.SubItems(7) = "贷": itm.SubItems(8) = zz.Cells(s, 7) - zz.Cells(s, 6)
    
    If zz.Cells(s, 6) = 0 And zz.Cells(s, 7) = 0 Then itm.SubItems(7) = "平": itm.SubItems(8) = 0
    '明细的借贷合计
    mxhj = zz.Cells(s, 6): mxhd = zz.Cells(s, 7)

For i = 2 To lastRow
'寻找选择会计科目
    If sh.Cells(i, 1) = km Or km = "银行存款" Then

    Set itm = ListView1.ListItems.Add
    
    itm.text = sh.Cells(i, 1)
    
    mxhj = mxhj + sh.Cells(i, 6): mxhd = mxhd + sh.Cells(i, 7)

        For j = 2 To lastColumn - 1
        
            itm.SubItems(j - 1) = sh.Cells(i, j)
        
        Next
    If sh.Cells(i, 6) - sh.Cells(i, 7) > 0 Then itm.SubItems(lastColumn - 1) = sh.Cells(i, 6) - sh.Cells(i, 7)
    If sh.Cells(i, 7) - sh.Cells(i, 6) > 0 Then itm.SubItems(lastColumn - 1) = sh.Cells(i, 7) - sh.Cells(i, 6)
    If sh.Cells(i, 6) = 0 And sh.Cells(i, 7) = 0 Then itm.SubItems(lastColumn - 1) = 0
    
    End If

Next

Set itm = ListView1.ListItems.Add

itm.SubItems(4) = "期末余额": itm.SubItems(5) = zz.Cells(s, 10): itm.SubItems(6) = zz.Cells(s, 11)
    
    If zz.Cells(s, 10) - zz.Cells(s, 11) > 0 Then itm.SubItems(7) = "借": itm.SubItems(8) = zz.Cells(s, 10) - zz.Cells(s, 11)
    
    If zz.Cells(s, 11) - zz.Cells(s, 10) > 0 Then itm.SubItems(7) = "贷": itm.SubItems(8) = zz.Cells(s, 11) - zz.Cells(s, 10)
    
    If zz.Cells(s, 10) = 0 And zz.Cells(s, 11) = 0 Then itm.SubItems(7) = "平": itm.SubItems(8) = 0
    
    mxhj = mxhj + sh.Cells(i, 10): mxhd = mxhd + sh.Cells(i, 11)
Set itm = ListView1.ListItems.Add

    itm.SubItems(4) = "合计": itm.SubItems(5) = mxhj: itm.SubItems(6) = mxhd
    
    If mxhj - mxhd > 0 Then itm.SubItems(7) = "借": itm.SubItems(8) = mxhj - mxhd
    
    If mxhd - mxhj > 0 Then itm.SubItems(7) = "贷": itm.SubItems(8) = mxhd - mxhj
    
    If mxhj - mxhd = 0 Then itm.SubItems(7) = "平": itm.SubItems(8) = 0



End If
End Sub

Private Sub CommandButton4_Click()
Unload Me
End Sub


Private Sub UserForm_Activate()

ListView1.ColumnHeaders.Clear:

ListView1.ListItems.Clear '清空列表

ComboBox2.AddItem ""

ComboBox1.AddItem "总账": ComboBox1.AddItem "明细账"

ComboBox1.AddItem "现金日记账": ComboBox1.AddItem "银行日记账": ComboBox1.AddItem "发生额及余额表"

ComboBox1.AddItem "资产负债表": ComboBox1.AddItem "利润表"

End Sub

Private Sub ymsy_Click()
Call 函数.利润结转
End Sub

