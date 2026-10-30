VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 期初余额 
   Caption         =   "期初余额"
   ClientHeight    =   9140.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   16560
   OleObjectBlob   =   "期初余额.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "期初余额"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton3_Click()
If 函数.核算("期初余额表(临时)", 6, 1) = 1 Then MsgBox "试算不平衡", , "提示" Else MsgBox "试算平衡", , "提示"

End Sub

Private Sub CommandButton4_Click()
Unload Me
系统平台.show
End Sub

Private Sub CommandButton5_Click()

jfy = TextBox4.text: dfy = TextBox5.text

If km1.Caption = "" Then MsgBox "未选择项目", , "提示": Exit Sub

Set sh = Sheets("期初余额表(临时)") '显示期初余额信息列表

s = WorksheetFunction.Match(km1.Caption, sh.[A:A], 0) '查找科目代码

If sh.Cells(s, 1) = sh.Cells(s, 2) And sh.Cells(s + 1, 2) = sh.Cells(s, 2) Then

MsgBox "该总科目有子科目，请调整子科目", , "提示": Exit Sub

End If

sh.Cells(s, 6) = CStr(jfy): sh.Cells(s, 7) = CStr(dfy)

MsgBox "请进行试算，试算平衡后请确定调整"
If sh.Cells(s, 3) <> "" Then

Call 函数.更新一级科目("期初余额表(临时)", 6)

Call 函数.更新一级科目("期初余额表(临时)", 7)
End If
Call Reflush
End Sub

Private Sub CommandButton6_Click()
If 函数.核算("期初余额表(临时)", 6, 1) = 1 Then MsgBox "试算不平衡,不能调整", , "提示": Exit Sub

Set sh1 = Sheets("期初余额表"): Set sh2 = Sheets("期初余额表(临时)")

sh2.Range("A1:G65535").Copy

sh1.Range("A1:G65535").PasteSpecial xlPasteAll

Sheets("期初余额表").Range("F2:G65535").Copy Sheets("发生额及余额表").Range("F2:G65535")

Call 更新期末余额

MsgBox "试算平衡,调整成功", , "提示"

End Sub

Private Sub ListView1_Click()

km1.Caption = ListView1.ListItems(ListView1.SelectedItem.Index).text

km2.Caption = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1)

End Sub

Private Sub TextBox1_Change()
If Len(TextBox1) >= 1 Then
        Call serch
        
    Else
        Call Reflush
    End If
End Sub

Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' 检查输入的字符是否为数字
    If KeyAscii < 48 Or KeyAscii > 57 Then
        ' 如果不是数字，则阻止该字符输入
        KeyAscii = 0
        MsgBox ("请输入数字！"), , "提示"
    End If
End Sub

Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' 检查输入的字符是否为数字
    If KeyAscii < 48 Or KeyAscii > 57 Then
        ' 如果不是数字，则阻止该字符输入
        KeyAscii = 0
        MsgBox ("请输入数字！"), , "提示"
    End If
End Sub
Private Sub UserForm_Activate()
  
Set sh1 = ThisWorkbook.Sheets("期初余额表")
Set sh2 = ThisWorkbook.Sheets("期初余额表(临时)")


sh1.Range("A1:G65535").Copy

sh2.Range("A1:G65535").PasteSpecial xlPasteAll

Call Reflush

End Sub

Public Sub Reflush()

ListView1.ColumnHeaders.Clear:

ListView1.ListItems.Clear '清空列表

km = TextBox1.text

Dim N, i As Integer
Set sh = Sheets("期初余额表(临时)") '显示期初余额信息列表
lastColumn = sh.Cells(1, sh.Columns.count).End(xlToLeft).Column

lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row

ListView1.ColumnHeaders.Add 1, , sh.Cells(1, 1)
For i = 2 To lastColumn - 2
ListView1.ColumnHeaders.Add i, , sh.Cells(1, i + 2), ListView1.Width / 5 '可设宽度
Next
ListView1.View = lvwReport: ListView1.FullRowSelect = True: ListView1.Gridlines = True

For i = 2 To lastRow
If kmna = "" Or sh.Cells(i, 4) Like "*" & km & "*" Then
    Set itm = ListView1.ListItems.Add
    
    itm.text = sh.Cells(i, 1)
    For j = 4 To lastColumn
    itm.SubItems(j - 3) = sh.Cells(i, j)
    Next
End If
Next
If ListView1.ListItems.count > 0 Then ListView1.ListItems(1).Selected = False

End Sub

Sub serch()
    ListView1.ColumnHeaders.Clear
    ListView1.ListItems.Clear '清空列表
    ListView1.View = lvwReport
    ListView1.Gridlines = True '是否有表格线，True有表格线
    ListView1.FullRowSelect = True
    Set sh = Sheets("期初余额表(临时)")
    lastColumn = sh.Cells(1, sh.Columns.count).End(xlToLeft).Column
    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    
    ListView1.ColumnHeaders.Add 1, , sh.Cells(1, 1), ListView1.Width / 5
    
    For i = 2 To lastColumn - 2
    
    ListView1.ColumnHeaders.Add i, , sh.Cells(1, i + 2), ListView1.Width / 5
    
    Next


    For i = 2 To lastRow
    If sh.Range("d" & i) Like "*" & TextBox1.text & "*" Then
        Set itm = ListView1.ListItems.Add
        itm.text = sh.Cells(i, 1)
        
        For j = 4 To lastColumn
        
        itm.SubItems(j - 3) = sh.Cells(i, j)
        Next j
        
    End If
   Next i

End Sub

