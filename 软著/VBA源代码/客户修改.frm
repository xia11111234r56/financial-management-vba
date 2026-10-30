VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 客户修改 
   Caption         =   "修改客户"
   ClientHeight    =   7340
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5580
   OleObjectBlob   =   "客户修改.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "客户修改"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 模块顶部定义一个变量保存当前行号（根据编号定位）
Private currentRow As Long

Private Sub UserForm_Initialize()
    LoadDepartments
    ShowEditControls False ' 初始状态隐藏编辑区
End Sub

' 初始化数据子程序，外部调用，用于加载客户信息
Public Sub InitializeData(ByVal 客户编号 As String)
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long

    Set ws = Sheets("客户档案")
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = 客户编号 Then
            currentRow = i
            ' 填充控件
            TextBox1.Value = ws.Cells(i, 1).Value    ' 编号，通常不可修改
            TextBox2.Value = ws.Cells(i, 2).Value    ' 名称
            TextBox5.Value = ws.Cells(i, 3).Value    ' 简称
            ComboBox1.Value = ws.Cells(i, 4).Value   ' 部门
            TextBox3.Value = ws.Cells(i, 6).Value    ' 联系方式
            
            ' 加载业务员（先设置部门再加载业务员列表）
            Call ComboBox1_Change
            ComboBox2.Value = ws.Cells(i, 5).Value   ' 业务员

            ShowEditControls True
            Exit Sub
        End If
    Next i

    MsgBox "未找到客户编号：" & 客户编号, vbExclamation
    ShowEditControls False
End Sub

Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Set ws = Sheets("客户档案")

    ' 校验：客户名称不能为空
    If Len(Trim(TextBox2.Value)) = 0 Then
        MsgBox "客户名称不能为空！", vbExclamation
        Exit Sub
    End If

    ' 写回工作表
    ws.Cells(currentRow, 2).Value = TextBox2.Value
    ws.Cells(currentRow, 3).Value = TextBox5.Value
    ws.Cells(currentRow, 4).Value = ComboBox1.Value
    ws.Cells(currentRow, 5).Value = ComboBox2.Value ' 即使为空也允许
    ws.Cells(currentRow, 6).Value = TextBox3.Value

    MsgBox "修改成功！", vbInformation

    Me.Hide
    客户管理.RefreshListView

End Sub

Private Sub CommandButton2_Click()
    Me.Hide

End Sub

Private Sub LoadDepartments()
    Dim wsDept As Worksheet
    Dim lastRow As Long, i As Long
    Set wsDept = Sheets("部门档案")

    lastRow = wsDept.Cells(wsDept.Rows.count, "A").End(xlUp).Row
    ComboBox1.Clear
    For i = 2 To lastRow
        ComboBox1.AddItem wsDept.Cells(i, 2).Value
    Next i

    ComboBox2.Clear ' 清空业务员列表
End Sub

Private Sub ComboBox1_Change()
    Dim wsPerson As Worksheet
    Dim lastRow As Long, i As Long
    Set wsPerson = Sheets("人员档案")

    ComboBox2.Clear
    lastRow = wsPerson.Cells(wsPerson.Rows.count, "A").End(xlUp).Row

    For i = 2 To lastRow
        If wsPerson.Cells(i, "E").Value = ComboBox1.Value And wsPerson.Cells(i, "G").Value = "是" Then
            ComboBox2.AddItem wsPerson.Cells(i, "B").Value
        End If
    Next i
End Sub

' 控件显示控制
Private Sub ShowEditControls(show As Boolean)
    TextBox1.Enabled = False   ' 编号不可编辑
    TextBox1.BackColor = IIf(show, vbBtnFace, vbWhite)

    TextBox2.Visible = show
    TextBox3.Visible = show
    TextBox5.Visible = show
    ComboBox1.Visible = show
    ComboBox2.Visible = show

    CommandButton1.Visible = show
    CommandButton2.Visible = show
End Sub

