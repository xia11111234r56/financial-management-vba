VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 人员修改 
   Caption         =   "修改人员"
   ClientHeight    =   7820
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5840
   OleObjectBlob   =   "人员修改.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "人员修改"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim wsDept As Worksheet
    Dim lastRow As Long, i As Long
    Dim deptCode As String, deptName As String

    Set wsDept = ThisWorkbook.Sheets("部门档案")

    ' 初始化性别
    ComboBox1.Clear
    ComboBox1.AddItem "男"
    ComboBox1.AddItem "女"

    ' 初始化在职状态
    ComboBox4.Clear
    ComboBox4.AddItem "在职"
    ComboBox4.AddItem "离职"

    ' 初始化是否业务员
    ComboBox5.Clear
    ComboBox5.AddItem "是"
    ComboBox5.AddItem "否"

    ' 初始化部门总称
    ComboBox3.Clear
    lastRow = wsDept.Cells(wsDept.Rows.count, "A").End(xlUp).Row

    For i = 2 To lastRow
        deptCode = Trim(wsDept.Cells(i, 1).Value)
        deptName = Trim(wsDept.Cells(i, 2).Value)
        If Len(deptCode) = 1 Then
            ComboBox3.AddItem deptName
        End If
    Next i
End Sub

Private Sub ComboBox3_Change()
    ' 部门总称变化时联动更新行政部门
    Dim wsDept As Worksheet
    Dim lastRow As Long, i As Long
    Dim mainDeptName As String, mainDeptCode As String
    Dim deptCode As String, deptName As String
    Dim found As Boolean

    Set wsDept = ThisWorkbook.Sheets("部门档案")
    mainDeptName = ComboBox3.Value
    ComboBox2.Clear
    found = False

    lastRow = wsDept.Cells(wsDept.Rows.count, "A").End(xlUp).Row

    ' 查找对应的主部门编号
    For i = 2 To lastRow
        If Trim(wsDept.Cells(i, 2).Value) = mainDeptName And Len(Trim(wsDept.Cells(i, 1).Value)) = 1 Then
            mainDeptCode = Trim(wsDept.Cells(i, 1).Value)
            Exit For
        End If
    Next i

    ' 查找所有与主部门对应的行政部门
    For i = 2 To lastRow
        deptCode = Trim(wsDept.Cells(i, 1).Value)
        deptName = Trim(wsDept.Cells(i, 2).Value)

        If Len(deptCode) = 3 And Left(deptCode, 1) = mainDeptCode Then
            ComboBox2.AddItem deptName
            found = True
        End If
    Next i

    ' 若未找到行政部门，则默认本部门为行政部门
    If Not found Then
        ComboBox2.AddItem mainDeptName
    End If
End Sub

' 初始化并加载指定人员信息
Public Sub InitializeData(ByVal personID As String)
    Dim ws As Worksheet
    Dim rowIndex As Long
    
    Set ws = ThisWorkbook.Sheets("人员档案")
    
    On Error GoTo NotFound
    rowIndex = WorksheetFunction.Match(Val(personID), ws.Range("A:A"), 0)
    
    ' 加载数据到控件
    TextBox1.Value = ws.Cells(rowIndex, 1).Value
    TextBox2.Value = ws.Cells(rowIndex, 2).Value
    ComboBox1.Value = ws.Cells(rowIndex, 3).Value
    ComboBox3.Value = ws.Cells(rowIndex, 4).Value
    
    ' 触发 ComboBox3_Change 事件，联动更新行政部门
    Call ComboBox3_Change
    
    ComboBox2.Value = ws.Cells(rowIndex, 5).Value
    ComboBox4.Value = ws.Cells(rowIndex, 6).Value
    ComboBox5.Value = ws.Cells(rowIndex, 7).Value

    ' 人员编号不可修改
    TextBox1.Enabled = False
    TextBox1.BackColor = &H8000000B  ' 灰色禁用效果
    
    Exit Sub

NotFound:
    MsgBox "人员编号未找到", vbExclamation, "提示"
End Sub

Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim no As String, na As String, sex As String
    Dim bmzc As String, xzbm As String, gyzt As String, ywy As String

    Set ws = ThisWorkbook.Sheets("人员档案")

    no = Trim(TextBox1.text)
    na = Trim(TextBox2.text)
    sex = ComboBox1.text
    bmzc = ComboBox3.text
    xzbm = ComboBox2.text
    gyzt = ComboBox4.text
    ywy = ComboBox5.text

    If no = "" Or na = "" Or sex = "" Or bmzc = "" Or xzbm = "" Or gyzt = "" Or ywy = "" Then
        MsgBox "存在空输入", , "提示"
        Exit Sub
    End If

    On Error GoTo NotFound
    rowIndex = WorksheetFunction.Match(Val(no), ws.Range("A:A"), 0)

    ws.Cells(rowIndex, 2).Value = na
    ws.Cells(rowIndex, 3).Value = sex
    ws.Cells(rowIndex, 4).Value = bmzc
    ws.Cells(rowIndex, 5).Value = xzbm
    ws.Cells(rowIndex, 6).Value = gyzt
    ws.Cells(rowIndex, 7).Value = ywy

    MsgBox "修改信息成功", , "提示"
    Unload Me
    
    ' 刷新人员管理主窗体列表
    人员管理.InitTreeView
    人员管理.FillListView

    Exit Sub

NotFound:
    MsgBox "人员编号未找到", , "提示"
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

