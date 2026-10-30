VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 供应商修改 
   Caption         =   "修改供应商"
   ClientHeight    =   5770
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5280
   OleObjectBlob   =   "供应商修改.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "供应商修改"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 窗体初始化：加载分管业务员列表
Private Sub UserForm_Initialize()
    Dim wsP As Worksheet
    Dim lastRow As Long, i As Long
    Set wsP = ThisWorkbook.Sheets("人员档案")
    lastRow = wsP.Cells(wsP.Rows.count, "A").End(xlUp).Row
    
    Me.ComboBox1.Clear
    For i = 2 To lastRow
        If Trim(wsP.Cells(i, "G").Value) = "是" Then
            Me.ComboBox1.AddItem wsP.Cells(i, "B").Value
        End If
    Next i
End Sub

' 根据编号加载并显示当前供应商信息
Public Sub InitializeData(ByVal no As Long)
    Dim ws As Worksheet
    Dim r As Variant
    Set ws = ThisWorkbook.Sheets("供应商档案")
    
    r = Application.Match(no, ws.Range("A:A"), 0)
    If IsError(r) Then
        MsgBox "供应商编号不存在！", vbExclamation, "提示"
        Exit Sub
    End If
    
    Me.TextBox1.Value = ws.Cells(r, 1).Value
    Me.TextBox2.Value = ws.Cells(r, 2).Value
    Me.TextBox3.Value = ws.Cells(r, 3).Value
    
    ' 确保业务员列表已加载
    Call UserForm_Initialize
    
    Me.ComboBox1.Value = ws.Cells(r, 4).Value
    
    ' 编号不可修改
    Me.TextBox1.Enabled = False
    Me.TextBox1.BackColor = &H8000000B
End Sub

' 确认修改
Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim r As Variant
    Dim no As Long
    
    Set ws = ThisWorkbook.Sheets("供应商档案")
    no = Val(Me.TextBox1.Value)
    
    r = Application.Match(no, ws.Range("A:A"), 0)
    If IsError(r) Then
        MsgBox "供应商编号不存在！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ws.Cells(r, 2).Value = Trim(Me.TextBox2.Value)
    ws.Cells(r, 3).Value = Trim(Me.TextBox3.Value)
    ws.Cells(r, 4).Value = Trim(Me.ComboBox1.Value)
    
    MsgBox "修改成功！", vbInformation, "提示"
    Me.Hide
    供应商管理.RefreshListView
End Sub

' 取消
Private Sub CommandButton2_Click()
    Me.Hide
End Sub

