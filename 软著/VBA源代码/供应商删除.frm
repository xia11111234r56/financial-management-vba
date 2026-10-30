VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 供应商删除 
   Caption         =   "删除供应商"
   ClientHeight    =   3900
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5780
   OleObjectBlob   =   "供应商删除.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "供应商删除"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Dim ws As Worksheet, r As Variant, no As Long
    Set ws = ThisWorkbook.Sheets("供应商档案")
    
    no = Val(Me.TextBox1.Value)
    If no = 0 Then
        MsgBox "请输入要删除的编号", vbExclamation, "提示": Exit Sub
    End If
    
    r = Application.Match(no, ws.Range("A:A"), 0)
    If IsError(r) Then
        MsgBox "供应商编号不存在！", vbExclamation, "提示": Exit Sub
    End If
    
    ws.Rows(r).Delete
    MsgBox "删除成功！", vbInformation, "提示"
    Me.Hide
    供应商管理.RefreshListView
End Sub

Private Sub CommandButton2_Click()
    Me.Hide
End Sub

