VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 人员删除 
   Caption         =   "删除人员"
   ClientHeight    =   3520
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6290
   OleObjectBlob   =   "人员删除.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "人员删除"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Dim sh As Worksheet
    Dim s As Variant
    Dim sel As String
    
    Set sh = Sheets("人员档案")
    sel = Trim(TextBox1.Value)
    
    If sel = "" Then
        MsgBox "请输入人员编号", vbExclamation, "提示"
        Exit Sub
    End If
    
    On Error Resume Next
    s = Application.Match(CLng(sel), sh.Range("A:A"), 0)
    On Error GoTo 0
    
    If IsError(s) Then
        MsgBox "未找到匹配的人员", vbExclamation, "提示"
        Exit Sub
    Else
        sh.Rows(s).Delete
        ' 刷新人员管理窗体列表和树
        人员管理.InitTreeView
        人员管理.FillListView
        MsgBox "删除成功！", vbInformation, "提示"
    End If
    
    TextBox1.Value = ""
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

