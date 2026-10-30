VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 部门删除 
   Caption         =   "删除部门"
   ClientHeight    =   3860
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5590
   OleObjectBlob   =   "部门删除.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "部门删除"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim sh As Worksheet
    Dim s As Variant
    Dim sel As String
    Dim selNum As Double
    
    Set sh = Sheets("部门档案")
    
    sel = Trim(TextBox1.Value)
    If sel = "" Then
        MsgBox "请输入部门编号", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 尝试转换成数字，如果失败就当成文本匹配
    On Error Resume Next
    selNum = CDbl(sel)
    On Error GoTo 0
    
    If IsNumeric(sel) Then
        ' 如果输入是数字，则匹配数字型部门编号
        s = Application.Match(selNum, sh.Range("A:A"), 0)
    Else
        ' 否则按文本匹配
        s = Application.Match(sel, sh.Range("A:A"), 0)
    End If
    
    If IsError(s) Then
        MsgBox "未找到匹配的部门", vbExclamation, "提示"
        Exit Sub
    Else
        sh.Rows(s).Delete
        Call 部门管理.Flesh
        Call 部门管理.Flesh2
        MsgBox "删除成功！", vbInformation, "提示"
    End If
    
    TextBox1.Value = ""
    Unload Me
End Sub



Private Sub CommandButton2_Click()
 Unload Me
End Sub

