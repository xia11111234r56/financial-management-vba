VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 客户删除 
   Caption         =   "删除客户"
   ClientHeight    =   3580
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5730
   OleObjectBlob   =   "客户删除.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "客户删除"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If Len(TextBox1.Value) = 0 Then
        MsgBox "编号不能为空", , "提示"
        Exit Sub
    End If

    Dim str As String
    Dim N As Integer
    Dim i As Integer
    Dim sig As Integer
    Dim ws As Worksheet
    Set ws = Sheets("客户档案")

    sig = 0
    str = TextBox1.Value
    N = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    For i = 2 To N
        If ws.Cells(i, 1).Value = str Then
            sig = 1
            Exit For
        End If
    Next

    If sig = 0 Then
        MsgBox "您输入的客户不存在！", , "提示"
        Exit Sub
    End If

    If i < N Then
        ws.Range(ws.Cells(i + 1, 1), ws.Cells(N, 6)).Cut ws.Cells(i, 1)
    End If
    ws.Rows(N).ClearContents

    MsgBox "删除成功！", , "提示"
    Unload Me

    ' ?添加刷新客户管理界面
    客户管理.RefreshListView

End Sub


Private Sub CommandButton2_Click()
    Unload Me

End Sub

