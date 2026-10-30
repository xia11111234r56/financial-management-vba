VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 用户删除 
   Caption         =   "UserForm2"
   ClientHeight    =   3690
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5820
   OleObjectBlob   =   "用户删除.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "用户删除"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If Len(TextBox1.Value) = 0 Then
        MsgBox "编号不能为空！", , "提示"
        Exit Sub
    End If
    Dim str As String
    Dim N As Long
    Dim i As Long
    Dim sig As Integer
    
    With Sheets("用户档案")
        N = .Cells(.Rows.count, 1).End(xlUp).Row
        sig = 0
        str = TextBox1.Value
        
        For i = 3 To N
            If .Cells(i, 1).Value = str Then
                sig = 1
                Exit For
            End If
        Next
        
        If sig = 0 Then
            MsgBox "您输入的用户不存在！", , "提示"
            Exit Sub
        End If
        
        Debug.Print i
        Debug.Print N
        
        If i < N Then
            .Range(.Cells(i + 1, 1), .Cells(N, 4)).Cut Destination:=.Range(.Cells(i, 1), .Cells(N - 1, 4))
        Else
            .Rows(N).ClearContents
        End If
    End With
    
    Unload Me
    用户管理.Hide
    用户管理.show
End Sub


Private Sub CommandButton2_Click()
 Unload Me
End Sub

