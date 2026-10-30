VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 登录页面 
   Caption         =   "登录"
   ClientHeight    =   3950
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6180
   OleObjectBlob   =   "登录页面.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "登录页面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox1_Click()

 If CheckBox1.Value = True Then
        TextBox2.PasswordChar = ""
    Else
        TextBox2.PasswordChar = "*"
    End If

End Sub

Private Sub CommandButton1_Click()
    Dim sh As Worksheet
    Dim user As Range
    Dim na As String
    Dim pw As String

    Set sh = Sheets("用户档案")
    na = ComboBox1.text
    pw = TextBox2.text

    If na = "" Or pw = "" Then
        MsgBox "未输入用户名或密码", vbExclamation
        Exit Sub
    End If

    Set user = sh.Range("B:B").Find(na, LookIn:=xlValues, LookAt:=xlWhole)

    If Not user Is Nothing Then
        If pw <> user.Offset(0, 1).Value Then
            MsgBox "密码错误！", vbCritical
            Exit Sub
        Else
            
        MsgBox "登录成功！", vbInformation
        currentUser = user.Row
            Me.Hide
            
            系统平台.show
        End If
    Else
        MsgBox "账户不存在！", vbExclamation
    End If
End Sub


Private Sub CommandButton2_Click()
Unload Me
    ThisWorkbook.save
    ThisWorkbook.Close
End Sub

Private Sub UserForm_Activate()
ComboBox1.text = ""
TextBox2.text = ""
'TextBox2.PasswordChar = "*"
End Sub

Private Sub UserForm_Initialize()
 Dim N As Long, i As Long
    Dim sht As Worksheet

    Set sht = Sheets("用户档案")
    ComboBox1.Clear
    N = sht.Cells(sht.Rows.count, "A").End(xlUp).Row

    For i = 2 To N
        If Trim(sht.Cells(i, "B").Value) <> "" Then
            ComboBox1.AddItem sht.Cells(i, "B").Value
        End If
    Next i

    CheckBox1.Value = False
    TextBox2.PasswordChar = "*"  ' 只在这里设置，确保不冲突
    TextBox2.text = ""           ' 初始化内容
End Sub

