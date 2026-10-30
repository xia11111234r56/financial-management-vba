VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 人员增加 
   Caption         =   "增加人员"
   ClientHeight    =   7540
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5750
   OleObjectBlob   =   "人员增加.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "人员增加"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ComboBox3_Change()
    ' 根据部门总称更新行政部门
    Dim deptSheet As Worksheet
    Set deptSheet = Sheets("部门档案")
    
    Dim selectedDept As String
    Dim parentCode As String
    Dim i As Long, N As Long
    Dim deptCode As String
    Dim deptName As String
    Dim matchFound As Boolean
    
    ComboBox2.Clear
    selectedDept = ComboBox3.Value
    matchFound = False
    
    N = deptSheet.Cells(deptSheet.Rows.count, "A").End(xlUp).Row
    
    ' 找出选中部门总称的编号（为一位数）
    For i = 2 To N
        If Trim(deptSheet.Cells(i, "B").Value) = selectedDept And Len(Trim(deptSheet.Cells(i, "A").Value)) = 1 Then
            parentCode = Trim(deptSheet.Cells(i, "A").Value)
            Exit For
        End If
    Next i
    
    ' 遍历查找子部门（编号为3位且以父编号开头）
    For i = 2 To N
        deptCode = Trim(deptSheet.Cells(i, "A").Value)
        deptName = Trim(deptSheet.Cells(i, "B").Value)
        If Len(deptCode) = 3 And Left(deptCode, 1) = parentCode Then
            ComboBox2.AddItem deptName
            matchFound = True
        End If
    Next i
    
    ' 如果没找到子部门，则直接加入部门总称作为行政部门
    If Not matchFound Then
        ComboBox2.AddItem selectedDept
    End If
End Sub

Private Sub CommandButton1_Click()
    Dim sh As Worksheet
    Set sh = Sheets("人员档案")
    
    Dim no As String
    Dim name As String
    Dim sex As String
    Dim div As String
    Dim status As String
    Dim N As Long, i As Long
    
    no = TextBox1.Value
    name = TextBox2.Value
    sex = ComboBox1.Value
    div = ComboBox3.Value
    status = ComboBox4.Value
    
    If Len(no) = 0 Or Len(name) = 0 Or Len(sex) = 0 Or Len(ComboBox2.Value) = 0 _
    Or Len(div) = 0 Or Len(status) = 0 Or Len(ComboBox5.Value) = 0 Then
        MsgBox "信息缺失，请重新填写！", , "提示"
        Exit Sub
    End If
    
    N = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    For i = 2 To N
        If sh.Cells(i, "A").Value = no Then
            MsgBox "该编号已存在，请重新输入！", , "提示"
            Exit Sub
        End If
    Next i

    With sh
    .Cells(N + 1, "A").Value = no
    .Cells(N + 1, "B").Value = name
    .Cells(N + 1, "C").Value = sex
    .Cells(N + 1, "D").Value = div
    .Cells(N + 1, "E").Value = ComboBox2.Value
    .Cells(N + 1, "F").Value = status
    .Cells(N + 1, "G").Value = ComboBox5.Value
End With

    MsgBox "增加成功！", , "提示"
    Unload Me
    人员管理.Hide
    人员管理.show
End Sub


Private Sub CommandButton2_Click()
 Unload Me
End Sub

Private Sub UserForm_Activate()
    Dim deptSheet As Worksheet
    Set deptSheet = Sheets("部门档案")
    
    Dim N As Long, i As Long
    Dim deptName As String
    Dim deptCode As String
    
    ComboBox1.Clear: ComboBox1.AddItem "男": ComboBox1.AddItem "女"
    ComboBox4.Clear: ComboBox4.AddItem "在职": ComboBox4.AddItem "离职"
    ComboBox5.Clear: ComboBox5.AddItem "是": ComboBox5.AddItem "否"
    ComboBox3.Clear
    N = deptSheet.Cells(deptSheet.Rows.count, "A").End(xlUp).Row
    For i = 2 To N
        deptCode = Trim(deptSheet.Cells(i, "A").Value)
        deptName = Trim(deptSheet.Cells(i, "B").Value)
        If Len(deptCode) = 1 Then
            ComboBox3.AddItem deptName
        End If
    Next i
End Sub


