VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 客户增加 
   Caption         =   "增加客户"
   ClientHeight    =   6930
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5910
   OleObjectBlob   =   "客户增加.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "客户增加"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
    Dim wsPerson As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set wsPerson = Sheets("人员档案")
    lastRow = wsPerson.Cells(wsPerson.Rows.count, "A").End(xlUp).Row
    
    ComboBox2.Clear
    
    If ComboBox1.Value = "" Then Exit Sub
    
    For i = 2 To lastRow
        If wsPerson.Cells(i, "E").Value = ComboBox1.Value And _
           LCase(wsPerson.Cells(i, "G").Value) = "是" Then
            ComboBox2.AddItem wsPerson.Cells(i, "B").Value
        End If
    Next i
End Sub


Private Sub CommandButton1_Click()
    Dim N As Long
    Dim i As Long
    Dim no As String
    Dim name As String
    Dim shortName As String
    Dim dept As String
    Dim salesman As String
    Dim tel As String
    
    N = Sheets("客户档案").Cells(Rows.count, "A").End(xlUp).Row

    ' 必填项校验：编号、名称、联系方式不能为空，分管业务员可空
    If Len(TextBox1.Value) = 0 Or Len(TextBox2.Value) = 0 Or Len(TextBox3.Value) = 0 Then
        MsgBox "必填项存在缺失，请重新填写！", , "提示"
        Exit Sub
    End If
    
    no = TextBox1.Value
    name = TextBox2.Value
    shortName = TextBox5.Value
    dept = ComboBox1.Value
    salesman = ComboBox2.Value ' 可为空
    tel = TextBox3.Value
    
    ' 检查编号和名称是否重复
    For i = 2 To N
        If Sheets("客户档案").Cells(i, "A").Value = no Then
            MsgBox "该编号已存在，请重新输入！", , "提示"
            Exit Sub
        End If
        If Sheets("客户档案").Cells(i, "B").Value = name Then
            MsgBox "该公司名称已存在，请重新输入！", , "提示"
            Exit Sub
        End If
    Next i
    
    ' 添加数据
    Sheets("客户档案").Cells(N + 1, "A").Value = no
    Sheets("客户档案").Cells(N + 1, "B").Value = name
    Sheets("客户档案").Cells(N + 1, "C").Value = shortName
    Sheets("客户档案").Cells(N + 1, "D").Value = dept
    Sheets("客户档案").Cells(N + 1, "E").Value = salesman
    Sheets("客户档案").Cells(N + 1, "F").Value = tel

    MsgBox "增加成功！", , "提示"
    
    Unload Me
    
    ' ?添加刷新客户管理列表视图
    客户管理.RefreshListView

End Sub



Private Sub CommandButton2_Click()
    Unload Me

End Sub

Private Sub UserForm_Initialize()
    Call LoadDepartments
End Sub

Private Sub LoadDepartments()
    Dim wsDept As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dictFirstLevel As Object
    Dim dictHasSecondLevel As Object
    Dim firstLevelNo As String
    Dim deptNo As String
    Dim key As Variant
    
    Set wsDept = Sheets("部门档案")
    lastRow = wsDept.Cells(wsDept.Rows.count, "A").End(xlUp).Row
    
    Set dictFirstLevel = CreateObject("Scripting.Dictionary")
    Set dictHasSecondLevel = CreateObject("Scripting.Dictionary")
    
    ComboBox1.Clear
    
    ' 先收集所有一级部门编号（文本）和名称
    For i = 2 To lastRow
        deptNo = CStr(wsDept.Cells(i, "A").Value)
        If Len(deptNo) = 1 Then
            dictFirstLevel(deptNo) = wsDept.Cells(i, "B").Value
        End If
    Next i
    
    ' 遍历找出所有二级部门，并标记对应一级部门有二级部门
    For i = 2 To lastRow
        deptNo = CStr(wsDept.Cells(i, "A").Value)
        If Len(deptNo) > 1 Then
            firstLevelNo = Left(deptNo, 1)
            If dictFirstLevel.exists(firstLevelNo) Then
                dictHasSecondLevel(firstLevelNo) = True
            End If
            ComboBox1.AddItem wsDept.Cells(i, "B").Value ' 添加二级部门名称
        End If
    Next i
    
    ' 添加没有二级部门的一级部门名称
    For Each key In dictFirstLevel.keys
        If Not dictHasSecondLevel.exists(key) Then
            ComboBox1.AddItem dictFirstLevel(key)
        End If
    Next key
End Sub

