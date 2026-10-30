VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 会计科目 
   Caption         =   "会计科目"
   ClientHeight    =   9070.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   16960
   OleObjectBlob   =   "会计科目.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "会计科目"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private originalCode As String
Private originalName As String


' —— 窗体初始化 ——
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("会计科目表")
    
    originalCode = ""
    originalName = ""
    
    ' 初始化 ListView1
    With Me.ListView1
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        .ColumnHeaders.Add , , "科目编码", 60
        .ColumnHeaders.Add , , "一级科目编码", 80
        .ColumnHeaders.Add , , "二级科目编码", 80
        .ColumnHeaders.Add , , "科目名称", 120
        .ColumnHeaders.Add , , "辅助核算", 100
        .ColumnHeaders.Add , , "方向", 50
        .ColumnHeaders.Add , , "日记账", 50
        .ColumnHeaders.Add , , "银行账", 50
        .ColumnHeaders.Add , , "币别/计量单位", 90
    End With
    
    ' 下拉框初始化
    With Me.ComboBox1
        .Clear
        .AddItem "客户往来": .AddItem "供应商往来"
        .AddItem "项目核算": .AddItem "部门核算"
    End With
    With Me.ComboBox2
        .Clear
        .AddItem "借": .AddItem "贷"
    End With
    
    ' 清空输入区
    Call ClearInputs
    
    ' 加载数据
    LoadAllSubjects
End Sub

' —— 点击 ListView 行时填入输入区 ——
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.TextBox1.Value = Item.text
    Me.TextBox2.Value = Item.SubItems(1)
    Me.TextBox3.Value = Item.SubItems(2)
    Me.TextBox4.Value = Item.SubItems(3)
    Me.ComboBox1.Value = Item.SubItems(4)
    Me.ComboBox2.Value = Item.SubItems(5)
    Me.CheckBox1.Value = (Item.SubItems(6) = "Y")
    Me.CheckBox2.Value = (Item.SubItems(7) = "Y")
    Me.TextBox5.Value = Item.SubItems(8)
    
    originalCode = Item.text
    originalName = Item.SubItems(3)
End Sub

' —— 科目编码变化时实时回填 ——
Private Sub TextBox1_Change()
    Dim ws As Worksheet, rng As Range
    Set ws = ThisWorkbook.Worksheets("会计科目表")
    
    originalCode = "": originalName = ""
    
    If Trim(Me.TextBox1.Value) = "" Then
        Call ClearInputs
        Exit Sub
    End If
    
    Set rng = ws.Range("A:A").Find(What:=Me.TextBox1.Value, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rng Is Nothing Then
        ' 回填
        Me.TextBox2 = ws.Cells(rng.Row, "B").Value
        Me.TextBox3 = ws.Cells(rng.Row, "C").Value
        Me.TextBox4 = ws.Cells(rng.Row, "D").Value
        Me.ComboBox1 = ws.Cells(rng.Row, "E").Value
        Me.ComboBox2 = ws.Cells(rng.Row, "F").Value
        Me.CheckBox1 = (ws.Cells(rng.Row, "G").Value = "Y")
        Me.CheckBox2 = (ws.Cells(rng.Row, "H").Value = "Y")

        Me.TextBox5 = ws.Cells(rng.Row, "I").Value
        originalCode = Me.TextBox1.Value
        originalName = Me.TextBox4.Value
    Else
        ' 自动填一级/二级
        If Len(Me.TextBox1) > 4 Then
            Me.TextBox2 = Left(Me.TextBox1, 4)
            Me.TextBox3 = Me.TextBox1
        Else
            Me.TextBox2 = Me.TextBox1
            Me.TextBox3 = ""
        End If
        Me.TextBox4 = "": Me.ComboBox1 = ""
        Me.ComboBox2 = "": Me.CheckBox1 = False
        Me.CheckBox2 = False: Me.TextBox5 = ""
    End If
End Sub

' —— 查询按钮 ——
Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim codeCrit As String, nameCrit As String
    Dim matchesCriteria As Boolean
    Dim itm As ListItem
    
    Set ws = ThisWorkbook.Worksheets("会计科目表")
    
    ' 获取用户输入的查询条件（去掉前后空格）
    codeCrit = Trim(Me.TextBox1.Value)
    nameCrit = Trim(Me.TextBox4.Value)
    
    ' 清空列表视图
    With Me.ListView1
        .ListItems.Clear
    End With
    
    ' 获取最后一行行号
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "科目表中没有数据！", vbExclamation
        Exit Sub
    End If
    
    ' 遍历所有科目记录（从第2行开始）
    For i = 2 To lastRow
        matchesCriteria = True  ' 假设满足所有条件
        
        ' 如果输入了科目编码关键字，进行模糊匹配
        If codeCrit <> "" Then
            ' InStr 返回1表示找到匹配，0表示未找到
            If InStr(1, ws.Cells(i, "A").Value, codeCrit, vbTextCompare) = 0 Then
                matchesCriteria = False
            End If
        End If
        
        ' 如果输入了科目名称关键字，进行模糊匹配
        If nameCrit <> "" Then
            If InStr(1, ws.Cells(i, "D").Value, nameCrit, vbTextCompare) = 0 Then
                matchesCriteria = False
            End If
        End If
        
        ' 如果选择了辅助核算条件，检查是否匹配
        If Me.ComboBox1.Value <> "" Then
            If ws.Cells(i, "E").Value <> Me.ComboBox1.Value Then
                matchesCriteria = False
            End If
        End If
        
        ' 如果选择了方向条件，检查是否匹配
        If Me.ComboBox2.Value <> "" Then
            If ws.Cells(i, "F").Value <> Me.ComboBox2.Value Then
                matchesCriteria = False
            End If
        End If
       If Me.CheckBox1.Value = True Then
    If ws.Cells(i, "G").Value <> "Y" Then
        matchesCriteria = False
    End If
End If

If Me.CheckBox2.Value = True Then
    If ws.Cells(i, "H").Value <> "Y" Then
        matchesCriteria = False
    End If
End If

        
        ' 如果所有条件都满足，则添加到 ListView1 中
        If matchesCriteria Then
            Set itm = Me.ListView1.ListItems.Add(, , ws.Cells(i, "A").Value)
            itm.SubItems(1) = ws.Cells(i, "B").Value  ' 一级科目编码
            itm.SubItems(2) = ws.Cells(i, "C").Value  ' 二级科目编码
            itm.SubItems(3) = ws.Cells(i, "D").Value  ' 科目名称
            itm.SubItems(4) = ws.Cells(i, "E").Value  ' 辅助核算
            itm.SubItems(5) = ws.Cells(i, "F").Value  ' 方向
            itm.SubItems(6) = ws.Cells(i, "G").Value  ' 日记账
            itm.SubItems(7) = ws.Cells(i, "H").Value  ' 银行账
            itm.SubItems(8) = ws.Cells(i, "I").Value  ' 币别/计量单位
        End If
    Next i
    
    ' 如果没有找到任何匹配结果，则提醒用户
    If Me.ListView1.ListItems.count = 0 Then
        MsgBox "未找到匹配的科目。", vbInformation
    End If
End Sub


' —— 增加按钮 ——
Private Sub CommandButton2_Click()
    Dim ws As Worksheet, lastRow As Long, rng As Range
    Set ws = ThisWorkbook.Worksheets("会计科目表")
    
    If Trim(Me.TextBox1) = "" Or Trim(Me.TextBox4) = "" Or Trim(Me.ComboBox2) = "" Then
    MsgBox "科目编码、名称和方向必填！", vbExclamation
    Exit Sub
End If
    
    ' 重复检查
    If Not ws.Range("A:A").Find(Me.TextBox1, , xlValues, xlWhole) Is Nothing Then
        MsgBox "编码已存在！", vbExclamation: Exit Sub
    End If
    If Not ws.Range("D:D").Find(Me.TextBox4, , xlValues, xlWhole) Is Nothing Then
        MsgBox "名称已存在！", vbExclamation: Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row + 1
    ws.Cells(lastRow, "A").Value = Me.TextBox1
    ws.Cells(lastRow, "B").Value = Me.TextBox2
    ws.Cells(lastRow, "C").Value = Me.TextBox3
    ws.Cells(lastRow, "D").Value = Me.TextBox4
    ws.Cells(lastRow, "E").Value = Me.ComboBox1
    ws.Cells(lastRow, "F").Value = Me.ComboBox2
   ws.Cells(lastRow, "G").Value = IIf(Me.CheckBox1, "Y", "")
ws.Cells(lastRow, "H").Value = IIf(Me.CheckBox2, "Y", "")

    ws.Cells(lastRow, "I").Value = Me.TextBox5
    
    Call SortSubjects
    LoadAllSubjects
    MsgBox "新增成功！", vbInformation
    ClearInputs
End Sub

' —— 修改按钮 ——
Private Sub CommandButton3_Click()
    Dim ws As Worksheet, rng As Range, rowIdx As Long
    Set ws = ThisWorkbook.Worksheets("会计科目表")
    
    If Trim(Me.TextBox1) = "" Or Trim(Me.TextBox4) = "" Or Trim(Me.ComboBox2) = "" Then
    MsgBox "科目编码、名称和方向必填！", vbExclamation
    Exit Sub
End If
    If originalCode = "" Then
        MsgBox "请先在列表中选中或查询后再修改！", vbExclamation: Exit Sub
    End If
    
    Set rng = ws.Range("A:A").Find(originalCode, , xlValues, xlWhole)
    If rng Is Nothing Then MsgBox "原始记录丢失！", vbExclamation: Exit Sub
    rowIdx = rng.Row
    
    ' 检查编码/名称冲突及变更
    If Me.TextBox1 <> originalCode Then
        If Not ws.Range("A:A").Find(Me.TextBox1, , xlValues, xlWhole) Is Nothing Then
            MsgBox "新编码已存在！", vbExclamation: Exit Sub
        End If
    End If
    If Me.TextBox4 <> originalName Then
        If Not ws.Range("D:D").Find(Me.TextBox4, , xlValues, xlWhole) Is Nothing Then
            MsgBox "新名称已存在！", vbExclamation: Exit Sub
        End If
    End If
    
    ' 写回
    ws.Cells(rowIdx, "A").Value = Me.TextBox1
    ws.Cells(rowIdx, "B").Value = Me.TextBox2
    ws.Cells(rowIdx, "C").Value = Me.TextBox3
    ws.Cells(rowIdx, "D").Value = Me.TextBox4
    ws.Cells(rowIdx, "E").Value = Me.ComboBox1
    ws.Cells(rowIdx, "F").Value = Me.ComboBox2
    ws.Cells(rowIdx, "G").Value = IIf(Me.CheckBox1, "Y", "")
ws.Cells(rowIdx, "H").Value = IIf(Me.CheckBox2, "Y", "")

    ws.Cells(rowIdx, "I").Value = Me.TextBox5
    
    Call SortSubjects
    LoadAllSubjects
    MsgBox "修改成功！", vbInformation
    ClearInputs
End Sub

' —— 删除按钮 ——
Private Sub CommandButton4_Click()
    Dim ws As Worksheet, rng As Range
    Set ws = ThisWorkbook.Worksheets("会计科目表")
    
    If Trim(Me.TextBox1) = "" Then
        MsgBox "请输入或选中要删除的编码！", vbExclamation: Exit Sub
    End If
    
    Set rng = ws.Range("A:A").Find(Me.TextBox1, , xlValues, xlWhole)
    If rng Is Nothing Then
        MsgBox "未找到该科目！", vbExclamation: Exit Sub
    End If
    
    ws.Rows(rng.Row).Delete
    LoadAllSubjects
    MsgBox "删除成功！", vbInformation
    ClearInputs
End Sub

' —— 刷新按钮 ——
Private Sub CommandButton5_Click()
    LoadAllSubjects
    ClearInputs
End Sub

' —— 返回按钮 ——
Private Sub CommandButton6_Click()
    Unload Me
    On Error Resume Next
   系统平台.show
End Sub

' —— 辅助：加载所有科目到 ListView ——
Private Sub LoadAllSubjects()
    Dim ws As Worksheet, lastRow As Long
    Dim dataArr As Variant, sortedArr() As Variant
    Dim i As Long, j As Long, k As Long
    Dim tempRow(1 To 9) As Variant
    Dim itm As ListItem

    Set ws = ThisWorkbook.Worksheets("会计科目表")
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    dataArr = ws.Range("A2:I" & lastRow).Value

    ' 拷贝数据
    ReDim sortedArr(1 To UBound(dataArr, 1), 1 To 9)
    For i = 1 To UBound(dataArr, 1)
        For j = 1 To 9
            sortedArr(i, j) = dataArr(i, j)
        Next j
    Next i

    ' 自定义排序科目编码
    For i = 1 To UBound(sortedArr, 1) - 1
        For j = i + 1 To UBound(sortedArr, 1)
            If CompareSubjectCode(CStr(sortedArr(i, 1)), CStr(sortedArr(j, 1))) > 0 Then

                ' 交换整行
                For k = 1 To 9
                    tempRow(k) = sortedArr(i, k)
                    sortedArr(i, k) = sortedArr(j, k)
                    sortedArr(j, k) = tempRow(k)
                Next k
            End If
        Next j
    Next i

    ' 加载到 ListView
    With Me.ListView1
        .ListItems.Clear
        For i = 1 To UBound(sortedArr, 1)
            If sortedArr(i, 1) <> "" Then
                Set itm = .ListItems.Add(, , sortedArr(i, 1))
                itm.SubItems(1) = sortedArr(i, 2)
                itm.SubItems(2) = sortedArr(i, 3)
                itm.SubItems(3) = sortedArr(i, 4)
                itm.SubItems(4) = sortedArr(i, 5)
                itm.SubItems(5) = sortedArr(i, 6)
                itm.SubItems(6) = sortedArr(i, 7)
                itm.SubItems(7) = sortedArr(i, 8)
                itm.SubItems(8) = sortedArr(i, 9)
            End If
        Next i
    End With
End Sub


Private Function CompareSubjectCode(code1 As String, code2 As String) As Integer
    Dim parts1() As String, parts2() As String
    Dim i As Long, maxParts As Long
    Dim n1 As Long, N2 As Long

    ' 分段，每2位为一段（如10, 02, 01）
    parts1 = SplitEveryN(code1, 2)
    parts2 = SplitEveryN(code2, 2)

    maxParts = Application.WorksheetFunction.Max(UBound(parts1), UBound(parts2))
    For i = 0 To maxParts
        If i > UBound(parts1) Then
            CompareSubjectCode = -1
            Exit Function
        ElseIf i > UBound(parts2) Then
            CompareSubjectCode = 1
            Exit Function
        End If

        n1 = Val(parts1(i))
        N2 = Val(parts2(i))

        If n1 <> N2 Then
            CompareSubjectCode = Sgn(n1 - N2)
            Exit Function
        End If
    Next i

    ' 完全相等
    CompareSubjectCode = 0
End Function

Private Function SplitEveryN(s As String, N As Integer) As String()
    Dim result() As String
    Dim i As Long, count As Long

    ReDim result(0 To (Len(s) - 1) \ N)
    count = 0
    For i = 1 To Len(s) Step N
        result(count) = Mid(s, i, N)
        count = count + 1
    Next i
    SplitEveryN = result
End Function


' —— 辅助：排序 ——
Private Sub SortSubjects()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataArr As Variant, sortedArr() As Variant
    Dim i As Long, j As Long, k As Long
    Dim tempRow(1 To 9) As Variant

    Set ws = ThisWorkbook.Worksheets("会计科目表")
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    dataArr = ws.Range("A2:I" & lastRow).Value

    ' 复制数据到 sortedArr
    ReDim sortedArr(1 To UBound(dataArr, 1), 1 To 9)
    For i = 1 To UBound(dataArr, 1)
        For j = 1 To 9
            sortedArr(i, j) = dataArr(i, j)
        Next j
    Next i

    ' 按编码进行冒泡排序
    For i = 1 To UBound(sortedArr, 1) - 1
        For j = i + 1 To UBound(sortedArr, 1)
          If CompareSubjectCode(CStr(sortedArr(i, 1)), CStr(sortedArr(j, 1))) > 0 Then

                ' 交换整行
                For k = 1 To 9
                    tempRow(k) = sortedArr(i, k)
                    sortedArr(i, k) = sortedArr(j, k)
                    sortedArr(j, k) = tempRow(k)
                Next k
            End If
        Next j
    Next i

    ' 清空原始数据并写入排序后的内容
    ws.Range("A2:I" & lastRow).ClearContents
    ws.Range("A2").Resize(UBound(sortedArr, 1), 9).Value = sortedArr
End Sub


' —— 辅助：清空输入区 ——
Private Sub ClearInputs()
    Me.TextBox1 = ""
    Me.TextBox2 = ""
    Me.TextBox3 = ""
    Me.TextBox4 = ""
    Me.ComboBox1 = ""
    Me.ComboBox2 = ""
    Me.CheckBox1 = False
    Me.CheckBox2 = False
    Me.TextBox5 = ""
    originalCode = ""
    originalName = ""
End Sub

