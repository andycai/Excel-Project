VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TeamTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit

Private m_proj, m_template_name, m_setting_name As String
Private isclear As Boolean
Private dict_clear As Object

Private m_handler, m_acceptor, m_handle_stat, m_accept_stat, m_keyword As String
Private m_dirname, m_filename, m_type, m_update_handle, m_update_accept, m_sheetname, m_clear_flag As String

'初始化
Public Sub Init(proj)
    m_template_name = "模板"
    m_setting_name = "说明"
    If proj = "" Then
        m_proj = ThisWorkbook.Sheets(m_setting_name).Range("A2").Value
    Else
        m_proj = proj
    End If
    
    m_handler = GetConfig(2)
    m_acceptor = GetConfig(3)
    m_handle_stat = GetConfig(4)
    m_accept_stat = GetConfig(5)
    m_keyword = GetConfig(6)
    m_dirname = GetConfig(7)
    m_filename = GetConfig(8)
    m_type = GetConfig(9)
    m_update_handle = GetConfig(10)
    m_update_accept = GetConfig(11)
    m_sheetname = GetConfig(12)
    m_clear_flag = GetConfig(13)
End Sub

'拉取任务
Public Sub PullTasks()
    Dim book As Workbook
    Dim f, file(), activefile, ret, sheetname As String
    
    Let activefile = ActiveWorkbook.name
    Let sheetname = ""
    
    isclear = False
    Set dict_clear = CreateObject("Scripting.Dictionary")
    
    ReDim file(1)
    Let file(1) = ThisWorkbook.Path & "\..\" & m_dirname & "\"
  
    If m_filename = "所有" Then
        Let f = Dir(file(1) & "*.xlsm")     '通配符*.*表示所有文件，*.xlsx Excel文件
    Else
        Let f = Dir(file(1) & m_filename & ".xlsm")
    End If
    Do Until f = ""
        If f <> activefile Then
            Set book = Workbooks.Open(file(1) & f)
            If m_type = "人名" Then
                sheetname = Left(f, InStr(1, f, ".") - 1)
            ElseIf m_type = "汇总" Then
                sheetname = m_sheetname
            End If
            Let ret = CopyTask(book, sheetname)
'            Let ret = CopySheets(book)
            book.Close
        End If
        Let f = Dir
    Loop
End Sub

'复制数据
Private Function CopyTask(book, sheetname)
    Dim i, j, k, count, rowcount As Integer
    Dim src_sheet, dist_sheet As Worksheet
    Dim found As Boolean
    
    Let count = 2
    For i = 1 To book.Sheets.count
        Set src_sheet = book.Sheets(i)
        If InStr(1, src_sheet.name, m_keyword) Then
            Let rowcount = src_sheet.Range("A65536").End(xlUp).row
            For j = 1 To rowcount
                If FilterName(src_sheet.Cells(j, 1).Value) And _
                    MatchName(m_handler, src_sheet.Cells(j, 6).Value) And _
                    MatchName(m_acceptor, src_sheet.Cells(j, 7).Value) And _
                    MatchStat(m_handle_stat, src_sheet.Cells(j, 4).Value) And _
                    MatchStat(m_accept_stat, src_sheet.Cells(j, 5).Value) Then
                    '系统，多个sheet
                    If sheetname = "" Then
                        Set dist_sheet = CreateSheet(src_sheet.name) 'sheet name 与 source 一致
                        
                        '清空sheet
                        If NeedClearSheet(m_clear_flag) And Not dict_clear.exists(src_sheet.name) Then
                            ClearSheet (src_sheet.name)
                            dict_clear(src_sheet.name) = 1
                        End If
                    Else '单个sheet
                        Set dist_sheet = CreateSheet(sheetname) '自定义sheet name
                        '清空sheet
                        If NeedClearSheet(m_clear_flag) And isclear = False Then
                            ClearSheet (sheetname)
                            Let isclear = True
                        End If
                    End If
                         
                    Let count = dist_sheet.UsedRange.Rows.count
                    Let found = False
                    
                    If Not NeedClearSheet(m_clear_flag) Then
                        For k = 1 To count
                            If Trim(dist_sheet.Cells(k, 3).Value) = Trim(src_sheet.Cells(j, 3).Value) Then '存在相同的任务
                                If NeedUpdateField(m_update_handle) Then
                                    src_sheet.Range("D" & j).Copy dist_sheet.Range("D" & k) '更新完成情况
                                End If
                                If NeedUpdateField(m_update_accept) Then
                                    src_sheet.Range("E" & j).Copy dist_sheet.Range("E" & k) '更新验收情况
                                End If
                                found = True
                                Exit For
                            End If
                        Next
                    End If
                    
                    If Not found Then
                        src_sheet.Rows(j & ":" & j).Copy dist_sheet.Range("A" & (count + 1)) '复制到指定位置
                    End If
                End If
            Next
        End If
    Next
End Function

'修正任务
Private Function ReviseTasks()
    Dim wbook As Workbook
    Dim src_sheet As Worksheet
    Dim i, j, k, rowcount, count As Integer
    
    Dim tasks As Object
    Set tasks = CreateObject("Scripting.Dictionary")
    
    Set wbook = Workbooks.Open(ThisWorkbook.Path & "\..\修正任务.xlsx")
    For i = 1 To wbook.Sheets.count
        Set src_sheet = wbook.Sheets(i)
        If InStr(1, src_sheet.name, "修正") Then
            Let rowcount = src_sheet.Range("A65536").End(xlUp).row
            For j = 1 To rowcount
                If src_sheet.Cells(j, 3).Value <> "内容" Then
                    tasks.Item(src_sheet.Cells(j, 3).Value) = Trim(src_sheet.Cells(j, 6).Value)
                End If
            Next
        End If
    Next
    wbook.Close
    
    Dim dist_sheet As Worksheet
    For j = 1 To ThisWorkbook.Sheets.count
        Set dist_sheet = ThisWorkbook.Sheets(j)
        Let count = dist_sheet.UsedRange.Rows.count
        For k = 1 To count
            ' 修正任务的负责人
            If tasks.exists(Trim(dist_sheet.Cells(k, 3).Value)) Then
                dist_sheet.Cells(k, 6).Value = tasks.Item(Trim(dist_sheet.Cells(k, 3).Value))
            End If
        Next
    Next
End Function

'复制sheet
Private Function CopySheets(wbook)
    Dim i, count As Integer
    Dim src_sheet As Worksheet
    
    For i = 1 To wbook.Sheets.count
        Set src_sheet = wbook.Sheets(i)
        If InStr(1, src_sheet.name, m_keyword) Then
            On Error Resume Next
            If ThisWorkbook.Sheets(src_sheet.name) Is Nothing Then '不存在sheet，直接复制sheet
                wbook.Sheets(src_sheet.name).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
            Else
                count = ThisWorkbook.Sheets(src_sheet.name).UsedRange.Rows.count
                wbook.Sheets(src_sheet.name).UsedRange.Copy ThisWorkbook.Sheets(src_sheet.name).Range("A" & (count + 1))
            End If
        End If
    Next
End Function

'是否清除sheet
Private Function NeedClearSheet(flag) As Boolean
    Dim ret As Boolean
    Let ret = False
    If Trim(flag) = "是" Then
        Let ret = True
    End If
    NeedClearSheet = ret
End Function

'过滤表头
Private Function FilterName(name_field) As Boolean
    Dim ret As Boolean
    Let ret = True
    If Trim(name_field) = "优先级" Then
        Let ret = False
    End If
    FilterName = ret
End Function

'匹配名字
Private Function MatchName(name, name_field) As Boolean
    Dim ret As Boolean
    Dim arr() As String
    Dim i As Integer
    
    Let ret = False
    Let name = Trim(name)
    Let name_field = Trim(name_field)
    Let arr() = Split(name, "|")
    If (name = "所有") Then
        Let ret = True
    Else
        For i = LBound(arr) To UBound(arr)
            If InStr(1, name_field, Trim(arr(i))) Then
                Let ret = True
                Exit For
            End If
        Next
    End If
    
    MatchName = ret
End Function

'匹配状态
Private Function MatchStat(stat, stat_field) As Boolean
    Dim ret As Boolean
    Dim arr() As String
    Dim i As Integer
    
    Let ret = False
    Let stat = Trim(stat)
    Let stat_field = Trim(stat_field)
    
    If (stat = "所有") Then
        Let ret = True
    Else
        Let arr() = Split(stat, "|")
        For i = LBound(arr) To UBound(arr)
            If Trim(arr(i)) = "空" And stat_field = "" Then
                Let ret = True
                Exit For
            ElseIf Trim(arr(i)) = "完成" And stat_field = "完成" Then
                Let ret = True
                Exit For
            ElseIf Trim(arr(i)) = "反馈" And stat_field <> "" And stat_field <> "完成" Then
                Let ret = True
                Exit For
            End If
        Next
    End If
    
    MatchStat = ret
End Function

'是否更新状态
Private Function NeedUpdateField(stat) As Boolean
    Dim ret As Boolean
    Let ret = False
    If Trim(stat) = "是" Then
        Let ret = True
    End If
    NeedUpdateField = ret
End Function

'创建sheet
Private Function CreateSheet(name) As Worksheet
    On Error Resume Next
    If ThisWorkbook.Sheets(name) Is Nothing Then
        ThisWorkbook.Sheets(m_template_name).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        ThisWorkbook.Sheets(ThisWorkbook.Sheets.count).name = name
    End If
    Set CreateSheet = ThisWorkbook.Sheets(name)
End Function

'清除sheet
Private Sub ClearSheet(name)
    Dim rsHeight As Integer
    On Error Resume Next
    If Not ThisWorkbook.Sheets(name) Is Nothing Then
'        ThisWorkbook.Sheets(name).Cells.Clear
'        ThisWorkbook.Sheets(name).Cells.ClearContents
        ThisWorkbook.Sheets(name).Cells.Delete

        ThisWorkbook.Sheets(m_template_name).UsedRange.Copy
        ThisWorkbook.Sheets(name).Range("A1").PasteSpecial Paste:=xlPasteValues
        ThisWorkbook.Sheets(name).Range("A1").PasteSpecial Paste:=xlPasteFormats
        ThisWorkbook.Sheets(name).Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
        rsHeight = ThisWorkbook.Sheets(m_template_name).Range("A1").EntireRow.Height
        ThisWorkbook.Sheets(name).Range("A1").RowHeight = rsHeight
        
    End If
End Sub

'获取配置
Private Function GetConfig(row) As String
    GetConfig = Trim(ThisWorkbook.Sheets(m_setting_name).Range(m_proj & row).Value)
End Function

'获取配置
Private Function GetOtherConfig(row, wbook) As String
    Dim col As String
    Let col = wbook.Sheets(m_setting_name).Range("A2").Value
    GetOtherConfig = Trim(wbook.Sheets(m_setting_name).Range(col & row).Value)
End Function

'获取配置
Private Function GetOtherCustomConfig(col, row, wbook) As String
    GetOtherCustomConfig = Trim(wbook.Sheets(m_setting_name).Range(col & row).Value)
End Function


