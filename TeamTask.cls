VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TeamTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_proj, m_template_name, m_setting_name As String
Private isclear As Boolean
Private dict_clear As Object

Private m_handler, m_acceptor, m_handle_stat, m_accept_stat, m_keyword As String
Private m_dirname, m_filename, m_type, m_update_handle, m_update_accept, m_sheetname, m_clear_flag As String
Private m_append As String

'��ʼ��
Public Sub Init(proj)
    m_template_name = "ģ��"
    m_setting_name = "˵��"
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
    m_append = GetConfig(14)
End Sub

'��ȡ����
Public Sub PullTasks()
    Dim book As Workbook
    Dim f, file(), activefile, ret, sheetname As String
    
    Let activefile = ActiveWorkbook.name
    Let sheetname = ""
    
    isclear = False
    Set dict_clear = CreateObject("Scripting.Dictionary")
    
    ReDim file(1)
    Let file(1) = ThisWorkbook.Path & "\..\" & m_dirname & "\"
  
    If m_filename = "����" Then
        Let f = Dir(file(1) & "*.xlsm")     'ͨ���*.*��ʾ�����ļ���*.xlsx Excel�ļ�
    Else
        Let f = Dir(file(1) & m_filename & ".xlsm")
    End If
    Do Until f = ""
        If f <> activefile Then
            Set book = Workbooks.Open(file(1) & f)
            If m_type = "����" Then
                sheetname = Left(f, InStr(1, f, ".") - 1)
            ElseIf m_type = "����" Then
                sheetname = m_sheetname
            End If
            Let ret = CopyTask(book, sheetname)
'            Let ret = CopySheets(book)
            book.Close
        End If
        Let f = Dir
    Loop
End Sub

'��������
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
                    'ϵͳ�����sheet
                    If sheetname = "" Then
                        Set dist_sheet = CreateSheet(src_sheet.name) 'sheet name �� source һ��
                        
                        '���sheet
                        If NeedClearSheet(m_clear_flag) And Not dict_clear.exists(src_sheet.name) Then
                            ClearSheet (src_sheet.name)
                            dict_clear(src_sheet.name) = 1
                        End If
                    Else '����sheet
                        Set dist_sheet = CreateSheet(sheetname) '�Զ���sheet name
                        '���sheet
                        If NeedClearSheet(m_clear_flag) And isclear = False Then
                            ClearSheet (sheetname)
                            Let isclear = True
                        End If
                    End If
                         
                    Let count = dist_sheet.UsedRange.Rows.count
                    Let found = False
                    
                    If Not NeedClearSheet(m_clear_flag) Then
                        For k = 1 To count
                            If Trim(dist_sheet.Cells(k, 3).Value) = Trim(src_sheet.Cells(j, 3).Value) Then '������ͬ������
                                If NeedUpdateField(m_update_handle) Then
                                    src_sheet.Range("D" & j).Copy dist_sheet.Range("D" & k) '����������
                                End If
                                If NeedUpdateField(m_update_accept) Then
                                    src_sheet.Range("E" & j).Copy dist_sheet.Range("E" & k) '�����������
                                End If
                                found = True
                                Exit For
                            End If
                        Next
                    End If
                    
                    If Not found And NeedAppend(m_append) Then
                        src_sheet.Rows(j & ":" & j).Copy dist_sheet.Range("A" & (count + 1)) '���Ƶ�ָ��λ��
                    End If
                End If
            Next
        End If
    Next
End Function

'����sheet
Private Function CopySheets(wbook)
    Dim i, count As Integer
    Dim src_sheet As Worksheet
    
    For i = 1 To wbook.Sheets.count
        Set src_sheet = wbook.Sheets(i)
        If InStr(1, src_sheet.name, m_keyword) Then
            On Error Resume Next
            If ThisWorkbook.Sheets(src_sheet.name) Is Nothing Then '������sheet��ֱ�Ӹ���sheet
                wbook.Sheets(src_sheet.name).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
            Else
                count = ThisWorkbook.Sheets(src_sheet.name).UsedRange.Rows.count
                wbook.Sheets(src_sheet.name).UsedRange.Copy ThisWorkbook.Sheets(src_sheet.name).Range("A" & (count + 1))
            End If
        End If
    Next
End Function

'�Ƿ����sheet
Private Function NeedClearSheet(flag) As Boolean
    Dim ret As Boolean
    Let ret = False
    If Trim(flag) = "��" Then
        Let ret = True
    End If
    NeedClearSheet = ret
End Function

'�Ƿ�׷���µ���
Private Function NeedAppend(flag) As Boolean
    Dim ret As Boolean
    Let ret = False
    If Trim(flag) = "��" Then
        Let ret = True
    End If
    NeedAppend = ret
End Function

'���˱�ͷ
Private Function FilterName(name_field) As Boolean
    Dim ret As Boolean
    Let ret = True
    If Trim(name_field) = "���ȼ�" Then
        Let ret = False
    End If
    FilterName = ret
End Function

'ƥ������
Private Function MatchName(name, name_field) As Boolean
    Dim ret As Boolean
    Dim arr() As String
    Dim i As Integer
    
    Let ret = False
    Let name = Trim(name)
    Let name_field = Trim(name_field)
    Let arr() = Split(name, "|")
    If (name = "����") Then
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

'ƥ��״̬
Private Function MatchStat(stat, stat_field) As Boolean
    Dim ret As Boolean
    Dim arr() As String
    Dim i As Integer
    
    Let ret = False
    Let stat = Trim(stat)
    Let stat_field = Trim(stat_field)
    
    If (stat = "����") Then
        Let ret = True
    Else
        Let arr() = Split(stat, "|")
        For i = LBound(arr) To UBound(arr)
            If Trim(arr(i)) = "��" And stat_field = "" Then
                Let ret = True
                Exit For
            ElseIf Trim(arr(i)) = "���" And stat_field = "���" Then
                Let ret = True
                Exit For
            ElseIf Trim(arr(i)) = "����" And stat_field <> "" And stat_field <> "���" Then
                Let ret = True
                Exit For
            End If
        Next
    End If
    
    MatchStat = ret
End Function

'�Ƿ����״̬
Private Function NeedUpdateField(stat) As Boolean
    Dim ret As Boolean
    Let ret = False
    If Trim(stat) = "��" Then
        Let ret = True
    End If
    NeedUpdateField = ret
End Function

'����sheet
Private Function CreateSheet(name) As Worksheet
    On Error Resume Next
    If ThisWorkbook.Sheets(name) Is Nothing Then
        ThisWorkbook.Sheets(m_template_name).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        ThisWorkbook.Sheets(ThisWorkbook.Sheets.count).name = name
    End If
    Set CreateSheet = ThisWorkbook.Sheets(name)
End Function

'���sheet
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

'��ȡ����
Private Function GetConfig(row) As String
    GetConfig = Trim(ThisWorkbook.Sheets(m_setting_name).Range(m_proj & row).Value)
End Function

'��ȡ����
Private Function GetOtherConfig(row, wbook) As String
    Dim col As String
    Let col = wbook.Sheets(m_setting_name).Range("A2").Value
    GetOtherConfig = Trim(wbook.Sheets(m_setting_name).Range(col & row).Value)
End Function

'��ȡ����
Private Function GetOtherCustomConfig(col, row, wbook) As String
    GetOtherCustomConfig = Trim(wbook.Sheets(m_setting_name).Range(col & row).Value)
End Function
