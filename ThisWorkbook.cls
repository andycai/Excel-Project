VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub Pull(proj)
    Dim task As TeamTask
    Set task = New TeamTask
    task.Init (proj)
    task.PullTasks
End Sub

'修正任务
Function ReviseTasks()
    Dim wbook As Workbook
    Dim src_sheet As Worksheet
    Dim i, j, k, rowcount, count As Integer
    
    Dim tasks As Object
    Set tasks = CreateObject("Scripting.Dictionary")
    
    Set wbook = Workbooks.Open(ThisWorkbook.Path & "\..\修正任务.xlsx")
    For i = 1 To wbook.Sheets.count
        Set src_sheet = wbook.Sheets(i)
        Let rowcount = src_sheet.Range("A65536").End(xlUp).row
        For j = 1 To rowcount
            If src_sheet.Cells(j, 3).Value <> "内容" Then
                tasks.Item(src_sheet.Cells(j, 3).Value) = Trim(src_sheet.Cells(j, 6).Value)
            End If
        Next
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
