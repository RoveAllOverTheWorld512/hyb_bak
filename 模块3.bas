Attribute VB_Name = "模块3"
Option Explicit

Sub 筛选自选股()
    Application.ScreenUpdating = False  '设置刷新关闭，可以提高运行速度
    Dim r As Integer, chrsht As String
    
    Dim MyRange As Range

    Dim i As Integer, j As Integer
    
    Dim wb As Workbook, wb1 As Workbook
    Dim curwbk As String, cursht As String, bk As String, sht As String
    Dim gs As String, s As String
    
    curwbk = ActiveWorkbook.Name    '当前工作簿
    cursht = ActiveSheet.Name       '当前工作表
    
    Workbooks.Open Filename:="D:\hyb\自选股列表.xlsx"       '打开

        
    Sheets("自选股").Select
    Cells(1, 3).Value = "序号"
    
    r = ActiveSheet.UsedRange.Rows.Count    '最大行数
    For i = 2 To r
        Cells(i, 3).Value = i - 1       '填写编号
    Next
        
    bk = ActiveWorkbook.Name               '获取工作簿名称
    sht = ActiveSheet.Name                 '获取工作表名称
    gs = "=VLOOKUP(MID(RC[-2],1,6),[" & bk & "]" & sht & "!R1C1:R" & r & "C3,3,FALSE)"
        
    Workbooks(curwbk).Activate
    Sheets(cursht).Activate
    r = ActiveSheet.UsedRange.Rows.Count    '最大行数
    
    
    Cells(1, 3).Value = "序号"
    
    Range("C2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .NumberFormatLocal = "0"
    End With

    Range("C2").Select
    ActiveCell.FormulaR1C1 = gs
        
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C" & r), Type:=xlFillDefault
    
    Range("C2").Select
    s = Split(ActiveCell.CurrentRegion.Address, ":")(1)

    Range("C2").Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "C2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A2:" & s)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("D2").Select
    ActiveWindow.FreezePanes = True '冻结窗口

        
    Windows("自选股列表.xlsx").Activate
    ActiveWindow.Close (False)

    Application.ScreenUpdating = True '设置刷新开启度

End Sub

Sub 筛选股东人数减少个股()
    Application.ScreenUpdating = False  '设置刷新关闭，可以提高运行速度
    Dim r As Integer, chrsht As String
    
    Dim MyRange As Range

    Dim i As Integer, j As Integer
    
    Dim wb As Workbook, wb1 As Workbook
    Dim curwbk As String, cursht As String, bk As String, sht As String
    Dim gs As String, s As String
    
    curwbk = ActiveWorkbook.Name    '当前工作簿
    cursht = ActiveSheet.Name       '当前工作表
    
    Workbooks.Open Filename:="D:\hyb\自选股列表.xlsx"       '打开

        
    Sheets("自选股").Select
    Cells(1, 3).Value = "序号"
    
    r = ActiveSheet.UsedRange.Rows.Count    '最大行数
    For i = 2 To r
        Cells(i, 3).Value = i - 1       '填写编号
    Next
        
    bk = ActiveWorkbook.Name               '获取工作簿名称
    sht = ActiveSheet.Name                 '获取工作表名称
    gs = "=VLOOKUP(MID(RC[-2],1,6),[" & bk & "]" & sht & "!R1C1:R" & r & "C3,3,FALSE)"
        
    Workbooks(curwbk).Activate
    Sheets(cursht).Activate
    r = ActiveSheet.UsedRange.Rows.Count    '最大行数
    
    
    Cells(1, 3).Value = "序号"
    
    Range("C2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .NumberFormatLocal = "0"
    End With

    Range("C2").Select
    ActiveCell.FormulaR1C1 = gs
        
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C" & r), Type:=xlFillDefault
    
    Range("C2").Select
    s = Split(ActiveCell.CurrentRegion.Address, ":")(1)

    Range("C2").Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "C2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A2:" & s)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Cells(1, 5).Value = "股东户数环比增长率%"
    
    Range("E2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .NumberFormatLocal = "0.00"
    End With

    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR((RC[1]/RC[2]-1)*100,"""")"
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E" & r), Type:=xlFillDefault

    Range("D2").Select
    ActiveWindow.FreezePanes = True '冻结窗口

        
    Windows("自选股列表.xlsx").Activate
    ActiveWindow.Close (False)

    Application.ScreenUpdating = True '设置刷新开启度

End Sub


