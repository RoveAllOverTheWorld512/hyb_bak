Attribute VB_Name = "ģ��3"
Option Explicit

Sub ɸѡ��ѡ��()
    Application.ScreenUpdating = False  '����ˢ�¹رգ�������������ٶ�
    Dim r As Integer, chrsht As String
    
    Dim MyRange As Range

    Dim i As Integer, j As Integer
    
    Dim wb As Workbook, wb1 As Workbook
    Dim curwbk As String, cursht As String, bk As String, sht As String
    Dim gs As String, s As String
    
    curwbk = ActiveWorkbook.Name    '��ǰ������
    cursht = ActiveSheet.Name       '��ǰ������
    
    Workbooks.Open Filename:="D:\hyb\��ѡ���б�.xlsx"       '��

        
    Sheets("��ѡ��").Select
    Cells(1, 3).Value = "���"
    
    r = ActiveSheet.UsedRange.Rows.Count    '�������
    For i = 2 To r
        Cells(i, 3).Value = i - 1       '��д���
    Next
        
    bk = ActiveWorkbook.Name               '��ȡ����������
    sht = ActiveSheet.Name                 '��ȡ����������
    gs = "=VLOOKUP(MID(RC[-2],1,6),[" & bk & "]" & sht & "!R1C1:R" & r & "C3,3,FALSE)"
        
    Workbooks(curwbk).Activate
    Sheets(cursht).Activate
    r = ActiveSheet.UsedRange.Rows.Count    '�������
    
    
    Cells(1, 3).Value = "���"
    
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
    ActiveWindow.FreezePanes = True '���ᴰ��

        
    Windows("��ѡ���б�.xlsx").Activate
    ActiveWindow.Close (False)

    Application.ScreenUpdating = True '����ˢ�¿�����

End Sub

Sub ɸѡ�ɶ��������ٸ���()
    Application.ScreenUpdating = False  '����ˢ�¹رգ�������������ٶ�
    Dim r As Integer, chrsht As String
    
    Dim MyRange As Range

    Dim i As Integer, j As Integer
    
    Dim wb As Workbook, wb1 As Workbook
    Dim curwbk As String, cursht As String, bk As String, sht As String
    Dim gs As String, s As String
    
    curwbk = ActiveWorkbook.Name    '��ǰ������
    cursht = ActiveSheet.Name       '��ǰ������
    
    Workbooks.Open Filename:="D:\hyb\��ѡ���б�.xlsx"       '��

        
    Sheets("��ѡ��").Select
    Cells(1, 3).Value = "���"
    
    r = ActiveSheet.UsedRange.Rows.Count    '�������
    For i = 2 To r
        Cells(i, 3).Value = i - 1       '��д���
    Next
        
    bk = ActiveWorkbook.Name               '��ȡ����������
    sht = ActiveSheet.Name                 '��ȡ����������
    gs = "=VLOOKUP(MID(RC[-2],1,6),[" & bk & "]" & sht & "!R1C1:R" & r & "C3,3,FALSE)"
        
    Workbooks(curwbk).Activate
    Sheets(cursht).Activate
    r = ActiveSheet.UsedRange.Rows.Count    '�������
    
    
    Cells(1, 3).Value = "���"
    
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
    
    Cells(1, 5).Value = "�ɶ���������������%"
    
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
    ActiveWindow.FreezePanes = True '���ᴰ��

        
    Windows("��ѡ���б�.xlsx").Activate
    ActiveWindow.Close (False)

    Application.ScreenUpdating = True '����ˢ�¿�����

End Sub


