Attribute VB_Name = "ģ��1"
Option Base 1

Sub gbgdgs()
    Dim gbarr(), gjarr(), gdarr()
    Dim i As Integer, j As Integer, k As Integer, pxq As String
    
    jjpx
    gjpx
    gdhspx
        
    Sheets("�ɼ���ɽ���").Select
    
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "�ܹɱ�(�ڹ�)"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "��ͨ�ɱ�(�ڹ�)"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "A�ɻ���������)"
    
    Range("A1").Select
       
    gjarr = Range("A1").CurrentRegion.Value
    
    pxq = Range("A1").CurrentRegion.Address
    
    gbarr = tqgb()
    
    j = 1
    k = UBound(gbarr)
    
    'i=1Ϊ������
    For i = 2 To UBound(gjarr)
        gjarr(i, 8) = (gbarr(j, 4) + gbarr(j, 9)) / 10 ^ 8
        gjarr(i, 9) = gbarr(j, 4) / 10 ^ 8
        If j < k Then
            If gjarr(i, 1) >= gbarr(j + 1, 1) Then
                j = j + 1
            End If
        End If
        
    Next
    
    gdarr = tqgdhs()
    
    j = 1
    k = UBound(gdarr)
    
    'i=1Ϊ������
    For i = 2 To UBound(gjarr)
        'iΪ���۱�ָ��
        'jΪ�ɶ���ָ��
        If j < k Then
            '����ɼ۱�ǰ���ڴ��ڹɶ���ǰ���ڣ���ɶ�������ǰ��1
            If gjarr(i, 1) > gdarr(j, 1) Then
                j = j + 1
            End If
        End If
        
        '����ɼ۱�ǰ���ڵ��ڹɶ���ǰ���ڣ���ɼ۱�ɶ���ֱ���ùɶ�����
        If gjarr(i, 1) = gdarr(j, 1) Then
            gjarr(i, 10) = gdarr(j, 3)
        Else
        
        
        '����ɼ۱����ڴ��ڵ��ڹɶ������ڣ���ɶ��������һ�ڹɶ���
        '����ɼ۱�����С�ڹɶ������ڣ���ɶ��������һ�ڹɶ���
            If gjarr(i, 1) >= gdarr(k, 1) Then
                gjarr(i, 10) = gdarr(k, 3)
            Else
                
                If j = 1 Then
                    gjarr(i, 10) = gdarr(1, 3)
                Else
                    gjarr(i, 10) = gdarr(j - 1, 3) + (gjarr(i, 1) - gdarr(j - 1, 1)) * gdarr(j, 11)
                End If
                
            End If
        End If
        
    Next
   
    
    Range(pxq).Value = gjarr
    
    k = UBound(gjarr)
    pxq = "H2:J" & CStr(k)
    Range(pxq).Select
    Selection.NumberFormatLocal = "#,##0.000"
    
    
        

End Sub
Sub jjpx()
'
' ���۹ɽ��ʱ���ʱ������
'
    Dim r As Integer, px As Range, pxqu As Range, cursht As String
    Dim i As Integer
    Dim ws As Worksheet
    cursht = ActiveSheet.Name
        
    Set ws = Sheets("���۹ɽ��ʱ���")
    ws.Select
    
    r = ws.UsedRange.Rows.Count
    For i = 1 To r
        If Cells(i, 2).Value = "���ǰ��ͨ��" Then
            Exit For
        End If
    Next
    
    Set pxqu = ws.Range("A" & i + 1 & ":I" & r)
    
    Set px = Range("A" & i + 1)
    
    pxqu.Select
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=px, _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange pxqu
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets(cursht).Select
    
End Sub

Sub gdhspx()
'
' �ɶ�������ʱ�����������
'
    Dim r As Integer, px As Range, pxqu As Range, cursht As String
    Dim i As Integer
    Dim ws As Worksheet
    
    cursht = ActiveSheet.Name
    
    Set ws = Sheets("�ɶ�����")
    ws.Select
    
    r = ws.UsedRange.Rows.Count
    For i = 1 To r
        If Cells(i, 2).Value = "�ܻ���" Then
            Exit For
        End If
    Next
    
    Set pxqu = ws.Range("A" & i + 1 & ":I" & r)
    
    Set px = Range("A" & i + 1)
    
   
    pxqu.Select
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=px, _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange pxqu
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Set pxqu = ws.Range("J" & i - 1 & ":J" & i)
    
    pxqu.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    
    ActiveCell.FormulaR1C1 = "����"
    
    Set pxqu = ws.Range("K" & i - 1 & ":K" & i)
    pxqu.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "ÿ��仯����"
    
    Set pxqu = ws.Range("I" & i - 1 & ":I" & i)
    pxqu.Select
    
    Selection.Copy
    
    Set pxqu = ws.Range("J" & i - 1 & ":K" & i)
    pxqu.Select
    
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False


    Set px = ws.Range("J" & i + 1)
    px.Select
    ActiveCell.FormulaR1C1 = "0"
    Set px = ws.Range("K" & i + 1)
    px.Select
    ActiveCell.FormulaR1C1 = "0"
    
    Set px = ws.Range("J" & i + 2)
    px.Select
    ActiveCell.FormulaR1C1 = "=RC[-9]-R[-1]C[-9]"
    
    Set px = ws.Range("K" & i + 2)
    px.Select
    ActiveCell.FormulaR1C1 = "=(RC[-9]-R[-1]C[-9])/RC[-1]"
        
    Set pxqu = ws.Range("J" & i + 2 & ":K" & r)
    
    Set px = ws.Range("J" & i + 2 & ":K" & i + 2)
    px.Select
    
    Selection.AutoFill Destination:=pxqu, Type:=xlFillDefault
    
        
    Set px = ws.Range("B" & i + 1)
    px.Select
    Selection.Copy
    
    Set pxqu = ws.Range("J" & i + 1 & ":K" & r)
    pxqu.Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Set pxqu = ws.Range("K" & i + 1 & ":K" & r)
    pxqu.Select
    Selection.NumberFormatLocal = "#,##0.000"

    
    Sheets(cursht).Select
    
End Sub


Public Function tqgb() As Variant
'
' ��ȡ�ɱ���Ϣ
'
    Dim r As Integer, pxqu As Range, cursht As String
    Dim i As Integer
    Dim ws As Worksheet
    cursht = ActiveSheet.Name
        
    Set ws = Sheets("���۹ɽ��ʱ���")
    ws.Select
    
    r = ws.UsedRange.Rows.Count
    For i = 1 To r
        If Cells(i, 2).Value = "���ǰ��ͨ��" Then
            Exit For
        End If
    Next
    
    Set pxqu = ws.Range("A" & i + 1 & ":I" & r)
    
    tqgb = pxqu.Value
    
    Sheets(cursht).Select

End Function

Public Function tqgdhs() As Variant
'
' ��ȡ�ɶ�������Ϣ
'
    Dim r As Integer, pxqu As Range, cursht As String
    Dim i As Integer
    Dim ws As Worksheet
    cursht = ActiveSheet.Name
        
    Set ws = Sheets("�ɶ�����")
    ws.Select
    
    r = ws.UsedRange.Rows.Count
    For i = 1 To r
        If Cells(i, 2).Value = "�ܻ���" Then
            Exit For
        End If
    Next
    
    Set pxqu = ws.Range("A" & i + 1 & ":K" & r)
    tqgdhs = pxqu.Value
    Sheets(cursht).Select

End Function

Sub gjpx()
'
' �ɼ۱�ʱ������
'
'
    Dim r As Integer, pxqu As String, cursht As String
    cursht = ActiveSheet.Name
    
    Sheets("�ɼ���ɽ���").Select
    
    pxqu = Range("A1").CurrentRegion.Address
    
    Range(pxqu).Select
    ActiveWorkbook.Worksheets("�ɼ���ɽ���").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�ɼ���ɽ���").Sort.SortFields.Add Key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�ɼ���ɽ���").Sort
        .SetRange Range(pxqu)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets(cursht).Select
    
End Sub


Sub gjzst(ByVal bgdate As Date, ByVal eddate As Date)
'
' �ɼ�����ͼ
'
'
    Dim rng0, rng1, rng2, rng3, rng4, rng5 As Range
    Dim ws As Worksheet
    Dim r, s As Integer, pxqu As String, cursht As String
    Dim arr()
'    Dim bgdate As String, eddate As String
    
'    bgdate = "2010/1/1"
'    eddate = "2017/12/31"
    
    cursht = ActiveSheet.Name
    
    Set ws = Sheets("�ɼ���ɽ���")
    ws.Select
    r = ActiveSheet.UsedRange.Rows.Count
    s = 2
    Set rng0 = ws.Range("A" & s & ":A" & r)
    Set rng1 = ws.Range("E" & s & ":E" & r)
    Set rng2 = ws.Range("J" & s & ":J" & r)
    Set rng3 = ws.Range("H" & s & ":H" & r)
    Set rng4 = ws.Range("I" & s & ":I" & r)
    Set rng5 = ws.Range("G" & s & ":G" & r)
    
    
    
    Sheets.Add After:=Sheets(Sheets.Count)
    
'    Sheets("Sheet1").Select
    
    Range("A1").Select
    
    ActiveSheet.DrawingObjects.Delete   'ɾ������ͼ�ζ���
    
    
    ActiveSheet.Shapes.AddChart(xlLine, 0, 0, 1100, 200).Select   '���һ����ͼ
    
    ActiveChart.SeriesCollection.NewSeries                                  '������ϵ�С����ش������ϵ�е� Series ����
    
    ActiveChart.SeriesCollection(1).Name = "�ɼ�"
    ActiveChart.SeriesCollection(1).Values = rng1
    ActiveChart.SeriesCollection(1).XValues = rng0

    ActiveChart.SeriesCollection.NewSeries                                  '������ϵ�С����ش������ϵ�е� Series ����
    ActiveChart.SeriesCollection(2).Name = "�ɶ�����"
    ActiveChart.SeriesCollection(2).Values = rng2
    ActiveChart.SeriesCollection(2).XValues = rng0
    ActiveChart.SeriesCollection(2).AxisGroup = 2           '��������
    
    With ActiveChart.Axes(xlCategory)
        .CategoryType = xlTimeScale
        .BaseUnit = xlDays
        .MajorUnit = 3
        .MajorUnitScale = xlMonths
        .MinorUnit = 1
        .MinorUnitScale = xlMonths
        .MinimumScale = DateValue(bgdate)
        .MaximumScale = DateValue(eddate)
        .HasMajorGridlines = True  'xlCategory��������࣬��X��
        .TickLabelPosition = xlTickLabelPositionNextToAxis  'xlTickLabelPositionNone  ˮƽ���ǩ����ʾ
        .TickLabels.Font.Size = 8   'ˮƽ���ǩ�����С
    End With

  
    ActiveChart.Axes(xlValue).HasMajorGridlines = Fasle     'xlValue��������ʾֵ����Y��
    ActiveChart.HasLegend = True    '��ʾͼ��
    ActiveChart.Legend.Position = xlLegendPositionTop   '��ʾ���ϲ�
    
    With ActiveChart.PlotArea
        .InsideTop = 20
        .InsideLeft = 40
        .InsideHeight = 110
        .InsideWidth = 1000
    End With
      
    ActiveChart.ChartArea.Border.Color = RGB(255, 255, 255)     '����ͼ�����߿�Ϊ��ɫ
        
  
    
    
    ActiveSheet.Shapes.AddChart(xlColumnClustered, 0, 190, 1100, 200).Select    '���һ��״ͼ
    
    ActiveChart.SeriesCollection.NewSeries                                  '������ϵ�С����ش������ϵ�е� Series ����
    
    ActiveChart.SeriesCollection(1).Name = "�ܹɱ�"
    ActiveChart.SeriesCollection(1).Values = rng3
    ActiveChart.SeriesCollection(1).XValues = rng0

    
    ActiveChart.SeriesCollection.NewSeries                                  '������ϵ�С����ش������ϵ�е� Series ����
    ActiveChart.SeriesCollection(2).Name = "��ͨ�ɱ�"
    ActiveChart.SeriesCollection(2).Values = rng4
    ActiveChart.SeriesCollection(2).XValues = rng0
    
    ActiveChart.SeriesCollection.NewSeries                                  '������ϵ�С����ش������ϵ�е� Series ����
    ActiveChart.SeriesCollection(3).Name = "�ɽ���"
    ActiveChart.SeriesCollection(3).Values = rng5
    ActiveChart.SeriesCollection(3).XValues = rng0
    ActiveChart.SeriesCollection(3).AxisGroup = 2           '��������
'    ActiveChart.Axes(xlValue, xlSecondary).MaximumScale = 60000000  '����������ε����ֵ

    With ActiveChart.Axes(xlCategory)
        .CategoryType = xlTimeScale
        .BaseUnit = xlDays
        .MajorUnit = 3
        .MajorUnitScale = xlMonths
        .MinorUnit = 1
        .MinorUnitScale = xlMonths
        .MinimumScale = bgdate
        .MaximumScale = eddate
        .HasMajorGridlines = True  'xlCategory��������࣬��X��
        .TickLabelPosition = xlTickLabelPositionNextToAxis  'xlTickLabelPositionNone  ˮƽ���ǩ����ʾ
        .TickLabels.Font.Size = 8   'ˮƽ���ǩ�����С
    End With

     
    
    ActiveChart.Axes(xlValue).HasMajorGridlines = Fasle     'xlValue��������ʾֵ����Y��
    ActiveChart.HasLegend = True    '��ʾͼ��
    ActiveChart.Legend.Position = xlLegendPositionBottom   '��ʾ�ڵײ�

    With ActiveChart.PlotArea
        .InsideTop = 10
        .InsideLeft = 40
        .InsideHeight = 110
        .InsideWidth = 1000
    End With

    ActiveChart.ChartArea.Border.Color = RGB(255, 255, 255)     '����ͼ�����߿�Ϊ��ɫ

End Sub


'Callback for customButton1 onAction
Sub sczst(control As IRibbonControl)
    Dim ws As Worksheet
    
    On Error GoTo err
    
    Set ws = Sheets("�ɼ���ɽ���")
    ws.Select
    createimg.Show
    Exit Sub
    
err:
    MsgBox "û�С��ɼ���ɽ��������޷�ִ�У�"
    
End Sub

'Callback for customButton1 onAction
Sub fzxg(control As IRibbonControl)
    selectstock.Show
End Sub

