Attribute VB_Name = "模块1"
Option Base 1

Sub gbgdgs()
    Dim gbarr(), gjarr(), gdarr()
    Dim i As Integer, j As Integer, k As Integer, pxq As String
    
    jjpx
    gjpx
    gdhspx
        
    Sheets("股价与成交量").Select
    
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "总股本(亿股)"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "流通股本(亿股)"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "A股户数（估算)"
    
    Range("A1").Select
       
    gjarr = Range("A1").CurrentRegion.Value
    
    pxq = Range("A1").CurrentRegion.Address
    
    gbarr = tqgb()
    
    j = 1
    k = UBound(gbarr)
    
    'i=1为标题行
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
    
    'i=1为标题行
    For i = 2 To UBound(gjarr)
        'i为估价表指针
        'j为股东表指针
        If j < k Then
            '如果股价表当前日期大于股东表当前日期，则股东表日期前移1
            If gjarr(i, 1) > gdarr(j, 1) Then
                j = j + 1
            End If
        End If
        
        '如果股价表当前日期等于股东表当前日期，则股价表股东数直接用股东表户数
        If gjarr(i, 1) = gdarr(j, 1) Then
            gjarr(i, 10) = gdarr(j, 3)
        Else
        
        
        '如果股价表日期大于等于股东表日期，则股东数用最后一期股东数
        '如果股价表日期小于股东表日期，则股东数用最后一期股东数
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
' 限售股解禁时间表按时间排序
'
    Dim r As Integer, px As Range, pxqu As Range, cursht As String
    Dim i As Integer
    Dim ws As Worksheet
    cursht = ActiveSheet.Name
        
    Set ws = Sheets("限售股解禁时间表")
    ws.Select
    
    r = ws.UsedRange.Rows.Count
    For i = 1 To r
        If Cells(i, 2).Value = "解禁前流通股" Then
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
' 股东户数表按时间排序与估算
'
    Dim r As Integer, px As Range, pxqu As Range, cursht As String
    Dim i As Integer
    Dim ws As Worksheet
    
    cursht = ActiveSheet.Name
    
    Set ws = Sheets("股东户数")
    ws.Select
    
    r = ws.UsedRange.Rows.Count
    For i = 1 To r
        If Cells(i, 2).Value = "总户数" Then
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
    
    ActiveCell.FormulaR1C1 = "天数"
    
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
    ActiveCell.FormulaR1C1 = "每天变化户数"
    
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
' 提取股本信息
'
    Dim r As Integer, pxqu As Range, cursht As String
    Dim i As Integer
    Dim ws As Worksheet
    cursht = ActiveSheet.Name
        
    Set ws = Sheets("限售股解禁时间表")
    ws.Select
    
    r = ws.UsedRange.Rows.Count
    For i = 1 To r
        If Cells(i, 2).Value = "解禁前流通股" Then
            Exit For
        End If
    Next
    
    Set pxqu = ws.Range("A" & i + 1 & ":I" & r)
    
    tqgb = pxqu.Value
    
    Sheets(cursht).Select

End Function

Public Function tqgdhs() As Variant
'
' 提取股东户数信息
'
    Dim r As Integer, pxqu As Range, cursht As String
    Dim i As Integer
    Dim ws As Worksheet
    cursht = ActiveSheet.Name
        
    Set ws = Sheets("股东户数")
    ws.Select
    
    r = ws.UsedRange.Rows.Count
    For i = 1 To r
        If Cells(i, 2).Value = "总户数" Then
            Exit For
        End If
    Next
    
    Set pxqu = ws.Range("A" & i + 1 & ":K" & r)
    tqgdhs = pxqu.Value
    Sheets(cursht).Select

End Function

Sub gjpx()
'
' 股价表按时间排序
'
'
    Dim r As Integer, pxqu As String, cursht As String
    cursht = ActiveSheet.Name
    
    Sheets("股价与成交量").Select
    
    pxqu = Range("A1").CurrentRegion.Address
    
    Range(pxqu).Select
    ActiveWorkbook.Worksheets("股价与成交量").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("股价与成交量").Sort.SortFields.Add Key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("股价与成交量").Sort
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
' 股价走势图
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
    
    Set ws = Sheets("股价与成交量")
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
    
    ActiveSheet.DrawingObjects.Delete   '删除所有图形对象
    
    
    ActiveSheet.Shapes.AddChart(xlLine, 0, 0, 1100, 200).Select   '添加一折线图
    
    ActiveChart.SeriesCollection.NewSeries                                  '创建新系列。返回代表该新系列的 Series 对象
    
    ActiveChart.SeriesCollection(1).Name = "股价"
    ActiveChart.SeriesCollection(1).Values = rng1
    ActiveChart.SeriesCollection(1).XValues = rng0

    ActiveChart.SeriesCollection.NewSeries                                  '创建新系列。返回代表该新系列的 Series 对象
    ActiveChart.SeriesCollection(2).Name = "股东户数"
    ActiveChart.SeriesCollection(2).Values = rng2
    ActiveChart.SeriesCollection(2).XValues = rng0
    ActiveChart.SeriesCollection(2).AxisGroup = 2           '次坐标轴
    
    With ActiveChart.Axes(xlCategory)
        .CategoryType = xlTimeScale
        .BaseUnit = xlDays
        .MajorUnit = 3
        .MajorUnitScale = xlMonths
        .MinorUnit = 1
        .MinorUnitScale = xlMonths
        .MinimumScale = DateValue(bgdate)
        .MaximumScale = DateValue(eddate)
        .HasMajorGridlines = True  'xlCategory坐标轴分类，即X轴
        .TickLabelPosition = xlTickLabelPositionNextToAxis  'xlTickLabelPositionNone  水平轴标签不显示
        .TickLabels.Font.Size = 8   '水平轴标签字体大小
    End With

  
    ActiveChart.Axes(xlValue).HasMajorGridlines = Fasle     'xlValue坐标轴显示值，即Y轴
    ActiveChart.HasLegend = True    '显示图例
    ActiveChart.Legend.Position = xlLegendPositionTop   '显示在上部
    
    With ActiveChart.PlotArea
        .InsideTop = 20
        .InsideLeft = 40
        .InsideHeight = 110
        .InsideWidth = 1000
    End With
      
    ActiveChart.ChartArea.Border.Color = RGB(255, 255, 255)     '设置图标区边框为白色
        
  
    
    
    ActiveSheet.Shapes.AddChart(xlColumnClustered, 0, 190, 1100, 200).Select    '添加一柱状图
    
    ActiveChart.SeriesCollection.NewSeries                                  '创建新系列。返回代表该新系列的 Series 对象
    
    ActiveChart.SeriesCollection(1).Name = "总股本"
    ActiveChart.SeriesCollection(1).Values = rng3
    ActiveChart.SeriesCollection(1).XValues = rng0

    
    ActiveChart.SeriesCollection.NewSeries                                  '创建新系列。返回代表该新系列的 Series 对象
    ActiveChart.SeriesCollection(2).Name = "流通股本"
    ActiveChart.SeriesCollection(2).Values = rng4
    ActiveChart.SeriesCollection(2).XValues = rng0
    
    ActiveChart.SeriesCollection.NewSeries                                  '创建新系列。返回代表该新系列的 Series 对象
    ActiveChart.SeriesCollection(3).Name = "成交量"
    ActiveChart.SeriesCollection(3).Values = rng5
    ActiveChart.SeriesCollection(3).XValues = rng0
    ActiveChart.SeriesCollection(3).AxisGroup = 2           '次坐标轴
'    ActiveChart.Axes(xlValue, xlSecondary).MaximumScale = 60000000  '设置纵坐标次的最大值

    With ActiveChart.Axes(xlCategory)
        .CategoryType = xlTimeScale
        .BaseUnit = xlDays
        .MajorUnit = 3
        .MajorUnitScale = xlMonths
        .MinorUnit = 1
        .MinorUnitScale = xlMonths
        .MinimumScale = bgdate
        .MaximumScale = eddate
        .HasMajorGridlines = True  'xlCategory坐标轴分类，即X轴
        .TickLabelPosition = xlTickLabelPositionNextToAxis  'xlTickLabelPositionNone  水平轴标签不显示
        .TickLabels.Font.Size = 8   '水平轴标签字体大小
    End With

     
    
    ActiveChart.Axes(xlValue).HasMajorGridlines = Fasle     'xlValue坐标轴显示值，即Y轴
    ActiveChart.HasLegend = True    '显示图例
    ActiveChart.Legend.Position = xlLegendPositionBottom   '显示在底部

    With ActiveChart.PlotArea
        .InsideTop = 10
        .InsideLeft = 40
        .InsideHeight = 110
        .InsideWidth = 1000
    End With

    ActiveChart.ChartArea.Border.Color = RGB(255, 255, 255)     '设置图标区边框为白色

End Sub


'Callback for customButton1 onAction
Sub sczst(control As IRibbonControl)
    Dim ws As Worksheet
    
    On Error GoTo err
    
    Set ws = Sheets("股价与成交量")
    ws.Select
    createimg.Show
    Exit Sub
    
err:
    MsgBox "没有【股价与成交量】表，无法执行！"
    
End Sub

'Callback for customButton1 onAction
Sub fzxg(control As IRibbonControl)
    selectstock.Show
End Sub

