Attribute VB_Name = "模块2"
Option Explicit
Public chartsht As String

Sub guben(control As IRibbonControl)
    gbjs.Show
End Sub
Sub test()
    Dim i As Integer, j As Integer
    For j = 21 To 30
        Cells(j, 2).NumberFormatLocal = "yyyy-mm-dd"
        Cells(j, 2).Value = DateValue(Cells(j, 2).Value)
    Next
   
End Sub

Sub pecl(ByVal bgdate As Date, ByVal eddate As Date)
'
' 估值分析表数据整理与排序
'
    Dim r As Integer, px As Range, pxqu As Range, cursht As String, chrsht As String
    
    Dim rng0, rng1, rng2, rng3, rng4, rng5 As Range
'    Dim bgdate As Date, eddate As Date

    Dim i As Integer, j As Integer
    Dim mx As Double, mn As Double
    
    Dim ws As Worksheet, ws1 As Worksheet
    cursht = ActiveSheet.Name
        
    Set ws = Sheets("估值分析")
    ws.Select
    
    r = ws.UsedRange.Rows.Count
    For i = 1 To r
        If Cells(i, 2).Value = "截止日期" Then
            Exit For
        End If
    Next
    
    
    '设置日期的格式
    '将文本转换成日期值,不能用Cdate函数
    For j = i + 1 To r
        Cells(j, 2).NumberFormatLocal = "yyyy-mm-dd"
        Cells(j, 2).Value = DateValue(Cells(j, 2).Value)
    Next
    
    '将文本转用分列方法换成数值,用Cdbl函数对不能转换的会出错
    Set pxqu = ws.Range("C" & i + 1 & ":C" & r)
    Set px = Range("C" & i + 1)
    pxqu.Select

'    MsgBox Application.IsText(px)
'   MsgBox VarType(px)

    
    
    Selection.NumberFormatLocal = "0.00_ "
    
    Selection.TextToColumns Destination:=px, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    
    
    Set pxqu = ws.Range("A" & i + 1 & ":G" & r)
    Set px = Range("B" & i + 1)
        
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
        
    For j = i + 1 To r
        If Cells(j, 2) >= bgdate Then
            Exit For
        End If
    Next
    i = j
     
    For j = i + 1 To r
        If Cells(j, 2) >= eddate Then
            Exit For
        End If
    Next
    If j < r Then
        r = j
    End If
   
    
    Set rng0 = ws.Range("B" & i + 1 & ":B" & r)
    Set rng1 = ws.Range("C" & i + 1 & ":C" & r)
    
    rng1.Select
    
    mx = Application.Max(Selection)

    mn = Application.Min(Selection)
    mn = Application.Min(0, mn)

    
    Sheets.Add After:=Sheets(Sheets.Count)
    chartsht = ActiveSheet.Name
    
    
    Range("A1").Select
    
    ActiveSheet.DrawingObjects.Delete   '删除所有图形对象
    
    
    ActiveSheet.Shapes.AddChart(xlLine, 0, 0, 1100, 200).Select   '添加一折线图
    With ActiveChart.PlotArea
        .InsideTop = 10
        .InsideLeft = 50
        .InsideHeight = 110
        .InsideWidth = 1000
        
    End With
    
    MsgBox ActiveChart.Name
    ActiveChart.HasTitle = False    '放在ActiveChart.HasLegend = True不起作用
    
    ActiveChart.SeriesCollection.NewSeries                                  '创建新系列。返回代表该新系列的 Series 对象
    
    ActiveChart.SeriesCollection(1).Name = "PE(TTM)"
    ActiveChart.SeriesCollection(1).Values = rng1
    ActiveChart.SeriesCollection(1).XValues = rng0
      
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

  
    ActiveChart.Axes(xlValue).HasMajorGridlines = False     'xlValue坐标轴显示值，即Y轴
    ActiveChart.Axes(xlValue).CrossesAt = mn
        
    ActiveChart.HasLegend = True    '显示图例
    ActiveChart.Legend.Position = xlLegendPositionBottom   '显示在下部
    
     
    ActiveChart.ChartArea.Border.Color = RGB(255, 255, 255)     '设置图标区边框为白色
        
      
    ActiveChart.HasTitle = False    '放在ActiveChart.HasLegend = True不起作用
       
   
    Sheets(cursht).Select
    
End Sub

Sub gjzs()

    Dim r As Integer, px As Range, pxqu As Range, cursht As String, chrsht As String
    
    Dim rng0, rng1, rng2, rng3, rng4, rng5 As Range
    Dim bgdate As Date, eddate As Date

    Dim i As Integer, j As Integer
    Dim mx As Double, mn As Double
    
    Dim ws As Worksheet, ws1 As Worksheet
    Dim chartname As String
    
    
    cursht = ActiveSheet.Name
    
    bgdate = DateValue("2012-10-01")
    eddate = DateValue("2017-10-31")
    
    Set ws1 = Sheets("股价与成交量")
    ws1.Select
    
    r = ws1.UsedRange.Rows.Count
    For j = 2 To r
        If Cells(j, 1) >= bgdate Then
            Exit For
        End If
    Next
    i = j
    
    For j = i + 1 To r
        If Cells(j, 1) >= eddate Then
            Exit For
        End If
    Next
    If j < r Then
        r = j
    End If

    Set rng2 = ws1.Range("A" & i + 1 & ":A" & r)
    Set rng3 = ws1.Range("E" & i + 1 & ":E" & r)
    
    Sheets.Add After:=Sheets(Sheets.Count)
    chrsht = ActiveSheet.Name
'    Sheets(chartsht).Select
    ActiveSheet.Shapes.AddChart(xlLine, 0, 0, 1100, 200).Select   '添加一折线图
    chartname = ActiveChart.Name
    
    
    
    With ActiveChart.PlotArea
        .InsideTop = 10
        .InsideLeft = 50
        .InsideHeight = 110
        .InsideWidth = 1000
        
    End With
    
  
    ActiveChart.Axes(xlValue).HasMajorGridlines = False     'xlValue坐标轴显示值，即Y轴
    ActiveChart.Axes(xlValue).CrossesAt = mn
    ActiveChart.HasTitle = False                '放在ActiveChart.HasLegend = True不起作用
       
    ActiveChart.HasLegend = True    '显示图例
    ActiveChart.Legend.Position = xlLegendPositionBottom        '显示在下部
        
      
    ActiveChart.ChartArea.Border.Color = RGB(255, 255, 255)     '设置图标区边框为白色
    ActiveChart.ChartArea.Format.Fill.Visible = msoFalse        '图表区域无填充
    ActiveChart.PlotArea.Format.Fill.Visible = msoFalse         '绘图区域无填充
      
    
    
    ActiveChart.SeriesCollection.NewSeries                                  '创建新系列。返回代表该新系列的 Series 对象
    ActiveChart.SeriesCollection(1).Name = "股价"
    ActiveChart.SeriesCollection(1).Values = rng3
    ActiveChart.SeriesCollection(1).XValues = rng2
    
    ActiveChart.SeriesCollection(1).Select
    With Selection.Border
         .ColorIndex = 3                    '线条颜色设置为红色
        .Weight = xlThick                   '线条宽度
        .LineStyle = xlContinuous           '线条样式
    End With
    
    With ActiveChart.Axes(xlCategory)
        .CategoryType = xlTimeScale
        .BaseUnit = xlDays
        .MajorUnit = 3
        .MajorUnitScale = xlMonths
        .MinorUnit = 1
        .MinorUnitScale = xlMonths
        .MinimumScale = bgdate
        .MaximumScale = eddate
        .Crosses = xlMaximum
        .HasMajorGridlines = True  'xlCategory坐标轴分类，即X轴
        .TickLabelPosition = xlTickLabelPositionNextToAxis  'xlTickLabelPositionNone  水平轴标签不显示
        .TickLabels.Font.Size = 8   '水平轴标签字体大小
    End With
    
    With Sheets(chrsht).ChartObjects(1).Chart
        .HasTitle = False
    End With

End Sub

Sub 设置图表标题()
    With Sheets("sheet6").ChartObjects(1).Chart
        .HasTitle = True
        .ChartTitle.Text = "我设置的图表标题"
    End With
End Sub

Sub 关闭图表标题()
    With Sheets("sheet6").ChartObjects(1).Chart
        .HasTitle = False
    End With
End Sub

Sub tmp()

    MsgBox ActiveChart.Name

    MsgBox ActiveChart.ChartTitle.Text
   
End Sub

