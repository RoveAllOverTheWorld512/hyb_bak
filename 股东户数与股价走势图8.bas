Attribute VB_Name = "股东户数与股价走势图"

'将字节转换成字符串
Public Function ByteToStr(B() As Byte) As String 'Byte数组转字符串
    Dim i, Tmp As String
    For Each i In B '枚举整个数组赋值给I
        If i > 127 Then '如果为汉字编码.(大于127为汉字,占两个字节)
            If Tmp <> "" Then '如果临时变量不为空(为空为第一字节)
                ByteToStr = ByteToStr & Chr(Tmp * 256 + i) '合并两个字节,转换为汉字.累加数据
                Tmp = "" '清空临时变量
            Else
                Tmp = i '储存临时变量
            End If
        Else
            ByteToStr = ByteToStr & Chr(i) '转换为汉字,累加数据
        End If
    Next
End Function

Sub All_Str2Date()
'将当前单元所在块所有日期列字符串转换成日期型
    Dim i As Integer, j As Integer
    Dim rng As Range, numfmt As String
    Dim curregadd As String, lt As String, rb As String
    Dim ltr As Integer, ltc As Integer, rbr As Integer, rbc As Integer
    Dim data, data1
    Application.ScreenUpdating = False  '设置刷新关闭，可以提高运行速度
    
    ActiveCell.CurrentRegion.Select
    curregadd = Selection.Address
    
    If Application.WorksheetFunction.CountA(Selection) = 0 Then
        MsgBox "请将光标放在排序区域。", vbOKOnly
        Exit Sub
    End If
    
    lt = Split(curregadd, ":")(0)       '左上角
    rb = Split(curregadd, ":")(1)       '右下角
    
    ltr = Range(lt).EntireRow.Row               '起始行
    ltc = Range(lt).EntireColumn.Column         '起始列
    rbr = Range(rb).EntireRow.Row               '终止行
    rbc = Range(rb).EntireColumn.Column         '终止列
    
    Set rng = Range(Cells(ltr, ltc), Cells(ltr, rbc))
    data = rng.Value
    For i = 1 To UBound(data, 2)
        If InStr(Replace(data(1, i), " ", ""), "日") > 0 Then
            For j = ltr To rbr
                Set rng = Cells(j, ltc + i - 1)
                data1 = rng.Value
                numfmt = rng.NumberFormat
                If rng.NumberFormat <> "yyyy-mm-dd;@" Then
                    rng.NumberFormatLocal = "yyyy-mm-dd;@"
                End If
                '单元格式为“常规"
                If numfmt = "General" Then
                    If VarType(data1) = vbDouble Then
                        data1 = CStr(data1)
                        If Len(data1) = 8 Then
                            data1 = Left(data1, 4) & "-" & Mid(data1, 5, 2) & "-" & Right(data1, 2)
                        End If
                    End If
                End If
                '单元格式为“文本”
                If numfmt = "@" Then
                    If Len(data1) = 8 Then
                        data1 = Left(data1, 4) & "-" & Mid(data1, 5, 2) & "-" & Right(data1, 2)
                    End If
                End If
                
                'IsDate函数，它判断表达式是否可以转换为日期格式而不是说数据类型是不是日期型
                'vartype(varname)函数指出变量的子类型,varname 参数是一个 Variant，包含用户定义类型变量之外的任何变量。
                    
                If IsDate(data1) Then
                    rng.Value = DateAdd("d", 0, data1)
                End If
            Next
        End If
    Next
    Application.ScreenUpdating = True  '设置刷新关闭，可以提高运行速度
    

End Sub
Function FullNameToPath(sFullName As String) As String
  '取路径， 不包括后面的反斜杠
  Dim k As Integer
  For k = Len(sFullName) To 1 Step -1
    If Mid(sFullName, k, 1) = "\" Then Exit For
  Next k
  If k < 1 Then
    FullNameToPath = ""
  Else
    FullNameToPath = Mid(sFullName, 1, k - 1)
  End If
End Function
'创建多级子目录
Sub MakeDir(Path As String)
    On Error Resume Next
    Dim s As String
    Dim i As Integer
    Dim v As Variant
    Dim sarr() As String
    sarr() = Split(Path, "\")
    i = 0
    For Each v In sarr()
        i = i + 1
        If i = 1 Then
            s = v
        Else
            s = s & "\" & v
            MkDir s
        End If
    Next
End Sub


Sub gdhs_gj()
'
' 股东户数与股价走势图
'
'
    Dim fn As String, pth As String
    gpdm1 = Sheets("_xlwings.conf").Range("B8").Value
    gpmc1 = Sheets("_xlwings.conf").Range("B9").Value
    
    fn = "D:\公司研究\" & gpmc1 & "\" & gpdm1 & gpmc1 & "股价走势分析.xlsx"
    pth = FullNameToPath(fn)
    
    MakeDir pth
    
    Sheets("股价与成交量").Select
    Range("A1").Select
    
    Range("D1").Value = "股东户数"
    Range("E1").Value = "总股本（万股）"
    Range("F1").Value = "流通股本（万股）"
    Range("G1").Value = "实际流通股本（万股）"
    
    gjrs = Range("A1").CurrentRegion.Rows.Count
    gjcs = Range("A1").CurrentRegion.Columns.Count
   
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A1").CurrentRegion
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    Sheets("股东户数").Select
    
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A1").CurrentRegion
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    gdrs = Range("A1").CurrentRegion.Rows.Count
    gdcs = Range("A1").CurrentRegion.Columns.Count
    
    Sheets("股价与成交量").Select
    Range("D2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-3],股东户数!R2C1:R" & CStr(gdrs) & "C2,2,TRUE)"
    
    
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D" & CStr(gjrs)), Type:=xlFillDefault
   
    Range("D2:D" & CStr(gjrs)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("历年股本变动").Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A1").CurrentRegion
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    gbrs = Range("A1").CurrentRegion.Rows.Count
    gbcs = Range("A1").CurrentRegion.Columns.Count

    Sheets("股价与成交量").Select
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],历年股本变动!R1C1:R" & CStr(gbrs) & "C4,2,TRUE)"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],历年股本变动!R1C1:R" & CStr(gbrs) & "C4,3,TRUE)"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-6],历年股本变动!R1C1:R" & CStr(gbrs) & "C4,4,TRUE)"
  
    Range("E2:G2").Select
    Selection.AutoFill Destination:=Range("E2:G" & CStr(gjrs)), Type:=xlFillDefault
    
    Range("E2:G" & CStr(gjrs)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    owb = ActiveWorkbook.Name
    
    Application.DisplayAlerts = False
    Sheets("股价与成交量").Copy
    ActiveWorkbook.SaveAs Filename:=fn, FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    Application.DisplayAlerts = True
    nwb = ActiveWorkbook.Name
    Workbooks(owb).Activate
    
    Sheets("股东户数").Select
    Sheets("股东户数").Copy Before:=Workbooks(nwb).Sheets(1)
    Workbooks(owb).Activate
    Sheets("历年股本变动").Select
    Sheets("历年股本变动").Copy Before:=Workbooks(nwb).Sheets(1)
    
    Sheets("股价与成交量").Select
    
    gjrs = Range("A1").CurrentRegion.Rows.Count
    gjcs = Range("A1").CurrentRegion.Columns.Count
   
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "股东户数与股价走势图"
    
    ActiveSheet.Shapes.AddChart(xlLine, 40, 10, 900, 350).Select   '添加一折线图

    
    ActiveChart.SetSourceData Source:=Sheets("股价与成交量").Range("C2:C" & CStr(gjrs))
    ActiveChart.SeriesCollection(1).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
    ActiveChart.SeriesCollection(1).Name = "股价(前复权)"
    
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Values = "='股价与成交量'!$D$2:$D$" & CStr(gjrs)
    ActiveChart.SeriesCollection(2).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
    ActiveChart.SeriesCollection(2).Name = "股东户数"
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).AxisGroup = 2
    ActiveChart.Legend.Position = xlBottom
    
    ActiveChart.Axes(xlCategory).CategoryType = xlTimeScale
    ActiveChart.Axes(xlCategory).BaseUnit = xlDays
    ActiveChart.Axes(xlCategory).MajorUnitScale = xlMonths
    ActiveChart.Axes(xlCategory).MajorUnit = 1
    
    ActiveChart.Axes(xlCategory).MinimumScale = #1/1/2013#          '41275 2013年01月01日
    ActiveChart.Axes(xlCategory).MaximumScale = #12/31/2017#        '43100
    
    ActiveChart.Axes(xlValue).HasTitle = True
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleVertical)
    ActiveChart.SetElement (msoElementSecondaryValueAxisTitleVertical)
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "股东户数"
'    ActiveChart.SetElement (msoElementSecondaryValueAxisTitleRotated)
    ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Text = "股价(前复权)"
    
    ActiveChart.HasTitle = True
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    ActiveChart.ChartTitle.Text = gpdm1 & gpmc1 & "股东户数与股价走势关系"
    
    ActiveChart.Axes(xlCategory).TickLabels.Orientation = xlTickLabelOrientationUpward
    
    ActiveChart.Axes(xlValue, xlPrimary).HasMajorGridlines = True
    ActiveChart.Axes(xlValue, xlPrimary).MajorGridlines.Border.ColorIndex = 5
    ActiveChart.Axes(xlValue, xlPrimary).MajorGridlines.Border.LineStyle = xlDash
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesMajor)
    ActiveChart.Axes(xlValue, xlSecondary).HasMajorGridlines = True
    ActiveChart.Axes(xlValue, xlSecondary).MajorGridlines.Border.ColorIndex = 12
    ActiveChart.Axes(xlValue, xlSecondary).MajorGridlines.Border.LineStyle = xlDash
    ActiveChart.SetElement (msoElementSecondaryValueGridLinesMajor)
    
    ActiveChart.ChartArea.Select
    With ActiveChart.PlotArea
        .InsideTop = 40
        .InsideLeft = 50
        .InsideHeight = 220
        .InsideWidth = 780
    End With
    
    
    ActiveSheet.Shapes.AddChart(xlLine, 40, 360, 900, 350).Select   '添加一折线图
    
    ActiveChart.SetSourceData Source:=Sheets("股价与成交量").Range("C2:C" & CStr(gjrs))
    ActiveChart.SeriesCollection(1).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
    ActiveChart.SeriesCollection(1).Name = "股价(前复权)"
    
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Values = "='股价与成交量'!$B$2:$B$" & CStr(gjrs)
    ActiveChart.SeriesCollection(2).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
    ActiveChart.SeriesCollection(2).Name = "成交量"
    ActiveChart.SeriesCollection(2).ChartType = xlColumnClustered
    
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).AxisGroup = 2
    ActiveChart.Legend.Position = xlBottom
    
    ActiveChart.Axes(xlCategory).CategoryType = xlTimeScale
    ActiveChart.Axes(xlCategory).BaseUnit = xlDays
    ActiveChart.Axes(xlCategory).MajorUnitScale = xlMonths
    ActiveChart.Axes(xlCategory).MajorUnit = 1
    
    ActiveChart.Axes(xlCategory).MinimumScale = #1/1/2013#          '41275 2013年01月01日
    ActiveChart.Axes(xlCategory).MaximumScale = #12/31/2017#        '43100
    
    ActiveChart.Axes(xlValue).HasTitle = True
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleVertical)
    ActiveChart.SetElement (msoElementSecondaryValueAxisTitleVertical)
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "成交量"
    ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Text = "股价(前复权)"
    
    ActiveChart.HasTitle = True
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    ActiveChart.ChartTitle.Text = gpdm1 & gpmc1 & "成交量与股价走势关系"
    
    ActiveChart.Axes(xlCategory).TickLabels.Orientation = xlTickLabelOrientationUpward
    
    ActiveChart.Axes(xlValue, xlPrimary).HasMajorGridlines = True
    ActiveChart.Axes(xlValue, xlPrimary).MajorGridlines.Border.ColorIndex = 5
    ActiveChart.Axes(xlValue, xlPrimary).MajorGridlines.Border.LineStyle = xlDash
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesMajor)
    ActiveChart.Axes(xlValue, xlSecondary).HasMajorGridlines = True
    ActiveChart.Axes(xlValue, xlSecondary).MajorGridlines.Border.ColorIndex = 12
    ActiveChart.Axes(xlValue, xlSecondary).MajorGridlines.Border.LineStyle = xlDash
    ActiveChart.SetElement (msoElementSecondaryValueGridLinesMajor)
    
    ActiveChart.ChartArea.Select
    With ActiveChart.PlotArea
        .InsideTop = 40
        .InsideLeft = 50
        .InsideHeight = 220
        .InsideWidth = 780
    End With
    
    ActiveSheet.Shapes.AddChart(xlColumnClustered, 40, 710, 900, 350).Select   '添加一柱状图
    
    ActiveChart.SetSourceData Source:=Sheets("股价与成交量").Range("E2:E" & CStr(gjrs))
    ActiveChart.SeriesCollection(1).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
    ActiveChart.SeriesCollection(1).Name = "总股本"
    
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Values = "='股价与成交量'!$G$2:$G$" & CStr(gjrs)
    ActiveChart.SeriesCollection(2).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
    ActiveChart.SeriesCollection(2).Name = "实际流通股本"
    
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(3).Values = "='股价与成交量'!$C$2:$C$" & CStr(gjrs)
    ActiveChart.SeriesCollection(3).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
    ActiveChart.SeriesCollection(3).Name = "股价(前复权)"
    ActiveChart.SeriesCollection(3).ChartType = xlLine
    
    ActiveChart.SeriesCollection(3).Select
    ActiveChart.SeriesCollection(3).AxisGroup = 2
    ActiveChart.Legend.Position = xlBottom
    
    ActiveChart.Axes(xlCategory).CategoryType = xlTimeScale
    ActiveChart.Axes(xlCategory).BaseUnit = xlDays
    ActiveChart.Axes(xlCategory).MajorUnitScale = xlMonths
    ActiveChart.Axes(xlCategory).MajorUnit = 1
    
    ActiveChart.Axes(xlCategory).MinimumScale = #1/1/2013#          '41275 2013年01月01日
    ActiveChart.Axes(xlCategory).MaximumScale = #12/31/2017#        '43100
    
    ActiveChart.Axes(xlValue).HasTitle = True
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleVertical)
    ActiveChart.SetElement (msoElementSecondaryValueAxisTitleVertical)
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "股本"
    ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Text = "股价"
    
    ActiveChart.HasTitle = True
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    ActiveChart.ChartTitle.Text = gpdm1 & gpmc1 & "总股本、实际流通股本与股价走势"
    
    ActiveChart.Axes(xlCategory).TickLabels.Orientation = xlTickLabelOrientationUpward
    
    ActiveChart.Axes(xlValue, xlPrimary).HasMajorGridlines = True
    ActiveChart.Axes(xlValue, xlPrimary).MajorGridlines.Border.ColorIndex = 5
    ActiveChart.Axes(xlValue, xlPrimary).MajorGridlines.Border.LineStyle = xlDash
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesMajor)
    ActiveChart.Axes(xlValue, xlSecondary).HasMajorGridlines = True
    ActiveChart.Axes(xlValue, xlSecondary).MajorGridlines.Border.ColorIndex = 12
    ActiveChart.Axes(xlValue, xlSecondary).MajorGridlines.Border.LineStyle = xlDash
    ActiveChart.SetElement (msoElementSecondaryValueGridLinesMajor)
    
    ActiveChart.ChartArea.Select
    With ActiveChart.PlotArea
        .InsideTop = 40
        .InsideLeft = 50
        .InsideHeight = 220
        .InsideWidth = 780
    End With
    
    ActiveWorkbook.Save

    
End Sub

'生成股票名称字典
Function gpmc_dic()
    Dim Header(1 To 50) As Byte
    Dim gpdm(1 To 6) As Byte
    Dim unknow1(1 To 17) As Byte
    Dim gpmc(1 To 8) As Byte
    Dim unknow2(1 To 283) As Byte
    
    Dim dm As String, mc As String
    Set gpmc_dic = CreateObject("Scripting.Dictionary")
    
    For n = 1 To 2
        fn = "C:\new_hxzq_hc\T0002\hq_cache\s" & Mid("hz", n, 1) & "m.tnf"
    
        Open fn For Binary As #1 '打开文本文件
        
        Get #1, , Header
        Do
            Get #1, , gpdm
            dm = ByteToStr(gpdm) & ".S" & UCase(Mid("hz", n, 1))
            
            Get #1, , unknow1
            Get #1, , gpmc
            For i = 1 To 8
                If gpmc(i) = 0 Then
                    gpmc(i) = 32    'x00的用x20空格代替
                End If
            Next
            mc = ByteToStr(gpmc)
            Get #1, , unknow2
            If (n = 1 And Left(dm, 1) = "6") Or (n = 2 And (Left(dm, 1) = "0" Or Left(dm, 2) = "30")) Then
            
                gpmc_dic.Add dm, Replace(Replace(mc, " ", ""), "*", "")
            
            End If
             
        Loop Until EOF(1)
        
        Close #1 '关闭文件
    Next
    
End Function

Sub get_data()
    mymodule = Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1))
    RunPython ("import " & mymodule & ";" & mymodule & ".getdata()")
End Sub


