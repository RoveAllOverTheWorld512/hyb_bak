Attribute VB_Name = "股东户数与股价走势图"
'通达信路径
Public tdxdir As String     '通达信安装目录
Public shday As String      '通达信沪市日线数据目录
Public szday As String      '通达信深市日线数据目录
Public bkdir As String      '通达信板块数据目录
Public hqdir As String      '通达信行情缓冲区目录
Public svdir As String      '数据保存目录

Sub TDXPATH()
'获取通达信安装路径
    Dim objWMI As Object
    
    Const HKEY_LOCAL_MACHINE = &H80000002

    Set objWMI = GetObject("winmgmts:\\.\root\default:StdRegProv")
    objWMI.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\华西证券华彩人生", "InstallLocation", tdxdir
     
    shday = tdxdir & "vipdoc\sh\lday\"
    szday = tdxdir & "vipdoc\sz\lday\"
    bkdir = tdxdir & "T0002\blocknew\"
    hqdir = tdxdir & "T0002\hq_cache\"
    svdir = Left(ThisWorkbook.FullName, 2) & "\公司研究\"
     
End Sub

'将字节转换成字符串
Public Function ByteToStr(B() As Byte) As String 'Byte数组转字符串
    Dim i, tmp As String
    For Each i In B '枚举整个数组赋值给I
        If i > 127 Then '如果为汉字编码.(大于127为汉字,占两个字节)
            If tmp <> "" Then '如果临时变量不为空(为空为第一字节)
                ByteToStr = ByteToStr & Chr(tmp * 256 + i) '合并两个字节,转换为汉字.累加数据
                tmp = "" '清空临时变量
            Else
                tmp = i '储存临时变量
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
Sub aaaa()
    For Each wb In Workbooks
        MsgBox wb.FullName
    Next
    
End Sub

Sub gdhs_gj()
'
' 股东户数与股价走势图
'
'
    On Error Resume Next
    '避免Selection.SpecialCells(xlCellTypeFormulas, xlErrors) = "-"出错
    
    Dim fn As String, pth As String
    gpdm1 = Sheets("_xlwings.conf").Range("B8").Value
    gpmc1 = Sheets("_xlwings.conf").Range("B9").Value
    
    TDXPATH
    
    fn = svdir & gpmc1 & "\" & gpdm1 & gpmc1 & "股价走势分析.xlsx"
    
    pth = FullNameToPath(fn)
    
    MakeDir pth
    
    Application.DisplayAlerts = False
    
    For Each wb In Workbooks
        If wb.FullName = fn Then
            wb.Close
        End If
    Next
    Application.DisplayAlerts = True
    
    Sheets("股价与成交量").Select
    Range("A1").Select
    
    Range("E1").Value = "总股本（万股）"
    Range("F1").Value = "流通股本（万股）"
    Range("G1").Value = "实际流通股本（万股）"
    Range("H1").Value = "静态市盈率"
    Range("I1").Value = "动态市盈率"
    Range("J1").Value = "市净率"
    Range("K1").Value = "股东户数"
    Range("L1").Value = "户均市值(万元)"
    
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
        
    Selection.SpecialCells(xlCellTypeFormulas, xlErrors) = "-"

    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 
    Sheets("市盈率与市净率").Select
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

    pers = Range("A1").CurrentRegion.Rows.Count
    pecs = Range("A1").CurrentRegion.Columns.Count

    Sheets("股价与成交量").Select
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-7],市盈率与市净率!R1C1:R" & CStr(pers) & "C4,2,TRUE)"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-8],市盈率与市净率!R1C1:R" & CStr(pers) & "C4,3,TRUE)"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-9],市盈率与市净率!R1C1:R" & CStr(pers) & "C4,4,TRUE)"
  
    Range("H2:J2").Select
    Selection.AutoFill Destination:=Range("H2:J" & CStr(gjrs)), Type:=xlFillDefault
    
    Range("H2:J" & CStr(gjrs)).Select
    
    Selection.SpecialCells(xlCellTypeFormulas, xlErrors) = "-"
    
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 
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
    Range("K2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-10],股东户数!R2C1:R" & CStr(gdrs) & "C2,2,TRUE)"
    
    
    Range("K2").Select
    Selection.AutoFill Destination:=Range("K2:K" & CStr(gjrs)), Type:=xlFillDefault
   
    Range("K2:K" & CStr(gjrs)).Select
    '将无法查到的值用“-”替代，避免出现#N/A，在后面取最大值时出错
    Selection.SpecialCells(xlCellTypeFormulas, xlErrors) = "-"

    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=RC[-8]*RC[-7]/RC[-1]"
    Range("L2").Select
    Selection.NumberFormatLocal = "0.00_ "
    Selection.AutoFill Destination:=Range("L2:L" & CStr(gjrs)), Type:=xlFillDefault
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
    
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    nwb = ActiveWorkbook.Name
    
    Workbooks(owb).Activate
    Sheets("股东户数").Select
    Sheets("股东户数").Copy Before:=Workbooks(nwb).Sheets(1)
    
    Workbooks(owb).Activate
    Sheets("历年股本变动").Select
    Sheets("历年股本变动").Copy Before:=Workbooks(nwb).Sheets(1)
    
    Workbooks(owb).Activate
    Sheets("市盈率与市净率").Select
    Sheets("市盈率与市净率").Copy Before:=Workbooks(nwb).Sheets(1)
    
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "股东户数与股价走势图"
    
    gj_gdhs
    gj_vol
    gj_gb
    gj_pe
    gj_pb
    gj_hjsz
    
    ActiveWorkbook.Save

    
End Sub
Sub gj_vol()

    Sheets("股价与成交量").Select
    
    gjrs = Range("A1").CurrentRegion.Rows.Count
    gjcs = Range("A1").CurrentRegion.Columns.Count
   
    Sheets("股东户数与股价走势图").Select
    ActiveSheet.Shapes.AddChart(xlLine, 40, 360, 900, 350).Select  '添加一折线图
    
    With ActiveChart
        .SetSourceData Source:=Sheets("股价与成交量").Range("C2:C" & CStr(gjrs))
        .SeriesCollection(1).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(1).Name = "股价(前复权)"
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Values = "='股价与成交量'!$B$2:$B$" & CStr(gjrs)
        .SeriesCollection(2).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(2).Name = "成交量"
        .SeriesCollection(2).ChartType = xlColumnClustered
        
        .SeriesCollection(1).AxisGroup = 2
        
        .Legend.Position = xlBottom
        
        .SetElement (msoElementChartTitleAboveChart)
        .SetElement (msoElementPrimaryValueAxisTitleVertical)
        .SetElement (msoElementSecondaryValueAxisTitleVertical)
        .SetElement (msoElementPrimaryValueGridLinesMajor)
        .SetElement (msoElementSecondaryValueGridLinesMajor)
        
        With .Axes(xlCategory)
            .CategoryType = xlTimeScale
            .BaseUnit = xlDays
            .MajorUnitScale = xlMonths
            .MajorUnit = 3
            .TickLabels.Orientation = xlTickLabelOrientationUpward
            
            .MinimumScale = #1/1/2012#          '41275 2013年01月01日
            .MaximumScale = #12/31/2017#        '43100
        End With
        
        .HasTitle = True
        .ChartTitle.Text = gpdm1 & gpmc1 & "成交量与股价走势"
        
        With .Axes(xlValue, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 5
            .MajorGridlines.Border.LineStyle = xlDash
            .HasTitle = True
            .AxisTitle.Text = "成交量"
        End With
        
        With .Axes(xlValue, xlSecondary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 12
            .MajorGridlines.Border.LineStyle = xlDash
            .HasTitle = True
            .AxisTitle.Text = "股价"
        End With
        
        With .PlotArea
            .InsideTop = 40
            .InsideLeft = 50
            .InsideHeight = 220
            .InsideWidth = 780
        End With
    End With
    Range("A1").Select
End Sub
Sub gj_gdhs()

    Sheets("股价与成交量").Select
    
    gjrs = Range("A1").CurrentRegion.Rows.Count
    gjcs = Range("A1").CurrentRegion.Columns.Count
   
    If gjrs - 1000 < 0 Then
        qsh = 2
    Else
        qsh = gjrs - 1000
    End If
    
    rng = Range("K" & CStr(qsh) & ":K" & CStr(gjrs))
    gdhsmax = WorksheetFunction.Max(rng)
   
    Sheets("股东户数与股价走势图").Select
    ActiveSheet.Shapes.AddChart(xlLine, 40, 10, 900, 350).Select   '添加一折线图
    
    With ActiveChart
        .SetSourceData Source:=Sheets("股价与成交量").Range("C2:C" & CStr(gjrs))
        .SeriesCollection(1).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(1).Name = "股价(前复权)"
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Values = "='股价与成交量'!$K$2:$K$" & CStr(gjrs)
        .SeriesCollection(2).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(2).Name = "股东户数"
        
        .SeriesCollection(1).AxisGroup = 2
        
        .Legend.Position = xlBottom
        
        .SetElement (msoElementChartTitleAboveChart)
        .SetElement (msoElementPrimaryValueAxisTitleVertical)
        .SetElement (msoElementSecondaryValueAxisTitleVertical)
        .SetElement (msoElementPrimaryValueGridLinesMajor)
        .SetElement (msoElementSecondaryValueGridLinesMajor)
        .SetElement (msoElementPrimaryCategoryGridLinesMajor)
        
        With .Axes(xlCategory)
            .CategoryType = xlTimeScale
            .BaseUnit = xlDays
            .MajorUnitScale = xlMonths
            .MajorUnit = 3
            .TickLabels.Orientation = xlTickLabelOrientationUpward
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .MajorGridlines.Border.ColorIndex = 15
            
            .MinimumScale = #1/1/2012#          '41275 2013年01月01日
            .MaximumScale = #12/31/2017#        '43100
        End With
        
        .HasTitle = True
        .ChartTitle.Text = gpdm1 & gpmc1 & "股东户数与股价走势"
        
        With .Axes(xlValue, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 5
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "股东户数"
            .MinimumScale = 0
            .MaximumScale = WorksheetFunction.RoundUp(gdhsmax / 1000, 0) * 1000
        End With
        
        With .Axes(xlValue, xlSecondary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 12
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "股价"
        End With
        
        With .PlotArea
            .InsideTop = 40
            .InsideLeft = 50
            .InsideHeight = 220
            .InsideWidth = 780
        End With
    End With
    Range("A1").Select
End Sub
Sub gj_hjsz()
'户均市值
    Sheets("股价与成交量").Select
    
    gjrs = Range("A1").CurrentRegion.Rows.Count
    gjcs = Range("A1").CurrentRegion.Columns.Count
   
    If gjrs - 1000 < 0 Then
        qsh = 2
    Else
        qsh = gjrs - 1000
    End If
    
    rng = Range("L" & CStr(qsh) & ":L" & CStr(gjrs))
    hjszmax = WorksheetFunction.Max(rng)
   
    Sheets("股东户数与股价走势图").Select
    ActiveSheet.Shapes.AddChart(xlLine, 40, 1760, 900, 350).Select   '添加一折线图
    
    With ActiveChart
        .SetSourceData Source:=Sheets("股价与成交量").Range("C2:C" & CStr(gjrs))
        .SeriesCollection(1).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(1).Name = "股价(前复权)"
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Values = "='股价与成交量'!$L$2:$L$" & CStr(gjrs)
        .SeriesCollection(2).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(2).Name = "户均市值(万元)"
        
        .SeriesCollection(1).AxisGroup = 2
        
        .Legend.Position = xlBottom
        
        .SetElement (msoElementChartTitleAboveChart)
        .SetElement (msoElementPrimaryValueAxisTitleVertical)
        .SetElement (msoElementSecondaryValueAxisTitleVertical)
        .SetElement (msoElementPrimaryValueGridLinesMajor)
        .SetElement (msoElementSecondaryValueGridLinesMajor)
        .SetElement (msoElementPrimaryCategoryGridLinesMajor)
        
        With .Axes(xlCategory)
            .CategoryType = xlTimeScale
            .BaseUnit = xlDays
            .MajorUnitScale = xlMonths
            .MajorUnit = 3
            .TickLabels.Orientation = xlTickLabelOrientationUpward
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .MajorGridlines.Border.ColorIndex = 15
           
            .MinimumScale = #1/1/2012#          '41275 2013年01月01日
            .MaximumScale = #12/31/2017#        '43100
        End With
        
        .HasTitle = True
        .ChartTitle.Text = gpdm1 & gpmc1 & "户均市值与股价走势"
        
        With .Axes(xlValue, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 5
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "户均市值"
            .MinimumScale = 0
            .MaximumScale = WorksheetFunction.RoundUp(hjszmax / 10, 0) * 10
        End With
        
        With .Axes(xlValue, xlSecondary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 12
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "股价"
        End With
        
        With .PlotArea
            .InsideTop = 40
            .InsideLeft = 50
            .InsideHeight = 220
            .InsideWidth = 780
        End With
    End With
    Range("A1").Select
End Sub

Sub gj_pe()

    Sheets("股价与成交量").Select
    
    gjrs = Range("A1").CurrentRegion.Rows.Count
    gjcs = Range("A1").CurrentRegion.Columns.Count
        
    If gjrs - 500 < 0 Then
        qsh = 2
    Else
        qsh = gjrs - 500
    End If
    
    rng = Range("I" & CStr(qsh) & ":I" & CStr(gjrs))
    pemax = WorksheetFunction.Max(rng)
    If pemax = 0 Then
        pemax = 10
    End If
    If pemax > 100 Then
        pemax = 100
    End If
   
    Sheets("股东户数与股价走势图").Select
    ActiveSheet.Shapes.AddChart(xlLine, 40, 1060, 900, 350).Select   '添加一折线图
    
    With ActiveChart
        .SetSourceData Source:=Sheets("股价与成交量").Range("H2:H" & CStr(gjrs))
        .SeriesCollection(1).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(1).Name = "静态市盈率"
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Values = "='股价与成交量'!$I$2:$I$" & CStr(gjrs)
        .SeriesCollection(2).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(2).Name = "滚动市盈率"
        
        .SeriesCollection.NewSeries
        .SeriesCollection(3).Values = "='股价与成交量'!$C$2:$C$" & CStr(gjrs)
        .SeriesCollection(3).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(3).Name = "股价(前复权)"
        .SeriesCollection(3).AxisGroup = 2
        
        .Legend.Position = xlBottom
        
        .SetElement (msoElementChartTitleAboveChart)
        .SetElement (msoElementPrimaryValueAxisTitleVertical)
        .SetElement (msoElementSecondaryValueAxisTitleVertical)
        .SetElement (msoElementPrimaryValueGridLinesMajor)
        .SetElement (msoElementSecondaryValueGridLinesMajor)
        .SetElement (msoElementPrimaryCategoryGridLinesMajor)
        
        With .Axes(xlCategory)
            .CategoryType = xlTimeScale
            .BaseUnit = xlDays
            .MajorUnitScale = xlMonths
            .MajorUnit = 3
            .TickLabels.Orientation = xlTickLabelOrientationUpward
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .MajorGridlines.Border.ColorIndex = 15
            
            .MinimumScale = #1/1/2012#          '41275 2013年01月01日
            .MaximumScale = #12/31/2017#        '43100
        End With
        
        .HasTitle = True
        .ChartTitle.Text = gpdm1 & gpmc1 & "市盈率与股价走势"
        
        With .Axes(xlValue, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 5
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "市盈率倍数"
            .MinimumScale = 0
            .MaximumScale = WorksheetFunction.RoundUp(pemax / 10, 0) * 10
        End With
        
        With .Axes(xlValue, xlSecondary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 12
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "股价"
        End With
        
        With .PlotArea
            .InsideTop = 40
            .InsideLeft = 50
            .InsideHeight = 220
            .InsideWidth = 780
        End With
    End With
    Range("A1").Select
End Sub
Sub gj_pb()

    Sheets("股价与成交量").Select
    
    gjrs = Range("A1").CurrentRegion.Rows.Count
    gjcs = Range("A1").CurrentRegion.Columns.Count
   
    If gjrs - 1000 < 0 Then
        qsh = 2
    Else
        qsh = gjrs - 1000
    End If
    
    rng = Range("J" & CStr(qsh) & ":J" & CStr(gjrs))
    
    pbmax = WorksheetFunction.Max(rng)
    If pbmax = 0 Then
        pbmax = 2
    End If
    If pbmax > 20 Then
        pbmax = 20
    End If
   
    Sheets("股东户数与股价走势图").Select
    ActiveSheet.Shapes.AddChart(xlLine, 40, 1410, 900, 350).Select   '添加一折线图
    
    With ActiveChart
        .SetSourceData Source:=Sheets("股价与成交量").Range("C2:C" & CStr(gjrs))
        .SeriesCollection(1).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(1).Name = "股价(前复权)"
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Values = "='股价与成交量'!$J$2:$J$" & CStr(gjrs)
        .SeriesCollection(2).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(2).Name = "市净率"
        
        .SeriesCollection(1).AxisGroup = 2
        
        .Legend.Position = xlBottom
        
        .SetElement (msoElementChartTitleAboveChart)
        .SetElement (msoElementPrimaryValueAxisTitleVertical)
        .SetElement (msoElementSecondaryValueAxisTitleVertical)
        .SetElement (msoElementPrimaryValueGridLinesMajor)
        .SetElement (msoElementSecondaryValueGridLinesMajor)
        .SetElement (msoElementPrimaryCategoryGridLinesMajor)
        
        With .Axes(xlCategory)
            .CategoryType = xlTimeScale
            .BaseUnit = xlDays
            .MajorUnitScale = xlMonths
            .MajorUnit = 3
            .TickLabels.Orientation = xlTickLabelOrientationUpward
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .MajorGridlines.Border.ColorIndex = 15
            
            .MinimumScale = #1/1/2012#          '41275 2013年01月01日
            .MaximumScale = #12/31/2017#        '43100
        End With
        
        .HasTitle = True
        .ChartTitle.Text = gpdm1 & gpmc1 & "市净率与股价走势"
        
        With .Axes(xlValue, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 5
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "市净率"
            .MinimumScale = 0
            .MaximumScale = WorksheetFunction.RoundUp(pbmax, 2)
        End With
        
        With .Axes(xlValue, xlSecondary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 12
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "股价"
        End With
        
        With .PlotArea
            .InsideTop = 40
            .InsideLeft = 50
            .InsideHeight = 220
            .InsideWidth = 780
        End With
    End With
    Range("A1").Select
End Sub
Sub gj_gb()

    Sheets("股价与成交量").Select
    
    gjrs = Range("A1").CurrentRegion.Rows.Count
    gjcs = Range("A1").CurrentRegion.Columns.Count
   
    Sheets("股东户数与股价走势图").Select
    ActiveSheet.Shapes.AddChart(xlColumnClustered, 40, 710, 900, 350).Select   '添加一柱状图
    
    With ActiveChart
        .SetSourceData Source:=Sheets("股价与成交量").Range("E2:E" & CStr(gjrs))
        .SeriesCollection(1).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(1).Name = "总股本"
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Values = "='股价与成交量'!$G$2:$G$" & CStr(gjrs)
        .SeriesCollection(2).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(2).Name = "实际流通股本"
        
        .SeriesCollection.NewSeries
        .SeriesCollection(3).Values = "='股价与成交量'!$C$2:$C$" & CStr(gjrs)
        .SeriesCollection(3).XValues = "='股价与成交量'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(3).Name = "股价(前复权)"
        .SeriesCollection(3).ChartType = xlLine
        .SeriesCollection(3).AxisGroup = 2
        
        .Legend.Position = xlBottom
        
        .SetElement (msoElementChartTitleAboveChart)
        .SetElement (msoElementPrimaryValueAxisTitleVertical)
        .SetElement (msoElementSecondaryValueAxisTitleVertical)
        .SetElement (msoElementPrimaryValueGridLinesMajor)
        .SetElement (msoElementSecondaryValueGridLinesMajor)
'        .SetElement (msoElementPrimaryCategoryGridLinesMajor)
        
        With .Axes(xlCategory)
            .CategoryType = xlTimeScale
            .BaseUnit = xlDays
            .MajorUnitScale = xlMonths
            .MajorUnit = 3
            .TickLabels.Orientation = xlTickLabelOrientationUpward
            
            .MinimumScale = #1/1/2012#          '41275 2013年01月01日
            .MaximumScale = #12/31/2017#        '43100
        End With
        
        .HasTitle = True
        .ChartTitle.Text = gpdm1 & gpmc1 & "总股本、实际流通股本与股价走势"
        
        With .Axes(xlValue, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 5
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "股本"
        End With
        
        With .Axes(xlValue, xlSecondary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 12
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "股价"
        End With
        
        With .PlotArea
            .InsideTop = 40
            .InsideLeft = 50
            .InsideHeight = 220
            .InsideWidth = 780
        End With
    End With
    Range("A1").Select
End Sub


'生成股票名称字典
Function gpmc_dic()
    Dim Header(1 To 50) As Byte
    Dim gpdm(1 To 6) As Byte
    Dim unknow1(1 To 17) As Byte
    Dim gpmc(1 To 8) As Byte
    Dim unknow2(1 To 283) As Byte
    
    Dim dm As String, mc As String
    
    TDXPATH
    
    Set gpmc_dic = CreateObject("Scripting.Dictionary")
    
    For n = 1 To 2
        fn = tdxhq & "s" & Mid("hz", n, 1) & "m.tnf"
    
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

Sub cs()
    Dim rng As Range
    Set rng = ActiveSheet.Cells.SpecialCells(xlCellTypeFormulas, xlErrors)
    MsgBox rng.Address
    
End Sub
