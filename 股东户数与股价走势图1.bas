Attribute VB_Name = "�ɶ�������ɼ�����ͼ"
'ͨ����·��
Public tdxdir As String     'ͨ���Ű�װĿ¼
Public shday As String      'ͨ���Ż�����������Ŀ¼
Public szday As String      'ͨ����������������Ŀ¼
Public bkdir As String      'ͨ���Ű������Ŀ¼
Public hqdir As String      'ͨ�������黺����Ŀ¼
Public svdir As String      '���ݱ���Ŀ¼

Sub TDXPATH()
'��ȡͨ���Ű�װ·��
    Dim objWMI As Object
    
    Const HKEY_LOCAL_MACHINE = &H80000002

    Set objWMI = GetObject("winmgmts:\\.\root\default:StdRegProv")
    objWMI.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\����֤ȯ��������", "InstallLocation", tdxdir
     
    shday = tdxdir & "vipdoc\sh\lday\"
    szday = tdxdir & "vipdoc\sz\lday\"
    bkdir = tdxdir & "T0002\blocknew\"
    hqdir = tdxdir & "T0002\hq_cache\"
    svdir = Left(ThisWorkbook.FullName, 2) & "\��˾�о�\"
     
End Sub

'���ֽ�ת�����ַ���
Public Function ByteToStr(B() As Byte) As String 'Byte����ת�ַ���
    Dim i, tmp As String
    For Each i In B 'ö���������鸳ֵ��I
        If i > 127 Then '���Ϊ���ֱ���.(����127Ϊ����,ռ�����ֽ�)
            If tmp <> "" Then '�����ʱ������Ϊ��(Ϊ��Ϊ��һ�ֽ�)
                ByteToStr = ByteToStr & Chr(tmp * 256 + i) '�ϲ������ֽ�,ת��Ϊ����.�ۼ�����
                tmp = "" '�����ʱ����
            Else
                tmp = i '������ʱ����
            End If
        Else
            ByteToStr = ByteToStr & Chr(i) 'ת��Ϊ����,�ۼ�����
        End If
    Next
End Function

Sub All_Str2Date()
'����ǰ��Ԫ���ڿ������������ַ���ת����������
    Dim i As Integer, j As Integer
    Dim rng As Range, numfmt As String
    Dim curregadd As String, lt As String, rb As String
    Dim ltr As Integer, ltc As Integer, rbr As Integer, rbc As Integer
    Dim data, data1
    Application.ScreenUpdating = False  '����ˢ�¹رգ�������������ٶ�
    
    ActiveCell.CurrentRegion.Select
    curregadd = Selection.Address
    
    If Application.WorksheetFunction.CountA(Selection) = 0 Then
        MsgBox "�뽫��������������", vbOKOnly
        Exit Sub
    End If
    
    lt = Split(curregadd, ":")(0)       '���Ͻ�
    rb = Split(curregadd, ":")(1)       '���½�
    
    ltr = Range(lt).EntireRow.Row               '��ʼ��
    ltc = Range(lt).EntireColumn.Column         '��ʼ��
    rbr = Range(rb).EntireRow.Row               '��ֹ��
    rbc = Range(rb).EntireColumn.Column         '��ֹ��
    
    Set rng = Range(Cells(ltr, ltc), Cells(ltr, rbc))
    data = rng.Value
    For i = 1 To UBound(data, 2)
        If InStr(Replace(data(1, i), " ", ""), "��") > 0 Then
            For j = ltr To rbr
                Set rng = Cells(j, ltc + i - 1)
                data1 = rng.Value
                numfmt = rng.NumberFormat
                If rng.NumberFormat <> "yyyy-mm-dd;@" Then
                    rng.NumberFormatLocal = "yyyy-mm-dd;@"
                End If
                '��Ԫ��ʽΪ������"
                If numfmt = "General" Then
                    If VarType(data1) = vbDouble Then
                        data1 = CStr(data1)
                        If Len(data1) = 8 Then
                            data1 = Left(data1, 4) & "-" & Mid(data1, 5, 2) & "-" & Right(data1, 2)
                        End If
                    End If
                End If
                '��Ԫ��ʽΪ���ı���
                If numfmt = "@" Then
                    If Len(data1) = 8 Then
                        data1 = Left(data1, 4) & "-" & Mid(data1, 5, 2) & "-" & Right(data1, 2)
                    End If
                End If
                
                'IsDate���������жϱ��ʽ�Ƿ����ת��Ϊ���ڸ�ʽ������˵���������ǲ���������
                'vartype(varname)����ָ��������������,varname ������һ�� Variant�������û��������ͱ���֮����κα�����
                    
                If IsDate(data1) Then
                    rng.Value = DateAdd("d", 0, data1)
                End If
            Next
        End If
    Next
    Application.ScreenUpdating = True  '����ˢ�¹رգ�������������ٶ�
    

End Sub
Function FullNameToPath(sFullName As String) As String
'ȡ·���� ����������ķ�б��
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
'�����༶��Ŀ¼
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
' �ɶ�������ɼ�����ͼ
'
'
    On Error Resume Next
    '����Selection.SpecialCells(xlCellTypeFormulas, xlErrors) = "-"����
    
    Dim fn As String, pth As String
    gpdm1 = Sheets("_xlwings.conf").Range("B8").Value
    gpmc1 = Sheets("_xlwings.conf").Range("B9").Value
    
    TDXPATH
    
    fn = svdir & gpmc1 & "\" & gpdm1 & gpmc1 & "�ɼ����Ʒ���.xlsx"
    
    pth = FullNameToPath(fn)
    
    MakeDir pth
    
    Application.DisplayAlerts = False
    
    For Each wb In Workbooks
        If wb.FullName = fn Then
            wb.Close
        End If
    Next
    Application.DisplayAlerts = True
    
    Sheets("�ɼ���ɽ���").Select
    Range("A1").Select
    
    Range("E1").Value = "�ܹɱ�����ɣ�"
    Range("F1").Value = "��ͨ�ɱ�����ɣ�"
    Range("G1").Value = "ʵ����ͨ�ɱ�����ɣ�"
    Range("H1").Value = "��̬��ӯ��"
    Range("I1").Value = "��̬��ӯ��"
    Range("J1").Value = "�о���"
    Range("K1").Value = "�ɶ�����"
    Range("L1").Value = "������ֵ(��Ԫ)"
    
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
    
    Sheets("����ɱ��䶯").Select
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

    Sheets("�ɼ���ɽ���").Select
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],����ɱ��䶯!R1C1:R" & CStr(gbrs) & "C4,2,TRUE)"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],����ɱ��䶯!R1C1:R" & CStr(gbrs) & "C4,3,TRUE)"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-6],����ɱ��䶯!R1C1:R" & CStr(gbrs) & "C4,4,TRUE)"
  
    Range("E2:G2").Select
    Selection.AutoFill Destination:=Range("E2:G" & CStr(gjrs)), Type:=xlFillDefault
    
    Range("E2:G" & CStr(gjrs)).Select
        
    Selection.SpecialCells(xlCellTypeFormulas, xlErrors) = "-"

    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 
    Sheets("��ӯ�����о���").Select
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

    Sheets("�ɼ���ɽ���").Select
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-7],��ӯ�����о���!R1C1:R" & CStr(pers) & "C4,2,TRUE)"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-8],��ӯ�����о���!R1C1:R" & CStr(pers) & "C4,3,TRUE)"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-9],��ӯ�����о���!R1C1:R" & CStr(pers) & "C4,4,TRUE)"
  
    Range("H2:J2").Select
    Selection.AutoFill Destination:=Range("H2:J" & CStr(gjrs)), Type:=xlFillDefault
    
    Range("H2:J" & CStr(gjrs)).Select
    
    Selection.SpecialCells(xlCellTypeFormulas, xlErrors) = "-"
    
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 
    Sheets("�ɶ�����").Select
    
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
    
    Sheets("�ɼ���ɽ���").Select
    Range("K2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-10],�ɶ�����!R2C1:R" & CStr(gdrs) & "C2,2,TRUE)"
    
    
    Range("K2").Select
    Selection.AutoFill Destination:=Range("K2:K" & CStr(gjrs)), Type:=xlFillDefault
   
    Range("K2:K" & CStr(gjrs)).Select
    '���޷��鵽��ֵ�á�-��������������#N/A���ں���ȡ���ֵʱ����
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
    Sheets("�ɼ���ɽ���").Copy
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
    Sheets("�ɶ�����").Select
    Sheets("�ɶ�����").Copy Before:=Workbooks(nwb).Sheets(1)
    
    Workbooks(owb).Activate
    Sheets("����ɱ��䶯").Select
    Sheets("����ɱ��䶯").Copy Before:=Workbooks(nwb).Sheets(1)
    
    Workbooks(owb).Activate
    Sheets("��ӯ�����о���").Select
    Sheets("��ӯ�����о���").Copy Before:=Workbooks(nwb).Sheets(1)
    
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "�ɶ�������ɼ�����ͼ"
    
    gj_gdhs
    gj_vol
    gj_gb
    gj_pe
    gj_pb
    gj_hjsz
    
    ActiveWorkbook.Save

    
End Sub
Sub gj_vol()

    Sheets("�ɼ���ɽ���").Select
    
    gjrs = Range("A1").CurrentRegion.Rows.Count
    gjcs = Range("A1").CurrentRegion.Columns.Count
   
    Sheets("�ɶ�������ɼ�����ͼ").Select
    ActiveSheet.Shapes.AddChart(xlLine, 40, 360, 900, 350).Select  '���һ����ͼ
    
    With ActiveChart
        .SetSourceData Source:=Sheets("�ɼ���ɽ���").Range("C2:C" & CStr(gjrs))
        .SeriesCollection(1).XValues = "='�ɼ���ɽ���'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(1).Name = "�ɼ�(ǰ��Ȩ)"
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Values = "='�ɼ���ɽ���'!$B$2:$B$" & CStr(gjrs)
        .SeriesCollection(2).XValues = "='�ɼ���ɽ���'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(2).Name = "�ɽ���"
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
            
            .MinimumScale = #1/1/2012#          '41275 2013��01��01��
            .MaximumScale = #12/31/2017#        '43100
        End With
        
        .HasTitle = True
        .ChartTitle.Text = gpdm1 & gpmc1 & "�ɽ�����ɼ�����"
        
        With .Axes(xlValue, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 5
            .MajorGridlines.Border.LineStyle = xlDash
            .HasTitle = True
            .AxisTitle.Text = "�ɽ���"
        End With
        
        With .Axes(xlValue, xlSecondary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 12
            .MajorGridlines.Border.LineStyle = xlDash
            .HasTitle = True
            .AxisTitle.Text = "�ɼ�"
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

    Sheets("�ɼ���ɽ���").Select
    
    gjrs = Range("A1").CurrentRegion.Rows.Count
    gjcs = Range("A1").CurrentRegion.Columns.Count
   
    If gjrs - 1000 < 0 Then
        qsh = 2
    Else
        qsh = gjrs - 1000
    End If
    
    rng = Range("K" & CStr(qsh) & ":K" & CStr(gjrs))
    gdhsmax = WorksheetFunction.Max(rng)
   
    Sheets("�ɶ�������ɼ�����ͼ").Select
    ActiveSheet.Shapes.AddChart(xlLine, 40, 10, 900, 350).Select   '���һ����ͼ
    
    With ActiveChart
        .SetSourceData Source:=Sheets("�ɼ���ɽ���").Range("C2:C" & CStr(gjrs))
        .SeriesCollection(1).XValues = "='�ɼ���ɽ���'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(1).Name = "�ɼ�(ǰ��Ȩ)"
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Values = "='�ɼ���ɽ���'!$K$2:$K$" & CStr(gjrs)
        .SeriesCollection(2).XValues = "='�ɼ���ɽ���'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(2).Name = "�ɶ�����"
        
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
            
            .MinimumScale = #1/1/2012#          '41275 2013��01��01��
            .MaximumScale = #12/31/2017#        '43100
        End With
        
        .HasTitle = True
        .ChartTitle.Text = gpdm1 & gpmc1 & "�ɶ�������ɼ�����"
        
        With .Axes(xlValue, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 5
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "�ɶ�����"
            .MinimumScale = 0
            .MaximumScale = WorksheetFunction.RoundUp(gdhsmax / 1000, 0) * 1000
        End With
        
        With .Axes(xlValue, xlSecondary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 12
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "�ɼ�"
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
'������ֵ
    Sheets("�ɼ���ɽ���").Select
    
    gjrs = Range("A1").CurrentRegion.Rows.Count
    gjcs = Range("A1").CurrentRegion.Columns.Count
   
    If gjrs - 1000 < 0 Then
        qsh = 2
    Else
        qsh = gjrs - 1000
    End If
    
    rng = Range("L" & CStr(qsh) & ":L" & CStr(gjrs))
    hjszmax = WorksheetFunction.Max(rng)
   
    Sheets("�ɶ�������ɼ�����ͼ").Select
    ActiveSheet.Shapes.AddChart(xlLine, 40, 1760, 900, 350).Select   '���һ����ͼ
    
    With ActiveChart
        .SetSourceData Source:=Sheets("�ɼ���ɽ���").Range("C2:C" & CStr(gjrs))
        .SeriesCollection(1).XValues = "='�ɼ���ɽ���'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(1).Name = "�ɼ�(ǰ��Ȩ)"
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Values = "='�ɼ���ɽ���'!$L$2:$L$" & CStr(gjrs)
        .SeriesCollection(2).XValues = "='�ɼ���ɽ���'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(2).Name = "������ֵ(��Ԫ)"
        
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
           
            .MinimumScale = #1/1/2012#          '41275 2013��01��01��
            .MaximumScale = #12/31/2017#        '43100
        End With
        
        .HasTitle = True
        .ChartTitle.Text = gpdm1 & gpmc1 & "������ֵ��ɼ�����"
        
        With .Axes(xlValue, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 5
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "������ֵ"
            .MinimumScale = 0
            .MaximumScale = WorksheetFunction.RoundUp(hjszmax / 10, 0) * 10
        End With
        
        With .Axes(xlValue, xlSecondary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 12
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "�ɼ�"
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

    Sheets("�ɼ���ɽ���").Select
    
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
   
    Sheets("�ɶ�������ɼ�����ͼ").Select
    ActiveSheet.Shapes.AddChart(xlLine, 40, 1060, 900, 350).Select   '���һ����ͼ
    
    With ActiveChart
        .SetSourceData Source:=Sheets("�ɼ���ɽ���").Range("H2:H" & CStr(gjrs))
        .SeriesCollection(1).XValues = "='�ɼ���ɽ���'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(1).Name = "��̬��ӯ��"
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Values = "='�ɼ���ɽ���'!$I$2:$I$" & CStr(gjrs)
        .SeriesCollection(2).XValues = "='�ɼ���ɽ���'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(2).Name = "������ӯ��"
        
        .SeriesCollection.NewSeries
        .SeriesCollection(3).Values = "='�ɼ���ɽ���'!$C$2:$C$" & CStr(gjrs)
        .SeriesCollection(3).XValues = "='�ɼ���ɽ���'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(3).Name = "�ɼ�(ǰ��Ȩ)"
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
            
            .MinimumScale = #1/1/2012#          '41275 2013��01��01��
            .MaximumScale = #12/31/2017#        '43100
        End With
        
        .HasTitle = True
        .ChartTitle.Text = gpdm1 & gpmc1 & "��ӯ����ɼ�����"
        
        With .Axes(xlValue, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 5
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "��ӯ�ʱ���"
            .MinimumScale = 0
            .MaximumScale = WorksheetFunction.RoundUp(pemax / 10, 0) * 10
        End With
        
        With .Axes(xlValue, xlSecondary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 12
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "�ɼ�"
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

    Sheets("�ɼ���ɽ���").Select
    
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
   
    Sheets("�ɶ�������ɼ�����ͼ").Select
    ActiveSheet.Shapes.AddChart(xlLine, 40, 1410, 900, 350).Select   '���һ����ͼ
    
    With ActiveChart
        .SetSourceData Source:=Sheets("�ɼ���ɽ���").Range("C2:C" & CStr(gjrs))
        .SeriesCollection(1).XValues = "='�ɼ���ɽ���'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(1).Name = "�ɼ�(ǰ��Ȩ)"
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Values = "='�ɼ���ɽ���'!$J$2:$J$" & CStr(gjrs)
        .SeriesCollection(2).XValues = "='�ɼ���ɽ���'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(2).Name = "�о���"
        
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
            
            .MinimumScale = #1/1/2012#          '41275 2013��01��01��
            .MaximumScale = #12/31/2017#        '43100
        End With
        
        .HasTitle = True
        .ChartTitle.Text = gpdm1 & gpmc1 & "�о�����ɼ�����"
        
        With .Axes(xlValue, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 5
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "�о���"
            .MinimumScale = 0
            .MaximumScale = WorksheetFunction.RoundUp(pbmax, 2)
        End With
        
        With .Axes(xlValue, xlSecondary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 12
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "�ɼ�"
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

    Sheets("�ɼ���ɽ���").Select
    
    gjrs = Range("A1").CurrentRegion.Rows.Count
    gjcs = Range("A1").CurrentRegion.Columns.Count
   
    Sheets("�ɶ�������ɼ�����ͼ").Select
    ActiveSheet.Shapes.AddChart(xlColumnClustered, 40, 710, 900, 350).Select   '���һ��״ͼ
    
    With ActiveChart
        .SetSourceData Source:=Sheets("�ɼ���ɽ���").Range("E2:E" & CStr(gjrs))
        .SeriesCollection(1).XValues = "='�ɼ���ɽ���'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(1).Name = "�ܹɱ�"
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Values = "='�ɼ���ɽ���'!$G$2:$G$" & CStr(gjrs)
        .SeriesCollection(2).XValues = "='�ɼ���ɽ���'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(2).Name = "ʵ����ͨ�ɱ�"
        
        .SeriesCollection.NewSeries
        .SeriesCollection(3).Values = "='�ɼ���ɽ���'!$C$2:$C$" & CStr(gjrs)
        .SeriesCollection(3).XValues = "='�ɼ���ɽ���'!$A$2:$A$" & CStr(gjrs)
        .SeriesCollection(3).Name = "�ɼ�(ǰ��Ȩ)"
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
            
            .MinimumScale = #1/1/2012#          '41275 2013��01��01��
            .MaximumScale = #12/31/2017#        '43100
        End With
        
        .HasTitle = True
        .ChartTitle.Text = gpdm1 & gpmc1 & "�ܹɱ���ʵ����ͨ�ɱ���ɼ�����"
        
        With .Axes(xlValue, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 5
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "�ɱ�"
        End With
        
        With .Axes(xlValue, xlSecondary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.ColorIndex = 12
            .MajorGridlines.Border.LineStyle = xlDot
            .MajorGridlines.Border.Weight = xlHairline
            .HasTitle = True
            .AxisTitle.Text = "�ɼ�"
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


'���ɹ�Ʊ�����ֵ�
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
    
        Open fn For Binary As #1 '���ı��ļ�
        
        Get #1, , Header
        Do
            Get #1, , gpdm
            dm = ByteToStr(gpdm) & ".S" & UCase(Mid("hz", n, 1))
            
            Get #1, , unknow1
            Get #1, , gpmc
            For i = 1 To 8
                If gpmc(i) = 0 Then
                    gpmc(i) = 32    'x00����x20�ո����
                End If
            Next
            mc = ByteToStr(gpmc)
            Get #1, , unknow2
            If (n = 1 And Left(dm, 1) = "6") Or (n = 2 And (Left(dm, 1) = "0" Or Left(dm, 2) = "30")) Then
            
                gpmc_dic.Add dm, Replace(Replace(mc, " ", ""), "*", "")
            
            End If
             
        Loop Until EOF(1)
        
        Close #1 '�ر��ļ�
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
