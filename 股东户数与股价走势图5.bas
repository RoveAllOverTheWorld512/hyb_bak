Attribute VB_Name = "�ɶ�������ɼ�����ͼ"
Sub gdhs_gj1(gpdm1 As String)
'
' �ɶ�������ɼ�����ͼ
'
'
'    gpdm1 = "300340"
    
    Set dic = gpmc_dic
   
    gpdm2 = gpdm1 & IIf(Left(gpdm1, 1) = "6", ".SH", ".SZ")
    gpmc1 = Trim(dic(gpdm2))
    
    gdhsfn = "D:\��˾�о�\" & gpmc1 & "\" & gpdm1 & gpmc1 & "�ɶ�����.xlsx"
    gjfn = "D:\��˾�о�\" & gpmc1 & "\" & gpdm1 & gpmc1 & "�ɼ���ɽ���.xlsx"
    
    Workbooks.Open Filename:=gjfn
    gjwb = ActiveWorkbook.Name
    Sheets("�ɼ���ɽ���").Select
    Range("A1").Select
    All_Str2Date
    
    gjrs = Range("A1").CurrentRegion.Rows.Count
    gjcs = Range("A1").CurrentRegion.Columns.Count
   
    
    Workbooks.Open Filename:=gdhsfn
    gdhswb = ActiveWorkbook.Name
    Sheets("�ɶ�����").Select
    
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "�ɼ�(ǰ��Ȩ)"
    All_Str2Date
    Cells.Select
    Cells.EntireColumn.AutoFit
    gdrs = Range("A1").CurrentRegion.Rows.Count
    gdcs = Range("A1").CurrentRegion.Columns.Count
    
    Range("P2").Select
    
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-13],[" & gjwb & "]�ɼ���ɽ���!R2C1:R" & CStr(gjrs) & "C" & CStr(gjcs) & ",11,TRUE)"
        
    Range("P2").Select
    Selection.AutoFill Destination:=Range("P2:P" & CStr(gdrs)), Type:=xlFillDefault
    
    Sheets.Add After:=Sheets(Sheets.Count)
    
    ActiveSheet.Shapes.AddChart(xlLine, 40, 0, 900, 260).Select   '���һ����ͼ
    
    ActiveChart.SetSourceData Source:=Sheets("�ɶ�����").Range("E2:E" & CStr(gdrs))
    ActiveChart.SeriesCollection(1).XValues = "='�ɶ�����'!$C$2:$C$" & CStr(gdrs)
    
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Values = "='�ɶ�����'!$P$2:$P$" & CStr(gdrs)
    ActiveChart.SeriesCollection(2).XValues = "='�ɶ�����'!$C$2:$C$" & CStr(gdrs)
    
    ActiveChart.SeriesCollection(2).Name = "�ɼ�(ǰ��Ȩ)"
    ActiveChart.SeriesCollection(1).Name = "�ɶ�����"
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).AxisGroup = 2
    ActiveChart.Legend.Select
    Selection.Position = xlBottom
    
    Windows(gjwb).Activate
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    Windows(gdhswb).Activate
    
End Sub

'���ֽ�ת�����ַ���
Public Function ByteToStr(B() As Byte) As String 'Byte����ת�ַ���
    Dim i, Tmp As String
    For Each i In B 'ö���������鸳ֵ��I
        If i > 127 Then '���Ϊ���ֱ���.(����127Ϊ����,ռ�����ֽ�)
            If Tmp <> "" Then '�����ʱ������Ϊ��(Ϊ��Ϊ��һ�ֽ�)
                ByteToStr = ByteToStr & Chr(Tmp * 256 + i) '�ϲ������ֽ�,ת��Ϊ����.�ۼ�����
                Tmp = "" '�����ʱ����
            Else
                Tmp = i '������ʱ����
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

Sub gdhs_gj2(gpdm1 As String)
'
' �ɶ�������ɼ�����ͼ
'
'
'    gpdm1 = "600201"
    
    '���ɹ�Ʊ�����ֵ�
    Set dic = gpmc_dic
   
    gpdm2 = gpdm1 & IIf(Left(gpdm1, 1) = "6", ".SH", ".SZ")
    gpmc1 = Trim(dic(gpdm2))
    
    gdhsfn = "D:\��˾�о�\" & gpmc1 & "\" & gpdm1 & gpmc1 & "�ɶ�����.xlsx"
    gjfn = "D:\��˾�о�\" & gpmc1 & "\" & gpdm1 & gpmc1 & "�ɼ���ɽ���.xlsx"
    
    Workbooks.Open Filename:=gjfn
    gjwb = ActiveWorkbook.Name
    Sheets("�ɼ���ɽ���").Select
    Range("A1").Select
    All_Str2Date
    
    Range("N1").Value = "�ɶ�����"
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
    
    
    Workbooks.Open Filename:=gdhsfn
    gdhswb = ActiveWorkbook.Name
    Sheets("�ɶ�����").Select
    All_Str2Date
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("C2"), _
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
    
    Windows(gjwb).Activate
    Sheets("�ɼ���ɽ���").Select
    Range("N2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-13],[" & gdhswb & "]�ɶ�����!R2C3:R" & CStr(gdrs) & "C5,3,TRUE)"
    
    
    Range("N2").Select
    Selection.AutoFill Destination:=Range("N2:N" & CStr(gjrs)), Type:=xlFillDefault
   
    Range("N2:N" & CStr(gjrs)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Windows(gdhswb).Activate
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    Windows(gjwb).Activate
   
    Sheets.Add After:=Sheets(Sheets.Count)
    
    ActiveSheet.Shapes.AddChart(xlLine, 40, 10, 900, 350).Select   '���һ����ͼ

    
    ActiveChart.SetSourceData Source:=Sheets("�ɼ���ɽ���").Range("K2:K" & CStr(gjrs))
    ActiveChart.SeriesCollection(1).XValues = "='�ɼ���ɽ���'!$A$2:$A$" & CStr(gjrs)
    ActiveChart.SeriesCollection(1).Name = "�ɼ�(ǰ��Ȩ)"
    
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Values = "='�ɼ���ɽ���'!$N$2:$N$" & CStr(gjrs)
    ActiveChart.SeriesCollection(2).XValues = "='�ɼ���ɽ���'!$A$2:$A$" & CStr(gjrs)
    ActiveChart.SeriesCollection(2).Name = "�ɶ�����"
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).AxisGroup = 2
    ActiveChart.Legend.Position = xlBottom
    
    ActiveChart.Axes(xlCategory).CategoryType = xlTimeScale
    ActiveChart.Axes(xlCategory).BaseUnit = xlDays
    ActiveChart.Axes(xlCategory).MajorUnitScale = xlMonths
    ActiveChart.Axes(xlCategory).MajorUnit = 1
    
    ActiveChart.Axes(xlCategory).MinimumScale = #1/1/2013#          '41275 2013��01��01��
    ActiveChart.Axes(xlCategory).MaximumScale = #12/31/2017#        '43100
    
    ActiveChart.Axes(xlValue).HasTitle = True
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleVertical)
    ActiveChart.SetElement (msoElementSecondaryValueAxisTitleVertical)
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "�ɶ�����"
'    ActiveChart.SetElement (msoElementSecondaryValueAxisTitleRotated)
    ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Text = "�ɼ�(ǰ��Ȩ)"
    
    ActiveChart.HasTitle = True
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    ActiveChart.ChartTitle.Text = gpdm1 & gpmc1 & "�ɶ�������ɼ����ƹ�ϵ"
    
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
Sub gdhs_gj()
'
' �ɶ�������ɼ�����ͼ
'
'
'    gpdm1 = "600201"
    
    '���ɹ�Ʊ�����ֵ�
    Set dic = gpmc_dic
   
    gpdm1 = Sheets("_xlwings.conf").Range("B8").Value
    gpmc1 = Sheets("_xlwings.conf").Range("B9").Value
    
    Sheets("�ɼ���ɽ���").Select
    Range("A1").Select
    
    Range("D1").Value = "�ɶ�����"
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
    Range("D2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-3],�ɶ�����!R2C1:R" & CStr(gdrs) & "C5,2,TRUE)"
    
    
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D" & CStr(gjrs)), Type:=xlFillDefault
   
    Range("D2:D" & CStr(gjrs)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
   
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "�ɶ�������ɼ�����ͼ"
    
    ActiveSheet.Shapes.AddChart(xlLine, 40, 10, 900, 350).Select   '���һ����ͼ

    
    ActiveChart.SetSourceData Source:=Sheets("�ɼ���ɽ���").Range("C2:C" & CStr(gjrs))
    ActiveChart.SeriesCollection(1).XValues = "='�ɼ���ɽ���'!$A$2:$A$" & CStr(gjrs)
    ActiveChart.SeriesCollection(1).Name = "�ɼ�(ǰ��Ȩ)"
    
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Values = "='�ɼ���ɽ���'!$D$2:$D$" & CStr(gjrs)
    ActiveChart.SeriesCollection(2).XValues = "='�ɼ���ɽ���'!$A$2:$A$" & CStr(gjrs)
    ActiveChart.SeriesCollection(2).Name = "�ɶ�����"
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).AxisGroup = 2
    ActiveChart.Legend.Position = xlBottom
    
    ActiveChart.Axes(xlCategory).CategoryType = xlTimeScale
    ActiveChart.Axes(xlCategory).BaseUnit = xlDays
    ActiveChart.Axes(xlCategory).MajorUnitScale = xlMonths
    ActiveChart.Axes(xlCategory).MajorUnit = 1
    
    ActiveChart.Axes(xlCategory).MinimumScale = #1/1/2013#          '41275 2013��01��01��
    ActiveChart.Axes(xlCategory).MaximumScale = #12/31/2017#        '43100
    
    ActiveChart.Axes(xlValue).HasTitle = True
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleVertical)
    ActiveChart.SetElement (msoElementSecondaryValueAxisTitleVertical)
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "�ɶ�����"
'    ActiveChart.SetElement (msoElementSecondaryValueAxisTitleRotated)
    ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Text = "�ɼ�(ǰ��Ȩ)"
    
    ActiveChart.HasTitle = True
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    ActiveChart.ChartTitle.Text = gpdm1 & gpmc1 & "�ɶ�������ɼ����ƹ�ϵ"
    
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

'���ɹ�Ʊ�����ֵ�
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


