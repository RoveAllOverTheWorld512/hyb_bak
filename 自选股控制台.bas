Attribute VB_Name = "��ѡ�ɿ���̨"
Sub zxg_list()

    Dim Header(1 To 50) As Byte
    Dim gpdm(1 To 6) As Byte
    Dim unknow1(1 To 17) As Byte
    Dim gpmc(1 To 8) As Byte
    Dim unknow2(1 To 283) As Byte
    
    Dim dm As String, mc As String
    
    Dim arr, string1 As String
    Dim i As Integer
    
    Application.ScreenUpdating = False  '����ˢ�¹رգ�������������ٶ�
    Application.DisplayAlerts = False '�رձ���
    
    If Workbooks.Count = 0 Then
        Workbooks.add
        cursht = "Sheet2"
    Else
        cursht = ActiveSheet.Name
        Sheets.add After:=Sheets(Sheets.Count)
    
    End If
'    MsgBox ActiveWorkbook.Name

'Ϊ�ù���������¼�����
    shtn = ActiveSheet.Name
    
    ShtCodeName = Application.VBE.MainWindow.Caption    '���������Ҫ����ˢ��VBE
    ShtCodeName = Sheets(shtn).CodeName
    
    
    With ActiveWorkbook.VBProject.VBComponents(ShtCodeName).CodeModule
    
        .InsertLines 1, "Private Sub Worksheet_Activate()"
        .InsertLines 2, "    For Each vbCmp In ThisWorkbook.VBProject.VBComponents"
        .InsertLines 3, "        fname = vbCmp.Name"
        .InsertLines 4, "        If fname = " & Chr(34) & "�ɶ�������ɼ�����ͼ" & Chr(34) & " Then"
'        .InsertLines 5, "            With Application.VBE.ActiveVBProject.VBComponents"
'        .InsertLines 6, "                .Remove .Item(fname)"
'        .InsertLines 7, "            End With"
        .InsertLines 6, "             Exit Sub"
        .InsertLines 8, "        End If"
        .InsertLines 9, "    Next vbCmp"
        .InsertLines 10, "    ShtCodeName = Application.VBE.MainWindow.Caption    '���������Ҫ����ˢ��VBE"
        .InsertLines 11, "    Application.VBE.ActiveVBProject.VBComponents.Import " & Chr(34) & "d:\hyb\�ɶ�������ɼ�����ͼ.bas" & Chr(34)
        .InsertLines 12, "End Sub"
    
        .InsertLines 13, "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)"
        .InsertLines 14, "    If Target.EntireColumn.Column = 1 Then"
        .InsertLines 15, "         Sheets(""_xlwings.conf"").Range(""B8"").Value = Target.Value"
        .InsertLines 16, "         Sheets(""_xlwings.conf"").Range(""B9"").Value = Cells(Target.EntireRow.Row, 2).Value"
        .InsertLines 17, "         get_data"
        .InsertLines 18, "         gdhs_gj"
        .InsertLines 19, "    End If"
        .InsertLines 20, "End Sub"
    
    
    End With


'���ɹ�Ʊ�����ֵ�
    Set dic = gpmc_dic

'��ȡ��ѡ������
    
    Open "C:\new_hxzq_hc\T0002\blocknew\zxg.blk" For Input As #1 '���ı��ļ�
    
    string1 = Input(LOF(1), 1)
    
    Close #1 '�ر��ļ�
    
    arr = Split(string1, Chr(13) + Chr(10))
   
    Cells(1, 1).Value = "��Ʊ����"
    Cells(1, 2).Value = "��Ʊ����"
    Cells(1, 3).Value = "�ɶ�����"
    Cells(1, 4).Value = "�б�(�����Ƹ���)"
    Cells(1, 5).Value = "�������"
    Cells(1, 6).Value = "ͬ��˳����"
    Cells(1, 7).Value = "�߹ֹܳ�"
    Cells(1, 8).Value = "��˾����"
    Cells(1, 9).Value = "�ȵ�����"
    Cells(1, 10).Value = "�����ֹ�"
    Cells(1, 11).Value = "�б�(ͬ��˳)"
    Cells(1, 12).Value = "���۹ɽ��"
    Cells(1, 13).Value = "��ɶ��ֹɱ䶯"
    Cells(1, 14).Value = "��ֵ����"
    Cells(1, 15).Value = "�������"
    Cells(1, 16).Value = "��ֵ������ͬ��˳��"
    
    j = 2
    
    For i = 0 To UBound(arr)
        
        If Len(arr(i)) = 7 Then
        
            dm = Mid(arr(i), 2)
            link1 = "http://data.eastmoney.com/gdhs/detail/" & dm & ".html"
            link2 = "http://data.eastmoney.com/report/" & dm & ".html"
            link3 = "http://news.stockstar.com/info/dstock.aspx?code=" & dm
            link4 = "http://stockpage.10jqka.com.cn/" & dm & "/"
            link5 = "http://stockpage.10jqka.com.cn/" & dm & "/event/#manager"
            link6 = "http://stockpage.10jqka.com.cn/" & dm & "/event/#remind"
            link7 = "http://search.10jqka.com.cn/search?&w=" & dm
            link8 = "http://stockpage.10jqka.com.cn/" & dm & "/position/#organhold"
            link9 = "http://stockpage.10jqka.com.cn/" & dm & "/worth/#stockreport"
'            link10 = "http://stockpage.10jqka.com.cn/" & dm & "/holder/#liftban"
            link10 = "http://data.eastmoney.com/dxf/q/" & dm & ".html"
            link11 = "http://stockpage.10jqka.com.cn/" & dm & "/event/#holder"
            link12 = "http://web-f10.gaotime.com/stock/" & dm & "/jzfx.html"
            link13 = "http://data.10jqka.com.cn/market/rzrqgg/code/" & dm & "/"
            link14 = "http://stockpage.10jqka.com.cn/" & dm & "/worth/"
            
            If Left(dm, 1) = "6" Then
                dm = dm & ".SH"
            Else
                dm = dm & ".SZ"
            End If
                
            Cells(j, 1).Value = dm
            Cells(j, 2).Value = dic(dm)
            
            ActiveSheet.Hyperlinks.add Anchor:=Cells(j, 3), Address:=link1, TextToDisplay:="�����Ƹ���"
            ActiveSheet.Hyperlinks.add Anchor:=Cells(j, 4), Address:=link2, TextToDisplay:="�����Ƹ���"
            ActiveSheet.Hyperlinks.add Anchor:=Cells(j, 5), Address:=link3, TextToDisplay:="֤ȯ֮����"
            ActiveSheet.Hyperlinks.add Anchor:=Cells(j, 6), Address:=link4, TextToDisplay:="ͬ��˳��"
            ActiveSheet.Hyperlinks.add Anchor:=Cells(j, 7), Address:=link5, TextToDisplay:="�߹ֹܳ�"
            ActiveSheet.Hyperlinks.add Anchor:=Cells(j, 8), Address:=link6, TextToDisplay:="��˾����"
            ActiveSheet.Hyperlinks.add Anchor:=Cells(j, 9), Address:=link7, TextToDisplay:="�ʲ���"
            ActiveSheet.Hyperlinks.add Anchor:=Cells(j, 10), Address:=link8, TextToDisplay:="�����ֹ�"
            ActiveSheet.Hyperlinks.add Anchor:=Cells(j, 11), Address:=link9, TextToDisplay:="ͬ��˳"
            ActiveSheet.Hyperlinks.add Anchor:=Cells(j, 12), Address:=link10, TextToDisplay:="���۹ɽ��"
            ActiveSheet.Hyperlinks.add Anchor:=Cells(j, 13), Address:=link11, TextToDisplay:="��ɶ��ֹɱ䶯"
            ActiveSheet.Hyperlinks.add Anchor:=Cells(j, 14), Address:=link12, TextToDisplay:="��ֵ����"
            ActiveSheet.Hyperlinks.add Anchor:=Cells(j, 15), Address:=link13, TextToDisplay:="�������"
            ActiveSheet.Hyperlinks.add Anchor:=Cells(j, 16), Address:=link14, TextToDisplay:="��ֵ������ͬ��˳��"
            
            j = j + 1
            
        End If
    Next
    
'������ʽ
    Cells.Select
    Cells.EntireColumn.AutoFit
    Rows(1).Select
    Selection.Font.Bold = True
    Range("A1").CurrentRegion.Select
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
'��������
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
'��������
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    r = Range("A1").CurrentRegion.Rows.Count
    
    Range("A1:B" & CStr(r)).Select
    Selection.Copy

'������ѡ��.txt
    Workbooks.add
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveWorkbook.SaveAs Filename:="D:\hyb\��ѡ��.txt", FileFormat:=xlUnicodeText _
        , CreateBackup:=False
    ActiveWindow.Close
    
    Application.ScreenUpdating = True
    
    Sheets(cursht).Activate
    Range("A1").Select
    
    MsgBox "��鿴������" & shtn & "��"
    
    Sheets(shtn).Activate   '������������ģ��
    Range("A1").Select
    
    fn = ActiveWorkbook.Name
    If Right(fn, 4) <> ".xlsm" Then
        i = InStr(fn, ".")
        If i = 0 Then
            fn = fn & ".xlsm"
        Else
            fn = Mid(fn, 1, i - 1) & ".xlsm"
        End If
    End If
    
    ActiveWorkbook.SaveAs Filename:=fn, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
'    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
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

'���ɹ�Ʊ�����ֵ�
Public Function gpmc_dic()
    Dim Header(1 To 50) As Byte
    Dim gpdm(1 To 6) As Byte
    Dim unknow1(1 To 17) As Byte
    Dim gpmc(1 To 8) As Byte
    Dim unknow2(1 To 254) As Byte
    Dim gppy(1 To 8) As Byte
    Dim unknow3(1 To 21) As Byte
    
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
            Get #1, , gppy
            For i = 1 To 8
                If gppy(i) = 0 Then
                    gppy(i) = 32    'x00����x20�ո����
                End If
            Next
            Get #1, , unknow3

            If (n = 1 And Left(dm, 1) = "6") Or (n = 2 And (Left(dm, 1) = "0" Or Left(dm, 2) = "30")) Then
            
                gpmc_dic.add dm, Replace(Replace(mc, " ", ""), "*", "") & "|" & py
            
            End If
             
        Loop Until EOF(1)
        
        Close #1 '�ر��ļ�
    Next
    
End Function

