Attribute VB_Name = "ģ��1"
Option Base 1

Sub xgxx(xgzd As Dictionary)
    Dim i As Integer, j As Integer
    Dim m As Integer, n As Integer, l As Integer
    Dim xgdm, data()
    Dim xxstr As String
    Dim rng As Range
    
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "�¹���Ϣ"
    Cells(1, 1) = "��Ʊ����"
    Cells(1, 1) = "�깺���"
    Cells(1, 1) = "��ǩ����"
    Cells(1, 1) = "���м�"
    Cells(1, 1) = "��������"
    
    xgdm = xgzd.Keys
    m = LBound(xgdm)
    n = UBound(xgdm)
    
    ReDim data(1 To n - m + 1, 1 To 5)
    
    j = 1
    For i = m To n
        data(j, 1) = xgdm(i)
        xxstr = xgzd(xgdm(i))
        l = InStr(xxstr, "|")
        data(j, 2) = Left(xxstr, l - 1)
        xxstr = Mid(xxstr, l + 1)
        l = InStr(xxstr, "|")
        data(j, 3) = Left(xxstr, l - 1)
        xxstr = Mid(xxstr, l + 1)
        l = InStr(xxstr, "|")
        data(j, 4) = Left(xxstr, l - 1)
        xxstr = Mid(xxstr, l + 1)
        data(j, 5) = xxstr
        j = j + 2
    Next

    Set rng = Range(Cells(2, 1), Cells(n - m + 2, 5))
    rng.Value = data

End Sub

Sub xx()
    Dim xgzd
    Set xgzd = CreateObject(Scripting.Dictionary)
    xgzd("SH600201") = "12345|1000|12.345|20150506"
    xgzd("SH600205") = "22345|1000|22.345|20150509"
    xgxx (zgzd)
End Sub
