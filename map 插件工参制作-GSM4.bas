Attribute VB_Name = "ģ��1"

Sub GSMn()
ThisWorkbook.Sheets.Add Before:=Worksheets(1)
ActiveSheet.Name = "sheet2"
Worksheets("sheet2").Activate
Range("a1") = "BSC"
Range("B1") = "LAC��"
Range("c1") = "վ���"
Range("d1") = "վ����"
Range("e1") = "С����"
Range("f1") = "С����"
Range("g1") = "����"
Range("h1") = "γ��"
Range("i1") = "��������"
Range("j1") = "��Ƶ��"
Range("k1") = "TCH"
Range("l1") = "�����"
Range("m1") = "�뾶_��"
Range("n1") = "����_��"
Sheet1.Activate
Range("c2:c8000").Copy Sheets("Sheet2").Range("a2")
Range("H2:H8000").Copy Sheets("Sheet2").Range("b2")
Range("I2:I8000").Copy Sheets("Sheet2").Range("c2")
Range("D2:D8000").Copy Sheets("Sheet2").Range("d2")
Range("I2:I8000").Copy Sheets("Sheet2").Range("e2")
Range("D2:D8000").Copy Sheets("Sheet2").Range("f2")
 Range("K2:K8000").Copy Sheets("Sheet2").Range("g2")
Range("L2:L8000").Copy Sheets("Sheet2").Range("h2")
                                                         
Range("AB2:AB8000").Copy Sheets("Sheet2").Range("J2")
Range("g2:g8000").Copy Sheets("Sheet2").Range("K2")
Range("M2:M8000").Copy Sheets("Sheet2").Range("L2")
     End Sub

