Attribute VB_Name = "模块1"

Sub GSMn()
ThisWorkbook.Sheets.Add Before:=Worksheets(1)
ActiveSheet.Name = "sheet2"
Worksheets("sheet2").Activate
Range("a1") = "BSC"
Range("B1") = "LAC号"
Range("c1") = "站点号"
Range("d1") = "站点名"
Range("e1") = "小区号"
Range("f1") = "小区名"
Range("g1") = "经度"
Range("h1") = "纬度"
Range("i1") = "覆盖类型"
Range("j1") = "主频点"
Range("k1") = "TCH"
Range("l1") = "方向角"
Range("m1") = "半径_米"
Range("n1") = "波瓣_度"
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

