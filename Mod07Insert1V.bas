Attribute VB_Name = "Mod07Insert1V"
'Name:Insert1V() 增加一个顶点
'入口:
'出口:
'调用例:

Public Sub Insert1V()
   Dim i As Integer, j As Integer, n As Integer, SegmentNumber As Integer
   Dim n1 As Double
   Dim Vnew As xy
   
   '
   SegmentNumber = UBound(uxy)             '线段个数
   ReDim DataNumberofSegment(1 To SegmentNumber)   '隶属于各线段s()的数据个数
   
   '隶属于各线段的数据个数DataNumberofSegment()
   For i = 1 To SegmentNumber
       DataNumberofSegment(i) = 0    '隶属于各线段s()的数据个数
   Next i
   
   For j = 1 To UBound(D)      '对数据循环
       If DtoVS(j) > 20000 Then i = DtoVS(j) - 20000: DataNumberofSegment(i) = DataNumberofSegment(i) + 1
   Next j
   
   '求最大DataNumberofSegment(i)中的i
   n1 = DataNumberofSegment(1)
   n = 1
   For i = 2 To SegmentNumber
       If DataNumberofSegment(i) > n1 Then n1 = DataNumberofSegment(i): n = i
   Next i
   '
   Vnew.X = (V(n).X + V(n + 1).X) / 2: Vnew.Y = (V(n).Y + V(n + 1).Y) / 2  '新顶点
   
   VnewSerialNumber = n + 1
   'FrmPC.Text1.Text = n1
   ReDim Preserve V(1 To SegmentNumber + 2)  '顶点
   For i = SegmentNumber + 2 To n + 2 Step -1  '后移1
       V(i).X = V(i - 1).X: V(i).Y = V(i - 1).Y
   Next i
   V(n + 1).X = Vnew.X: V(n + 1).Y = Vnew.Y
   Call SegmentExpression(V, tmin)           '求线段的uxy(),tsx()   V(1)与V(2)间是线段uxy(1)
End Sub
