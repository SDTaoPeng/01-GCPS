Attribute VB_Name = "Mod07Insert1V"
'Name:Insert1V() ����һ������
'���:
'����:
'������:

Public Sub Insert1V()
   Dim i As Integer, j As Integer, n As Integer, SegmentNumber As Integer
   Dim n1 As Double
   Dim Vnew As xy
   
   '
   SegmentNumber = UBound(uxy)             '�߶θ���
   ReDim DataNumberofSegment(1 To SegmentNumber)   '�����ڸ��߶�s()�����ݸ���
   
   '�����ڸ��߶ε����ݸ���DataNumberofSegment()
   For i = 1 To SegmentNumber
       DataNumberofSegment(i) = 0    '�����ڸ��߶�s()�����ݸ���
   Next i
   
   For j = 1 To UBound(D)      '������ѭ��
       If DtoVS(j) > 20000 Then i = DtoVS(j) - 20000: DataNumberofSegment(i) = DataNumberofSegment(i) + 1
   Next j
   
   '�����DataNumberofSegment(i)�е�i
   n1 = DataNumberofSegment(1)
   n = 1
   For i = 2 To SegmentNumber
       If DataNumberofSegment(i) > n1 Then n1 = DataNumberofSegment(i): n = i
   Next i
   '
   Vnew.X = (V(n).X + V(n + 1).X) / 2: Vnew.Y = (V(n).Y + V(n + 1).Y) / 2  '�¶���
   
   VnewSerialNumber = n + 1
   'FrmPC.Text1.Text = n1
   ReDim Preserve V(1 To SegmentNumber + 2)  '����
   For i = SegmentNumber + 2 To n + 2 Step -1  '����1
       V(i).X = V(i - 1).X: V(i).Y = V(i - 1).Y
   Next i
   V(n + 1).X = Vnew.X: V(n + 1).Y = Vnew.Y
   Call SegmentExpression(V, tmin)           '���߶ε�uxy(),tsx()   V(1)��V(2)�����߶�uxy(1)
End Sub
