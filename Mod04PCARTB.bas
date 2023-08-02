Attribute VB_Name = "Mod04PCARTB"
'Name:PCARTB()---��һ���ɷ�=(ta*t,tb*t)��ta,tb(ȫ�ֱ���)
'Function:(1)RTB�㷨�����һ���ɷ�(��0��)(����E(X)=0)
'���:(ȫ�ֱ���) ���ݵ����Nd,���ݵ㼯��[D().x,D().y]
'���ڲ�������һ���ɷ�=(ta*t,tb*t)��ta,tb(ȫ�ֱ���)
Public Sub PCARTB(ByRef DataPoint() As xy)   'PCA-RTB Rotations and Translations of Blocks
   Dim i As Integer, j As Integer
   Dim tx1 As Double, tx2 As Double, tx As Double
   
   'PCA-RTB
   i = LBound(D): j = UBound(D)
   ReDim t(i To j)         'ͶӰָ��(Projection Index)
   '----------------------------------------------------
   'S=tu->             'Step 0
   ta = 0.1: tb = Sqr(1 - ta * ta)           '��ʼֵ�趨
   '
   j = 1
   Do
        'step 1
        For i = 1 To UBound(DataPoint)
            t(i) = DataPoint(i).X * ta + DataPoint(i).Y * tb
        Next i
        'step 2
        tx1 = 0: tx2 = 0
        For i = 1 To UBound(DataPoint)
            tx1 = tx1 + t(i) * DataPoint(i).X
            tx2 = tx2 + t(i) * DataPoint(i).Y
        Next i
        tx = Sqr(tx1 * tx1 + tx2 * tx2)
        ta = tx1 / tx: tb = tx2 / tx      '��ȫ�ֱ���ta,tb��ֵ
        '
        j = j + 1
   Loop Until j >= 3000
End Sub

