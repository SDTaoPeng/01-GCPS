Attribute VB_Name = "Mod04PCARTB"
'Name:PCARTB()---第一主成分=(ta*t,tb*t)的ta,tb(全局变量)
'Function:(1)RTB算法计算第一主成分(过0点)(假设E(X)=0)
'入口:(全局变量) 数据点个数Nd,数据点集合[D().x,D().y]
'出口参数：第一主成分=(ta*t,tb*t)的ta,tb(全局变量)
Public Sub PCARTB(ByRef DataPoint() As xy)   'PCA-RTB Rotations and Translations of Blocks
   Dim i As Integer, j As Integer
   Dim tx1 As Double, tx2 As Double, tx As Double
   
   'PCA-RTB
   i = LBound(D): j = UBound(D)
   ReDim t(i To j)         '投影指标(Projection Index)
   '----------------------------------------------------
   'S=tu->             'Step 0
   ta = 0.1: tb = Sqr(1 - ta * ta)           '初始值设定
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
        ta = tx1 / tx: tb = tx2 / tx      '给全局变量ta,tb赋值
        '
        j = j + 1
   Loop Until j >= 3000
End Sub

