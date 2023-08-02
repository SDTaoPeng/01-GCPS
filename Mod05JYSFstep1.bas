Attribute VB_Name = "Mod05JYSFstep1"
'Name:JYSAstep1()  简易算法step1-数据规范点值D():第一主成分的ta,tb(全局变量)->顶点V(1) V(2) V(3) 投影指标tmin tmax
'入口:[D().x,D().y]---数据规范点值 (下标包含个数的信息)
'     第一主成分=(ta*t,tb*t)的ta,tb(全局变量)
'出口:V(1) V(2) V(3) (下标包含个数的信息),tmax ,tmin (投影指标最小、最大值)

Public Sub JYSAstep1()
    Dim i As Integer
    Dim ET As Double
    Dim V1X As xy
    Dim t() As Double, P() As xy
    '
    '(1.1) 求数据集在第一主成分上的投影P()
    ReDim t(1 To UBound(D)): ReDim P(1 To UBound(D))
    ET = 0
    For i = 1 To Nd
       '投影指标
       t(i) = D(i).x * ta + D(i).y * tb
       ET = ET + t(i)
       '投影
       P(i).x = t(i) * ta
       P(i).y = t(i) * tb
    Next i
    '求最大、最小t
    tmax = t(1): tmin = t(1)
    For i = 2 To Nd
        If tmax < t(i) Then tmax = t(i)
        If tmin > t(i) Then tmin = t(i)
    Next i
    V1X.x = (ET / Nd) * ta
    V1X.y = (ET / Nd) * tb
    '
    Kv = 3
    ReDim V(1 To Kv)
    V(1).x = tmin * ta: V(1).y = tmin * tb
    V(2).x = V1X.x: V(2).y = V1X.y
    V(3).x = tmax * ta: V(3).y = tmax * tb
    
End Sub
