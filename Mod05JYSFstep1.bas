Attribute VB_Name = "Mod05JYSFstep1"
'Name:JYSAstep1()  �����㷨step1-���ݹ淶��ֵD():��һ���ɷֵ�ta,tb(ȫ�ֱ���)->����V(1) V(2) V(3) ͶӰָ��tmin tmax
'���:[D().x,D().y]---���ݹ淶��ֵ (�±������������Ϣ)
'     ��һ���ɷ�=(ta*t,tb*t)��ta,tb(ȫ�ֱ���)
'����:V(1) V(2) V(3) (�±������������Ϣ),tmax ,tmin (ͶӰָ����С�����ֵ)

Public Sub JYSAstep1()
    Dim i As Integer
    Dim ET As Double
    Dim V1X As xy
    Dim t() As Double, P() As xy
    '
    '(1.1) �����ݼ��ڵ�һ���ɷ��ϵ�ͶӰP()
    ReDim t(1 To UBound(D)): ReDim P(1 To UBound(D))
    ET = 0
    For i = 1 To Nd
       'ͶӰָ��
       t(i) = D(i).x * ta + D(i).y * tb
       ET = ET + t(i)
       'ͶӰ
       P(i).x = t(i) * ta
       P(i).y = t(i) * tb
    Next i
    '�������Сt
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
