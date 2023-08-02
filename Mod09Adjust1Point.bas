Attribute VB_Name = "Mod09"
'入口参数:m-顶点序号,dv-调整步长
'全局变量:V()-顶点数组 GV()
'Public Pi() As Double              '顶点的角度惩罚
'Public PV() As Double              '顶点的角度惩罚总和
'Public DairTa() As Double          '顶点的距离约束总和
'Public GV() As Double              '顶点的距离约束+角度惩罚 总和
'Public DistanceofDtoVSZ As Double  '数据点到曲线的总距离平方
'Const Namnapp = 0.13
Public Sub Adjust1Point(ByVal m As Integer, ByVal dv As Double)
    Dim Vtemp0 As xy, Vtemp1 As xy, GVtemp As Double, n As Integer
    Dim d1 As Double, d2 As Double
    Dim i As Integer, j As Integer
    Dim DistanceofDtoVSZtemp As Double   '数据点到曲线的总距离平方（上次）
    Dim bz As Boolean
    Dim Xjiaxs As Double, Xjianxs As Double, Yjiaxs As Double, Yjianxs As Double '设4个系数
    Xjiaxs = 1#: Xjianxs = 1#: Yjiaxs = 1#: Yjianxs = 1#                         '防止溢出
    
    If V(m).X >= 0.99 Then Xjiaxs = 0
    If V(m).X <= -0.99 Then Xjianxs = 0
    
    If V(m).Y >= 0.99 Then Yjiaxs = 0
    If V(m).Y <= -0.99 Then Yjianxs = 0
    
'    If (Abs(V(m).y) >= 0.92) Then Yjiaxs = 0
'    If Abs(V(m).x <= 0.08) Then Xjianxs = 0
'    If Abs(V(m).y <= 0.08) Then Yjianxs = 0
  
  
    '保存当前顶点在Vtemp0及Vtemp1中,保存当前GV(m)在GVtemp中
    '保存DistanceofDtoVSZ在DistanceofDtoVSZtemp中
    Vtemp0.X = V(m).X: Vtemp0.Y = V(m).Y: Vtemp1.X = V(m).X: Vtemp1.Y = V(m).Y
    GVtemp = GV(m)
    DistanceofDtoVSZtemp = DistanceofDtoVSZ
    'n = m
    'If m = LBound(V) Then n = m + 1
    'If m = UBound(V) Then n = m - 1
    i = 0
    i = i + 1
    V(m).X = Vtemp0.X: V(m).Y = Vtemp0.Y - Yjianxs * dv
    'Call Adjust1PointSub(m, DistanceofDtoVSZtemp, Vtemp1)      '调整
    Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
    
    i = i + 1
    V(m).X = Vtemp0.X: V(m).Y = Vtemp0.Y + Yjiaxs * dv
     'Call Adjust1PointSub(m, DistanceofDtoVSZtemp, Vtemp1)     '调整
     Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整

    i = i + 1
    V(m).X = Vtemp0.X - Xjianxs * dv: V(m).Y = Vtemp0.Y
    'Call Adjust1PointSub(m, DistanceofDtoVSZtemp, Vtemp1)           '调整
     Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
     
     i = i + 1
    V(m).X = Vtemp0.X + Xjiaxs * dv: V(m).Y = Vtemp0.Y
     'Call Adjust1PointSub(m, DistanceofDtoVSZtemp, Vtemp1)           '调整
    Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
    
    i = i + 1
    V(m).X = Vtemp0.X + Xjiaxs * dv: V(m).Y = Vtemp0.Y + Yjiaxs * dv
     'Call Adjust1PointSub(m, DistanceofDtoVSZtemp, Vtemp1)           '调整
     Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
     
     i = i + 1
    V(m).X = Vtemp0.X + Xjiaxs * dv: V(m).Y = Vtemp0.Y - Yjianxs * dv
     'Call Adjust1PointSub(m, DistanceofDtoVSZtemp, Vtemp1)           '调整
     Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
     
    i = i + 1
    V(m).X = Vtemp0.X - Xjianxs * dv: V(m).Y = Vtemp0.Y + Yjiaxs * dv
    'Call Adjust1PointSub(m, DistanceofDtoVSZtemp, Vtemp1)           '调整
     Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
    '
    i = i + 1
    V(m).X = Vtemp0.X - Xjianxs * dv: V(m).Y = Vtemp0.Y - Yjianxs * dv
    'Call Adjust1PointSub(m, DistanceofDtoVSZtemp, Vtemp1)           '调整
    Call Adjust1PointSub1(i, m, DistanceofDtoVSZtemp, Vtemp1)   '调整
    
    d1 = MoveDirectionDistance(1): j = 1               '求出最小的
    For i = 2 To 8
        If MoveDirectionDistance(i) < d1 Then
           'd1 = MoveDirectionDistance(i): j = i
           d1 = MoveDirectionDistance(i)
           j = i
        End If
    Next i
  
  
    Dim DistanceofDtoVSZtemp1 As Double   '临时变量，上次
    Dim DistanceofDtoVSZtemp2 As Double   '临时变量，上次
    Dim DistanceofDtoVSZtemp3 As Double   '临时变量，上次
    'Dim MoveDirectionDistance1 As Double  '临时变量，当前
    'DistanceofDtoVSZtemp1 = DistanceofDtoVSZtemp  '临时变量，上次
    'MoveDirectionDistance1 = MoveDirectionDistance(j)  '临时变量，当前
    'DistanceofDtoVSZtemp2 = Sqr(DistanceofDtoVSZtemp1 * DistanceofDtoVSZtemp1 - 1)
'    DistanceofDtoVSZtemp3 = Log(1 + Exp(DistanceofDtoVSZtemp1)) / Log(10) / DistanceofDtoVSZtemp1
    'DistanceofDtoVSZtemp1 = DistanceofDtoVSZtemp2 * DistanceofDtoVSZtemp

    'DistanceofDtoVSZtemp1 = Log(1 + Exp(2 * DistanceofDtoVSZtemp / 1000)) / Log(2.71) / (2 * DistanceofDtoVSZtemp)

    'If (MoveDirectionDistance(j) - DistanceofDtoVSZtemp < DistanceofDtoVSZtemp1) Then
    'If (MoveDirectionDistance(j) < DistanceofDtoVSZtemp) Then
    If (DistanceofDtoVSZtemp - MoveDirectionDistance(j) > 0.002) Then
          V(m).X = MoveDirectionV(j).X: V(m).Y = MoveDirectionV(j).Y   '顶点新的位置
          Call SegmentExpression(V, tmin)        '求线段的uxy(),tsx()   V(1)与V(2)间是线段uxy(1)
          Call DataProject(D(), V, uxy, tsx)     '入口:数据点,顶点,各线段单位矢量,各线段投影指标初值
         
         For i = LBound(V) To UBound(V)                   '判定各夹角是否为锐角
            bz = False
            If (Pi(i) - 1) >= 0.01 Then bz = True: Exit For  '角度惩罚
         Next i
         
         If (bz = True) Then
                V(m).X = Vtemp0.X: V(m).Y = Vtemp0.Y
         End If
    Else
         V(m).X = Vtemp0.X: V(m).Y = Vtemp0.Y
    End If
    Call SegmentExpression(V, tmin)        '求线段的uxy(),tsx()   V(1)与V(2)间是线段uxy(1)
    Call DataProject(D(), V, uxy, tsx)     '入口:数据点,顶点,各线段单位矢量,各线段投影指标初值
Adjust1Point_Exit:

End Sub
'Public Sub Adjust1PointSub(ByVal m As Integer, ByVal DistanceofDtoVSZtemp As Double, ByRef Vtemp1 As xy)
'    Dim bz As Boolean
'    Dim i As Integer
'    If (Abs(V(m).x) < 0.999) And (Abs(V(m).y) < 0.999) Then
'        Call SegmentExpression(V, tmin)        '求线段的uxy(),tsx()   V(1)与V(2)间是线段uxy(1)
'        Call DataProject(D(), V, uxy, tsx)     '入口:数据点,顶点,各线段单位矢量,各线段投影指标初值
'                                               'DataProject会改变DistanceofDtoVSZ
'        If (DistanceofDtoVSZtemp > DistanceofDtoVSZ) Then
'           'DistanceofDtoVSZ = DistanceofDtoVSZtemp
'           For i = LBound(V) To UBound(V)                   '判定各夹角是否为锐角
'              bz = False
'              If (Pi(i) - 1) >= 0 Then bz = True: Exit For  '角度惩罚
'           Next i
'           If bz = False Then Vtemp1.x = V(m).x: Vtemp1.y = V(m).y
'       End If
'    End If
'End Sub




Public Sub Adjust1PointSub1(ByVal i As Integer, ByVal m As Integer, ByVal DistanceofDtoVSZtemp As Double, ByRef Vtemp1 As xy)
     Call SegmentExpression(V, tmin)        '求线段的uxy(),tsx()   V(1)与V(2)间是线段uxy(1)
     Call DataProject(D(), V, uxy, tsx)     '入口:数据点,顶点,各线段单位矢量,各线段投影指标初值
     MoveDirectionDistance(i) = DistanceofDtoVSZ '数据点到曲线的总距离平方
     MoveDirectionV(i).X = V(m).X: MoveDirectionV(i).Y = V(m).Y        '8个新顶点
End Sub



