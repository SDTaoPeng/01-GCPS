Attribute VB_Name = "Mod061SegmentExpression"
'Name: SegmentExpression(顶点V(),投影指标最小值tmin)
'入口:顶点V(),投影指标最小值tmin (全局变量)
'出口:顶点V()不变
'     uxy() 各线段单位矢量数组  tsx() 各线段投影指标初值数组
'调用例:Call SegmentExpression(V, tmin)   '求线段的uxy(),tsx()

Public Sub SegmentExpression(ByRef Vex() As xy, ByVal tmin As Double)
   '
   Dim i As Integer, j As Integer, K As Integer
   Dim t1 As Double, d1 As Double
   Dim Vpoint As xy
   '
   ReDim tsx(1 To UBound(Vex))        '重新定义各线段投影指标初值数组
   ReDim uxy(1 To UBound(Vex) - 1)    '重新各线段单位矢量数组
   tsx(1) = tmin                      '第1线段的投影指标初值
   '各线段单位矢量数组uxy()
   For i = 1 To UBound(Vex) - 1
       '两点间欧氏距离
       d1 = Sqr((Vex(i + 1).X - Vex(i).X) * (Vex(i + 1).X - Vex(i).X) + (Vex(i + 1).Y - Vex(i).Y) * (Vex(i + 1).Y - Vex(i).Y))
       '第i线段单位矢量
       uxy(i).X = (Vex(i + 1).X - Vex(i).X) / d1
       uxy(i).Y = (Vex(i + 1).Y - Vex(i).Y) / d1
   Next i
  
   '计算投影指标tsx()
   For i = 1 To UBound(Vex) - 1
'       If Abs(uxy(i).X) > 0.5 Then    '原则上两式计算结果相同，但考虑一个很小情况该用另一个
'          tsx(i + 1) = (Vex(i + 1).X - Vex(i).X) / uxy(i).X + tsx(i)
'       Else
'          tsx(i + 1) = (Vex(i + 1).Y - Vex(i).Y) / uxy(i).Y + tsx(i)   '与上式计算结果相同
'       End If
       
        d1 = Sqr((Vex(i + 1).X - Vex(i).X) * (Vex(i + 1).X - Vex(i).X) + (Vex(i + 1).Y - Vex(i).Y) * (Vex(i + 1).Y - Vex(i).Y))
        tsx(i + 1) = d1 + tsx(i)
            
   Next i
   '线段的表达方式:第i个线段(由第i到第i+1顶点构成)  t1属于[tsx(i),tsx(i+1)]
    'Vpoint.x = Vex(i).x + (t1 -tsx(i) ) * uxy(i).x
    'Vpoint.y = Vex(i).y + (t1 - tsx(i)) * uxy(i).y
End Sub

'Name: DataProject()
'入口:D()数据点,V()顶点,uxy() 各线段单位矢量数组,tsx() 各线段投影指标初值
'出口:全局变量:PV():顶点的角度惩罚 DairTa():顶点的距离约束总和:GV():顶点的距离约束+角度惩罚 总和
'      DistanceofDtoVSZ :数据点到曲线的总距离平方
'调用例:Call DataProject(D(), V, uxy, tsx)     '入口:数据点,顶点,各线段单位矢量,各线段投影指标初值

Public Sub DataProject(ByRef D() As xy, ByRef v() As xy, ByRef uxy() As xy, ByRef tsx() As Double)
   Dim i As Integer, j As Integer, n As Integer, K As Integer
   Dim t1 As Double, Drtsx As Double, d1 As Double, namnapp As Double, namnap As Double
   Dim ProjectPoint As xy
 
   namnapp = 0.13
   
   
   ReDim DtoVS(1 To UBound(D))             '数据点投影标识 1-20000 为属于顶点 20000以上为属于线段
   ReDim DistanceofDtoVS(1 To UBound(D))   '数据点到投影处的距离平方
   '计算DtoVS(1 To UBound(D)),DistanceofDtoVS(1 To UBound(D))
   For j = 1 To UBound(D)  '对数据点循环(开始)
        DtoVS(j) = 0                    '数据点投影标识(初值---定一个不可能的数,以便循环编程统一)
        DistanceofDtoVS(j) = 1000       '数据点到投影处的距离平方(初值---定一个不可能的大数,以便循环编程统一[数据点已标准化])
        For i = 1 To UBound(v) - 1      '对线段循环(开始)
             '对数据点D(j)计算到各线段uxy(i)或顶点V(i)距离平方Drtsx(分三种情况进行计算)
             t1 = (D(j).X - v(i).X) * uxy(i).X + (D(j).Y - v(i).Y) * uxy(i).Y + tsx(i)  'D(j)到线段(i)的投影指标
             If t1 <= tsx(i) Then    '在V(i)之外,取到顶点v(i)的距离
                Drtsx = (D(j).X - v(i).X) * (D(j).X - v(i).X) + (D(j).Y - v(i).Y) * (D(j).Y - v(i).Y)
                If Drtsx <= DistanceofDtoVS(j) Then DtoVS(j) = i: DistanceofDtoVS(j) = Drtsx '投影标识,点到投影处的距离平方
             Else
                If t1 >= tsx(i + 1) Then   '在V(i+1)之外,取到顶点v(i+1)的距离
                  Drtsx = (D(j).X - v(i + 1).X) * (D(j).X - v(i + 1).X) + (D(j).Y - v(i + 1).Y) * (D(j).Y - v(i + 1).Y)
                  If Drtsx <= DistanceofDtoVS(j) Then DtoVS(j) = i + 1: DistanceofDtoVS(j) = Drtsx '投影标识,点到投影处的距离平方
                Else                       '在V(i)-V(i+1)之间,计算投影点,求到投影点的距离
                  ProjectPoint.X = v(i).X + (t1 - tsx(i)) * uxy(i).X
                  ProjectPoint.Y = v(i).Y + (t1 - tsx(i)) * uxy(i).Y
                  Drtsx = (D(j).X - ProjectPoint.X) * (D(j).X - ProjectPoint.X) + (D(j).Y - ProjectPoint.Y) * (D(j).Y - ProjectPoint.Y)
                  If Drtsx <= DistanceofDtoVS(j) Then DtoVS(j) = 20000 + i: DistanceofDtoVS(j) = Drtsx '投影标识,点到投影处的距离平方
                End If
             End If
         Next i                         '对线段循环(结束)
    Next j                '对数据点循环(结束)
    '验证
'    i = 56
'    FrmPC.PicC_Qc.Print i
'    FrmPC.PicC_Qc.Print DtoVS(i)
'    FrmPC.PicC_Qc.Print DistanceofDtoVS(i)
'    Call DrawData(FrmPC.PicC_Qc, D(i), vbBlue, "DrawForkX", 150)     '在图片框中,画点,颜色,形状,大小
     '-------------------------------------------------------------------------------------------------
     '对于顶点进行计算
     '顶点优化步变量
     K = UBound(uxy)             '线段个数
     ReDim Cgm(1 To K)           '隶属于线段数据到该线段Si的距离平方
     ReDim VV(1 To K + 1)        '隶属于顶点的数据到该顶点Vi的距离平方
     ReDim u2(1 To K)            '各线段长度平方
     ReDim Pi(1 To K + 1)        '顶点的角度惩罚
     ReDim PV(1 To K + 1)        '顶点的角度惩罚总和
     ReDim DairTa(1 To K + 1)    '顶点的距离约束总和
     ReDim GV(1 To K + 1)        '顶点的距离约束+角度惩罚 总和
     '------------------------------------------------------------------------
     '计算Cgm(1 To k + 1) 隶属于顶点后线段数据到该线段Si的距离平方
     For i = 1 To K                '对线段循环
         Cgm(i) = 0
         For j = 1 To UBound(D)    '对数据循环
             If ((DtoVS(j) - 20000)) = i Then Cgm(i) = Cgm(i) + DistanceofDtoVS(j)
         Next j
     Next i
     '计算VV(1 To k + 1) 隶属于顶点的数据到该顶点Vi的距离平方
     For i = 1 To K + 1            '对线段循环
         VV(i) = 0
         For j = 1 To UBound(D)    '对数据循环
             If DtoVS(j) = i Then VV(i) = VV(i) + DistanceofDtoVS(j)
         Next j
     Next i
     '计算u2(1 To k)   '各线段长度平方
     For i = 1 To K             '对线段循环
         u2(i) = (v(i + 1).X - v(i).X) * (v(i + 1).X - v(i).X) + (v(i + 1).Y - v(i).Y) * (v(i + 1).Y - v(i).Y)
     Next i
     '计算Pi(1 To k + 1) 顶点的角度惩罚
     Pi(1) = 0: Pi(K + 1) = 0
     For i = 2 To K
         '
         d1 = (v(i - 1).X - v(i).X) * (v(i + 1).X - v(i).X) + (v(i - 1).Y - v(i).Y) * (v(i + 1).Y - v(i).Y)
         
         t1 = Sqr((v(i - 1).X - v(i).X) * (v(i - 1).X - v(i).X) + (v(i - 1).Y - v(i).Y) * (v(i - 1).Y - v(i).Y))
         t1 = t1 * Sqr((v(i + 1).X - v(i).X) * (v(i + 1).X - v(i).X) + (v(i + 1).Y - v(i).Y) * (v(i + 1).Y - v(i).Y))
          'COSri
         Pi(i) = 1 + d1 / t1    '取r=1
     Next i
     '计算PV(1 To k + 1) 顶点的角度惩罚总和
     For i = 1 To K + 1 '对顶点循环
         If i = 1 Then PV(i) = u2(1) + Pi(2)
         If i = 2 Then PV(i) = u2(1) + Pi(2) + Pi(3)
         If (i > 2) And (i < K) Then PV(i) = Pi(i - 1) + Pi(i) + Pi(i + 1)
         If i = K Then PV(i) = Pi(i - 1) + Pi(i) + u2(i)
         If i = K + 1 Then PV(i) = Pi(i - 1) + u2(i - 1)
         PV(i) = PV(i) / (K + 1)
     Next i
 
     '计算DairTa(1 To k + 1)   顶点的距离约束总和
      n = UBound(D)
     For i = 1 To K + 1 '对顶点循环
         If i = 1 Then DairTa(i) = VV(i) + Cgm(i)                                 'i=1
         If (i > 1) And (i < K + 1) Then DairTa(i) = Cgm(i - 1) + VV(i) + Cgm(i)  '1<i<k+1
         If i = K + 1 Then DairTa(i) = Cgm(i - 1) + VV(i)                         'i=k+1
         DairTa(i) = DairTa(i) / n
     Next i
     '计算GV(1 To k + 1) 顶点的距离约束+角度惩罚 总和
     d1 = 0
     For i = 1 To n: d1 = d1 + DistanceofDtoVS(i): Next i '数据点到曲线的总距离平方
     DistanceofDtoVSZ = d1  '数据点到曲线的总距离平方
     namnap = namnapp * K * (1 / ((n) ^ (1 / 3))) * Sqr(d1)
     For i = 1 To K + 1 '对顶点循环
         GV(i) = DairTa(i) + PV(i) * namnap                             '惩罚距离
     Next i
     'FrmPC.PicC_Qc.Print GV(2)
End Sub

'Name: UpdateArray(旧顶点数组V(),临时数组Vextemp(),第m个点需要删除,点值为valuex和valuey)
'入口:顶点数组V(),删除点m和value
'出口:顶点数组V()更新
'     数组更新函数（删除数组V中的点m）

Public Sub UpdateArray(ByRef Vex() As xy, ByVal m As Integer, ByVal valuex As Double, ByVal valuey As Double)
   '
    Dim i As Integer
    Dim t As Integer
    Dim Vtemp As xy, Vextemp() As xy             '中间变量点Vtemp和数组Vextemp

    ReDim Vextemp(1 To UBound(Vex) - 1)            '中间变量数组Vextemp
    
    For i = 1 To UBound(Vex)            '对数组Vex()循环
        If (v(i).X <> valuex) And (v(i).Y <> valuey) Then
            If (i < m) Then
                Vextemp(i).X = v(i).X
                Vextemp(i).Y = v(i).Y
            Else
                If (i = m) Then
                    t = t + 1          '定义无用变量
                Else
                    If (i > m) Then
                        Vextemp(i - 1).X = v(i).X
                        Vextemp(i - 1).Y = v(i).Y
                    End If
                End If
            End If
        End If
'ExitLoop:
    Next i
    
    ReDim Vex(1 To UBound(Vextemp))
    For i = 1 To UBound(Vextemp)            '对数组Vex()循环
        v(i).X = Vextemp(i).X
        v(i).Y = Vextemp(i).Y
    Next i
End Sub

'Name: CalculateAngle(点vL,v,vR；点距离disvLv和disvvR)
'入口:3个顶点的坐标，和2条边的长度
'出口:计算角度

Public Sub CalculateAngle(ByRef vL As xy, ByRef v As xy, ByRef vR As xy, ByVal disvLv As Double, ByVal disvvR As Double, ByVal angle As Double)
    
    Dim xtemp As Double
    xtemp = (vL.X - v.X) * (vR.X - v.X) + (vL.Y - v.Y) * (vR.Y - v.Y)
    xtemp = xtemp / disvLv / disvvR
    angle = Atn(Sqr(1 - X ^ 2) / xtemp)

End Sub

