Attribute VB_Name = "Mod061SegmentExpression"
'Name: SegmentExpression(����V(),ͶӰָ����Сֵtmin)
'���:����V(),ͶӰָ����Сֵtmin (ȫ�ֱ���)
'����:����V()����
'     uxy() ���߶ε�λʸ������  tsx() ���߶�ͶӰָ���ֵ����
'������:Call SegmentExpression(V, tmin)   '���߶ε�uxy(),tsx()

Public Sub SegmentExpression(ByRef Vex() As xy, ByVal tmin As Double)
   '
   Dim i As Integer, j As Integer, K As Integer
   Dim t1 As Double, d1 As Double
   Dim Vpoint As xy
   '
   ReDim tsx(1 To UBound(Vex))        '���¶�����߶�ͶӰָ���ֵ����
   ReDim uxy(1 To UBound(Vex) - 1)    '���¸��߶ε�λʸ������
   tsx(1) = tmin                      '��1�߶ε�ͶӰָ���ֵ
   '���߶ε�λʸ������uxy()
   For i = 1 To UBound(Vex) - 1
       '�����ŷ�Ͼ���
       d1 = Sqr((Vex(i + 1).X - Vex(i).X) * (Vex(i + 1).X - Vex(i).X) + (Vex(i + 1).Y - Vex(i).Y) * (Vex(i + 1).Y - Vex(i).Y))
       '��i�߶ε�λʸ��
       uxy(i).X = (Vex(i + 1).X - Vex(i).X) / d1
       uxy(i).Y = (Vex(i + 1).Y - Vex(i).Y) / d1
   Next i
  
   '����ͶӰָ��tsx()
   For i = 1 To UBound(Vex) - 1
'       If Abs(uxy(i).X) > 0.5 Then    'ԭ������ʽ��������ͬ��������һ����С���������һ��
'          tsx(i + 1) = (Vex(i + 1).X - Vex(i).X) / uxy(i).X + tsx(i)
'       Else
'          tsx(i + 1) = (Vex(i + 1).Y - Vex(i).Y) / uxy(i).Y + tsx(i)   '����ʽ��������ͬ
'       End If
       
        d1 = Sqr((Vex(i + 1).X - Vex(i).X) * (Vex(i + 1).X - Vex(i).X) + (Vex(i + 1).Y - Vex(i).Y) * (Vex(i + 1).Y - Vex(i).Y))
        tsx(i + 1) = d1 + tsx(i)
            
   Next i
   '�߶εı�﷽ʽ:��i���߶�(�ɵ�i����i+1���㹹��)  t1����[tsx(i),tsx(i+1)]
    'Vpoint.x = Vex(i).x + (t1 -tsx(i) ) * uxy(i).x
    'Vpoint.y = Vex(i).y + (t1 - tsx(i)) * uxy(i).y
End Sub

'Name: DataProject()
'���:D()���ݵ�,V()����,uxy() ���߶ε�λʸ������,tsx() ���߶�ͶӰָ���ֵ
'����:ȫ�ֱ���:PV():����ĽǶȳͷ� DairTa():����ľ���Լ���ܺ�:GV():����ľ���Լ��+�Ƕȳͷ� �ܺ�
'      DistanceofDtoVSZ :���ݵ㵽���ߵ��ܾ���ƽ��
'������:Call DataProject(D(), V, uxy, tsx)     '���:���ݵ�,����,���߶ε�λʸ��,���߶�ͶӰָ���ֵ

Public Sub DataProject(ByRef D() As xy, ByRef v() As xy, ByRef uxy() As xy, ByRef tsx() As Double)
   Dim i As Integer, j As Integer, n As Integer, K As Integer
   Dim t1 As Double, Drtsx As Double, d1 As Double, namnapp As Double, namnap As Double
   Dim ProjectPoint As xy
 
   namnapp = 0.13
   
   
   ReDim DtoVS(1 To UBound(D))             '���ݵ�ͶӰ��ʶ 1-20000 Ϊ���ڶ��� 20000����Ϊ�����߶�
   ReDim DistanceofDtoVS(1 To UBound(D))   '���ݵ㵽ͶӰ���ľ���ƽ��
   '����DtoVS(1 To UBound(D)),DistanceofDtoVS(1 To UBound(D))
   For j = 1 To UBound(D)  '�����ݵ�ѭ��(��ʼ)
        DtoVS(j) = 0                    '���ݵ�ͶӰ��ʶ(��ֵ---��һ�������ܵ���,�Ա�ѭ�����ͳһ)
        DistanceofDtoVS(j) = 1000       '���ݵ㵽ͶӰ���ľ���ƽ��(��ֵ---��һ�������ܵĴ���,�Ա�ѭ�����ͳһ[���ݵ��ѱ�׼��])
        For i = 1 To UBound(v) - 1      '���߶�ѭ��(��ʼ)
             '�����ݵ�D(j)���㵽���߶�uxy(i)�򶥵�V(i)����ƽ��Drtsx(������������м���)
             t1 = (D(j).X - v(i).X) * uxy(i).X + (D(j).Y - v(i).Y) * uxy(i).Y + tsx(i)  'D(j)���߶�(i)��ͶӰָ��
             If t1 <= tsx(i) Then    '��V(i)֮��,ȡ������v(i)�ľ���
                Drtsx = (D(j).X - v(i).X) * (D(j).X - v(i).X) + (D(j).Y - v(i).Y) * (D(j).Y - v(i).Y)
                If Drtsx <= DistanceofDtoVS(j) Then DtoVS(j) = i: DistanceofDtoVS(j) = Drtsx 'ͶӰ��ʶ,�㵽ͶӰ���ľ���ƽ��
             Else
                If t1 >= tsx(i + 1) Then   '��V(i+1)֮��,ȡ������v(i+1)�ľ���
                  Drtsx = (D(j).X - v(i + 1).X) * (D(j).X - v(i + 1).X) + (D(j).Y - v(i + 1).Y) * (D(j).Y - v(i + 1).Y)
                  If Drtsx <= DistanceofDtoVS(j) Then DtoVS(j) = i + 1: DistanceofDtoVS(j) = Drtsx 'ͶӰ��ʶ,�㵽ͶӰ���ľ���ƽ��
                Else                       '��V(i)-V(i+1)֮��,����ͶӰ��,��ͶӰ��ľ���
                  ProjectPoint.X = v(i).X + (t1 - tsx(i)) * uxy(i).X
                  ProjectPoint.Y = v(i).Y + (t1 - tsx(i)) * uxy(i).Y
                  Drtsx = (D(j).X - ProjectPoint.X) * (D(j).X - ProjectPoint.X) + (D(j).Y - ProjectPoint.Y) * (D(j).Y - ProjectPoint.Y)
                  If Drtsx <= DistanceofDtoVS(j) Then DtoVS(j) = 20000 + i: DistanceofDtoVS(j) = Drtsx 'ͶӰ��ʶ,�㵽ͶӰ���ľ���ƽ��
                End If
             End If
         Next i                         '���߶�ѭ��(����)
    Next j                '�����ݵ�ѭ��(����)
    '��֤
'    i = 56
'    FrmPC.PicC_Qc.Print i
'    FrmPC.PicC_Qc.Print DtoVS(i)
'    FrmPC.PicC_Qc.Print DistanceofDtoVS(i)
'    Call DrawData(FrmPC.PicC_Qc, D(i), vbBlue, "DrawForkX", 150)     '��ͼƬ����,����,��ɫ,��״,��С
     '-------------------------------------------------------------------------------------------------
     '���ڶ�����м���
     '�����Ż�������
     K = UBound(uxy)             '�߶θ���
     ReDim Cgm(1 To K)           '�������߶����ݵ����߶�Si�ľ���ƽ��
     ReDim VV(1 To K + 1)        '�����ڶ�������ݵ��ö���Vi�ľ���ƽ��
     ReDim u2(1 To K)            '���߶γ���ƽ��
     ReDim Pi(1 To K + 1)        '����ĽǶȳͷ�
     ReDim PV(1 To K + 1)        '����ĽǶȳͷ��ܺ�
     ReDim DairTa(1 To K + 1)    '����ľ���Լ���ܺ�
     ReDim GV(1 To K + 1)        '����ľ���Լ��+�Ƕȳͷ� �ܺ�
     '------------------------------------------------------------------------
     '����Cgm(1 To k + 1) �����ڶ�����߶����ݵ����߶�Si�ľ���ƽ��
     For i = 1 To K                '���߶�ѭ��
         Cgm(i) = 0
         For j = 1 To UBound(D)    '������ѭ��
             If ((DtoVS(j) - 20000)) = i Then Cgm(i) = Cgm(i) + DistanceofDtoVS(j)
         Next j
     Next i
     '����VV(1 To k + 1) �����ڶ�������ݵ��ö���Vi�ľ���ƽ��
     For i = 1 To K + 1            '���߶�ѭ��
         VV(i) = 0
         For j = 1 To UBound(D)    '������ѭ��
             If DtoVS(j) = i Then VV(i) = VV(i) + DistanceofDtoVS(j)
         Next j
     Next i
     '����u2(1 To k)   '���߶γ���ƽ��
     For i = 1 To K             '���߶�ѭ��
         u2(i) = (v(i + 1).X - v(i).X) * (v(i + 1).X - v(i).X) + (v(i + 1).Y - v(i).Y) * (v(i + 1).Y - v(i).Y)
     Next i
     '����Pi(1 To k + 1) ����ĽǶȳͷ�
     Pi(1) = 0: Pi(K + 1) = 0
     For i = 2 To K
         '
         d1 = (v(i - 1).X - v(i).X) * (v(i + 1).X - v(i).X) + (v(i - 1).Y - v(i).Y) * (v(i + 1).Y - v(i).Y)
         
         t1 = Sqr((v(i - 1).X - v(i).X) * (v(i - 1).X - v(i).X) + (v(i - 1).Y - v(i).Y) * (v(i - 1).Y - v(i).Y))
         t1 = t1 * Sqr((v(i + 1).X - v(i).X) * (v(i + 1).X - v(i).X) + (v(i + 1).Y - v(i).Y) * (v(i + 1).Y - v(i).Y))
          'COSri
         Pi(i) = 1 + d1 / t1    'ȡr=1
     Next i
     '����PV(1 To k + 1) ����ĽǶȳͷ��ܺ�
     For i = 1 To K + 1 '�Զ���ѭ��
         If i = 1 Then PV(i) = u2(1) + Pi(2)
         If i = 2 Then PV(i) = u2(1) + Pi(2) + Pi(3)
         If (i > 2) And (i < K) Then PV(i) = Pi(i - 1) + Pi(i) + Pi(i + 1)
         If i = K Then PV(i) = Pi(i - 1) + Pi(i) + u2(i)
         If i = K + 1 Then PV(i) = Pi(i - 1) + u2(i - 1)
         PV(i) = PV(i) / (K + 1)
     Next i
 
     '����DairTa(1 To k + 1)   ����ľ���Լ���ܺ�
      n = UBound(D)
     For i = 1 To K + 1 '�Զ���ѭ��
         If i = 1 Then DairTa(i) = VV(i) + Cgm(i)                                 'i=1
         If (i > 1) And (i < K + 1) Then DairTa(i) = Cgm(i - 1) + VV(i) + Cgm(i)  '1<i<k+1
         If i = K + 1 Then DairTa(i) = Cgm(i - 1) + VV(i)                         'i=k+1
         DairTa(i) = DairTa(i) / n
     Next i
     '����GV(1 To k + 1) ����ľ���Լ��+�Ƕȳͷ� �ܺ�
     d1 = 0
     For i = 1 To n: d1 = d1 + DistanceofDtoVS(i): Next i '���ݵ㵽���ߵ��ܾ���ƽ��
     DistanceofDtoVSZ = d1  '���ݵ㵽���ߵ��ܾ���ƽ��
     namnap = namnapp * K * (1 / ((n) ^ (1 / 3))) * Sqr(d1)
     For i = 1 To K + 1 '�Զ���ѭ��
         GV(i) = DairTa(i) + PV(i) * namnap                             '�ͷ�����
     Next i
     'FrmPC.PicC_Qc.Print GV(2)
End Sub

'Name: UpdateArray(�ɶ�������V(),��ʱ����Vextemp(),��m������Ҫɾ��,��ֵΪvaluex��valuey)
'���:��������V(),ɾ����m��value
'����:��������V()����
'     ������º�����ɾ������V�еĵ�m��

Public Sub UpdateArray(ByRef Vex() As xy, ByVal m As Integer, ByVal valuex As Double, ByVal valuey As Double)
   '
    Dim i As Integer
    Dim t As Integer
    Dim Vtemp As xy, Vextemp() As xy             '�м������Vtemp������Vextemp

    ReDim Vextemp(1 To UBound(Vex) - 1)            '�м��������Vextemp
    
    For i = 1 To UBound(Vex)            '������Vex()ѭ��
        If (v(i).X <> valuex) And (v(i).Y <> valuey) Then
            If (i < m) Then
                Vextemp(i).X = v(i).X
                Vextemp(i).Y = v(i).Y
            Else
                If (i = m) Then
                    t = t + 1          '�������ñ���
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
    For i = 1 To UBound(Vextemp)            '������Vex()ѭ��
        v(i).X = Vextemp(i).X
        v(i).Y = Vextemp(i).Y
    Next i
End Sub

'Name: CalculateAngle(��vL,v,vR�������disvLv��disvvR)
'���:3����������꣬��2���ߵĳ���
'����:����Ƕ�

Public Sub CalculateAngle(ByRef vL As xy, ByRef v As xy, ByRef vR As xy, ByVal disvLv As Double, ByVal disvvR As Double, ByVal angle As Double)
    
    Dim xtemp As Double
    xtemp = (vL.X - v.X) * (vR.X - v.X) + (vL.Y - v.Y) * (vR.Y - v.Y)
    xtemp = xtemp / disvLv / disvvR
    angle = Atn(Sqr(1 - X ^ 2) / xtemp)

End Sub

