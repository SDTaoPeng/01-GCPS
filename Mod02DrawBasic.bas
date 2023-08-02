Attribute VB_Name = "Mod02DrawBasic"

'-------------------------------------------------------------------------------------------------------
'Name:Drawcoordinate(ͼƬ������,��������ɫ,�̶���ɫ,�߿���ɫ)
'Function:(1)��ͼƬ���л����꼰�̶�;
'         (2)����ԭ���Ƶ�ͼƬ������;
'         (3)�����ȫ�ֱ��� Kx��Ky (��λ����������Ļ���ȶ�����λ)
'         (4)�����ȫ�ֱ��� Px0��Py0(0��λ��)
'         (5)�ѱ���������(-1,-1)-(1,1)֮��,����Ϊԭ��
'         (6)��ɫʹ��VB��׼ɫ,��vbBlue,vbWhite,vbYellow��
'         (7)����������,��ʹ��vbWhite��������ɫ���̶���ɫ���������߿򣬿���vbWhite
'���ڲ�����ȫ�ֱ��� Kx  Ky  Px0  Py0
'��������Call Drawcoordinate(PicC_Qc, vbBlue, vbRed, vbGreen)  '��ͼƬ���л�������,��������ɫ,�̶���ɫ,�߿���ɫ
Public Sub Drawcoordinate(ByVal Pic As PictureBox, ByVal CoordinateColor As Long, ByVal ScaleColor As Long, frameColor As Long)
    Dim Dx As Double, X As Double, Y As Double
    Dim Px As Double, Py As Double
    'ȷ��0��λ��
    Px0 = (Pic.Width - 200) / 2 + 100: Py0 = (Pic.Height - 200) / 2 + 100
    '�������꼰����̶�
    Pic.Line (100, Py0)-(Pic.Width - 100, Py0), CoordinateColor, B                    '��������
    Kx = (Pic.Width - 200) / 2                                 '��λ���ȵ���Ļ���ȶ�����λ
    For X = -1 To 1 Step 0.1                                                     '��������̶�
       Px = Kx * X + Px0: Py = Py0 - 80: Px1 = Kx * X + Px0: Py1 = Py + 160
       Pic.Line (Px, Py)-(Px1, Py1), ScaleColor, B
    Next X
     '�������꼰����̶�
    Pic.Line (Px0, 100)-(Px0, Pic.Height - 100), CoordinateColor, B                   '��������
    Ky = -(Pic.Height - 200) / 2                                '��λ���ȵ���Ļ���ȶ�����λ
    For Y = -1 To 1 Step 0.1                                                      '��������̶�
        Px = Px0 - 80: Py = Ky * Y + Py0: Px1 = Px + 160: Py1 = Ky * Y + Py0
        Pic.Line (Px, Py)-(Px1, Py1), ScaleColor, B
    Next Y
    '������
    Pic.Line (100, 100)-(Pic.Width - 100, Pic.Height - 100), frameColor, B
End Sub


'-------------------------------------------------------------------------------------------------
'Name:DrawData(ͼƬ�����ơ�����㡢��ɫ����״����С)   ����һ�㣩
'Function:(1)��ͼƬ����ĳһ�㻭������״
'         (2)�����Ϊ�Զ�������xy�ж�ֵ(xy.x,xy.y)
'         (3)��ɫʹ��VB��׼ɫ,��vbBlue,vbWhite,vbYellow��
'         (4)��״Ϊ�ַ���,��"DrawCircle","DrawForkX","DrawRectangle"����
'         (5)��ֻ��һ��,���ð뾶Ϊ1��Բ
'���ڲ�������
'��������Call DrawData(PicC_Qc, PointA, vbRed, "DrawCircle", 1)     '��ͼƬ����,����,��ɫ,��״,��С
Public Sub DrawData(ByVal Pic As PictureBox, DataPoint As xy, Color As Long, ByVal shape As String, ByVal r As Integer)
        Dim Px As Double, Py As Double
        Px = Kx * DataPoint.X + Px0: Py = Ky * DataPoint.Y + Py0
     Select Case shape
        Case "DrawCircle"     '��(DataPoint.x,DataPoint.y)����һ�뾶Ϊr����ɫΪcolor��Բ(�뾶Ϊ1��Ϊ����)
             Pic.Circle (Px, Py), r, Color
        Case "DrawForkX"      '��(DataPoint.x,DataPoint.y)����һ�뾶Ϊr����ɫΪcolor��"��"
             Pic.Line (Px - r, Py - r)-(Px + r, Py + r), Color
             Pic.Line (Px + r, Py - r)-(Px - r, Py + r), Color
             Pic.DrawWidth = 10
             
        Case "DrawRectangle"  '��(DataPoint.x,DataPoint.y)����һ�뾶Ϊr����ɫΪcolor�ľ���
             Pic.Line (Px - r, Py - r)-(Px + r, Py + r), Color, B
     End Select
End Sub


'-------------------------------------------------------------------------------------------------
'Name:DrawLine(ͼƬ�����ơ������A�������B����ɫ����ϸ)   ����������ֱ�ߣ�
'Function:(1)��ͼƬ���л�������״---����
'         (2)�����Ϊ�Զ�������xy�ж�ֵ(xy.x,xy.y)
'         (3)��ɫʹ��VB��׼ɫ,��vbBlue,vbWhite,vbYellow��
'���ڲ�������
'��������Call DrawLine(PicC_Qc, PointA,  PointB,vbRed, 1)     '��ͼƬ����,����,��ɫ,��ϸ
Public Sub DrawLine(ByVal Pic As PictureBox, DataPointA As xy, DataPointB As xy, Color As Long, ByVal r As Integer)
        Dim Pxa As Double, Pya As Double
        Dim Pxb As Double, Pyb As Double
        Pxa = Kx * DataPointA.X + Px0: Pya = Ky * DataPointA.Y + Py0
        Pxb = Kx * DataPointB.X + Px0: Pyb = Ky * DataPointB.Y + Py0
        Pic.Line (Pxa, Pya)-(Pxb, Pyb), Color
        Pic.Line (Pxa, Pya - r)-(Pxb, Pyb - r), Color
        Pic.Line (Pxa - r, Pya)-(Pxb - r, Pyb), Color
        Pic.Line (Pxa, Pya + r)-(Pxb, Pyb + r), Color
        Pic.Line (Pxa + r, Pya)-(Pxb + r, Pyb), Color
End Sub








