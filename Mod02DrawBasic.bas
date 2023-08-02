Attribute VB_Name = "Mod02DrawBasic"

'-------------------------------------------------------------------------------------------------------
'Name:Drawcoordinate(图片框名称,坐标轴颜色,刻度颜色,边框颜色)
'Function:(1)在图片框中画坐标及刻度;
'         (2)坐标原点移到图片框中心;
'         (3)计算出全局变量 Kx、Ky (单位物理量的屏幕长度度量单位)
'         (4)计算出全局变量 Px0、Py0(0点位置)
'         (5)把变量限制在(-1,-1)-(1,1)之间,中心为原点
'         (6)颜色使用VB标准色,如vbBlue,vbWhite,vbYellow等
'         (7)若不画坐标,可使用vbWhite的坐标颜色及刻度颜色，若不画边框，可用vbWhite
'出口参数：全局变量 Kx  Ky  Px0  Py0
'调用例：Call Drawcoordinate(PicC_Qc, vbBlue, vbRed, vbGreen)  '在图片框中画坐标轴,坐标轴颜色,刻度颜色,边框颜色
Public Sub Drawcoordinate(ByVal Pic As PictureBox, ByVal CoordinateColor As Long, ByVal ScaleColor As Long, frameColor As Long)
    Dim Dx As Double, X As Double, Y As Double
    Dim Px As Double, Py As Double
    '确定0点位置
    Px0 = (Pic.Width - 200) / 2 + 100: Py0 = (Pic.Height - 200) / 2 + 100
    '画横坐标及坐标刻度
    Pic.Line (100, Py0)-(Pic.Width - 100, Py0), CoordinateColor, B                    '画横坐标
    Kx = (Pic.Width - 200) / 2                                 '单位长度的屏幕长度度量单位
    For X = -1 To 1 Step 0.1                                                     '画横坐标刻度
       Px = Kx * X + Px0: Py = Py0 - 80: Px1 = Kx * X + Px0: Py1 = Py + 160
       Pic.Line (Px, Py)-(Px1, Py1), ScaleColor, B
    Next X
     '画纵坐标及坐标刻度
    Pic.Line (Px0, 100)-(Px0, Pic.Height - 100), CoordinateColor, B                   '画纵坐标
    Ky = -(Pic.Height - 200) / 2                                '单位长度的屏幕长度度量单位
    For Y = -1 To 1 Step 0.1                                                      '画纵坐标刻度
        Px = Px0 - 80: Py = Ky * Y + Py0: Px1 = Px + 160: Py1 = Ky * Y + Py0
        Pic.Line (Px, Py)-(Px1, Py1), ScaleColor, B
    Next Y
    '画矩形
    Pic.Line (100, 100)-(Pic.Width - 100, Pic.Height - 100), frameColor, B
End Sub


'-------------------------------------------------------------------------------------------------
'Name:DrawData(图片框名称、坐标点、颜色、形状、大小)   （画一点）
'Function:(1)在图片框中某一点画基本形状
'         (2)坐标点为自定义类型xy有二值(xy.x,xy.y)
'         (3)颜色使用VB标准色,如vbBlue,vbWhite,vbYellow等
'         (4)形状为字符型,有"DrawCircle","DrawForkX","DrawRectangle"三种
'         (5)若只画一点,可用半径为1的圆
'出口参数：无
'调用例：Call DrawData(PicC_Qc, PointA, vbRed, "DrawCircle", 1)     '在图片框中,画点,颜色,形状,大小
Public Sub DrawData(ByVal Pic As PictureBox, DataPoint As xy, Color As Long, ByVal shape As String, ByVal r As Integer)
        Dim Px As Double, Py As Double
        Px = Kx * DataPoint.X + Px0: Py = Ky * DataPoint.Y + Py0
     Select Case shape
        Case "DrawCircle"     '在(DataPoint.x,DataPoint.y)处画一半径为r、颜色为color的圆(半径为1即为画点)
             Pic.Circle (Px, Py), r, Color
        Case "DrawForkX"      '在(DataPoint.x,DataPoint.y)处画一半径为r、颜色为color的"×"
             Pic.Line (Px - r, Py - r)-(Px + r, Py + r), Color
             Pic.Line (Px + r, Py - r)-(Px - r, Py + r), Color
             Pic.DrawWidth = 10
             
        Case "DrawRectangle"  '在(DataPoint.x,DataPoint.y)处画一半径为r、颜色为color的矩形
             Pic.Line (Px - r, Py - r)-(Px + r, Py + r), Color, B
     End Select
End Sub


'-------------------------------------------------------------------------------------------------
'Name:DrawLine(图片框名称、坐标点A、坐标点B、颜色、粗细)   （画两点间的直线）
'Function:(1)在图片框中画基本形状---线条
'         (2)坐标点为自定义类型xy有二值(xy.x,xy.y)
'         (3)颜色使用VB标准色,如vbBlue,vbWhite,vbYellow等
'出口参数：无
'调用例：Call DrawLine(PicC_Qc, PointA,  PointB,vbRed, 1)     '在图片框中,画线,颜色,粗细
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








