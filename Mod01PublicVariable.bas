Attribute VB_Name = "Mod01PublicVariable"
'画图时使用的变量（Drawcoordinate模块计算得出）
Public Px0 As Integer, Py0 As Integer '(0点位置)
Public Kx As Double, Ky As Double     '(单位物理量的屏幕长度度量单位)



Public DataFileName As String
Public PCAandTxyFileName As String



'投影存盘使用的数据
'平面数据点类型定义
Public Type xy
     X As Double
     Y As Double
End Type
Public CurcvsPoint(0 To 9999) As xy    '曲线点
'全局变量声明
Public Tfx() As Double    '投影指标   (向量,以各数据点为元素)
Public Fxy() As Double    '数据点矩阵 (第1下标为数据点序号,第2下标为数据维数)
                          '每个数据为行向量(Fxy(i,1),Fxy(i,2))
                          
Public xymax As Double, Sumx As Double, Sumy As Double

'从文件中读数据
Public Nd As Integer            '原始数据个数
Public D() As xy                '数据点(D().x,D().y)
'PCA-RTB
Public t() As Double               '投影指标(Projection Index)
Public ta As Double, tb As Double  'PCA-RTB参数
'
Public V() As xy                      '顶点
Public tmax As Double, tmin As Double '投影指标最小、最大值
'
Public uxy() As xy                       'uxy()各线段单位矢量数组
Public tsx() As Double                   'tsx()各线段投影指标初值数组
Public DataNumberofSegment() As Integer  '隶属于各线段的数据个数

'
Public DtoVS() As Integer          '数据点投影标识 1-20000 为属于顶点 20000以上为属于线段
Public DistanceofDtoVS() As Double '数据点到投影处的距离平方
'顶点优化步变量
Public Cgm() As Double             '隶属于线段数据到该线段Si的距离平方
Public VV() As Double              '隶属于顶点的数据到该顶点Vi的距离平方
Public u2() As Double              '各线段长度平方

Public Pi() As Double              '顶点的角度惩罚
Public PV() As Double              '顶点的角度惩罚总和
Public DairTa() As Double          '顶点的距离约束总和
Public GV() As Double              '顶点的距离约束+角度惩罚 总和
Public DistanceofDtoVSZ As Double  '数据点到曲线的总距离平方

Public MoveDirectionDistance(1 To 8) As Double  '数据点到曲线的总距离平方（当前）
Public MoveDirectionV(1 To 8) As xy             '8个新顶点


'
Public VnewSerialNumber  As Integer




   


