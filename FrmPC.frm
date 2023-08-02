VERSION 5.00
Begin VB.Form FrmPC 
   BackColor       =   &H00E0E0E0&
   Caption         =   "闭合主曲线程序"
   ClientHeight    =   10350
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   16320
   Icon            =   "FrmPC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10350
   ScaleWidth      =   16320
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox TxtState 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   1
      Left            =   0
      TabIndex        =   24
      Text            =   "本过程得出[t,(x,y)]的关系,作为有监督BNNM的学习样本(输入)"
      Top             =   9240
      Width           =   6855
   End
   Begin VB.TextBox TxtState 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   0
      Left            =   210
      TabIndex        =   23
      Text            =   "利用角度惩罚方法求闭合主曲线(用分段直线表达,求出数据点的投影指标存盘)"
      Top             =   150
      Width           =   8385
   End
   Begin VB.Frame Frame2 
      Caption         =   "参数显示"
      Height          =   1935
      Left            =   960
      TabIndex        =   5
      Top             =   7080
      Width           =   4665
      Begin VB.TextBox TxtPara 
         Height          =   270
         Index           =   5
         Left            =   840
         TabIndex        =   15
         Text            =   "显示参数"
         Top             =   1500
         Width           =   3705
      End
      Begin VB.TextBox TxtPara 
         Height          =   270
         Index           =   4
         Left            =   840
         TabIndex        =   13
         Text            =   "显示参数"
         Top             =   1200
         Width           =   3705
      End
      Begin VB.TextBox TxtPara 
         Height          =   270
         Index           =   3
         Left            =   870
         TabIndex        =   11
         Text            =   "显示参数"
         Top             =   540
         Width           =   3705
      End
      Begin VB.TextBox TxtPara 
         Height          =   270
         Index           =   2
         Left            =   870
         TabIndex        =   9
         Text            =   "显示参数"
         Top             =   900
         Width           =   3705
      End
      Begin VB.TextBox TxtPara 
         Height          =   270
         Index           =   1
         Left            =   870
         TabIndex        =   6
         Text            =   "显示参数"
         Top             =   240
         Width           =   3705
      End
      Begin VB.Label Label1 
         Caption         =   "COS[0,-1]"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   570
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "总惩罚"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1530
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "距离惩罚"
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   10
         Top             =   1170
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "角度惩罚"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   8
         Top             =   870
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "顶点个数"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   270
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "操作区"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6165
      Left            =   960
      TabIndex        =   1
      Top             =   690
      Width           =   4725
      Begin VB.CommandButton CmdAuto 
         Caption         =   "自动"
         Height          =   495
         Left            =   3630
         TabIndex        =   28
         Top             =   2970
         Width           =   765
      End
      Begin VB.Frame Frame5 
         Caption         =   "[t,(x,y)]存盘路径与文件名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   240
         TabIndex        =   25
         Top             =   3630
         Width           =   4395
         Begin VB.TextBox SaveFileName 
            Height          =   465
            Left            =   210
            TabIndex        =   26
            Text            =   "SaveFileName"
            Top             =   390
            Width           =   3945
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "选择数据文件(SRCdata目录内)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2475
         Left            =   210
         TabIndex        =   20
         Top             =   390
         Width           =   4425
         Begin VB.TextBox TxtFile 
            Height          =   885
            Left            =   180
            MultiLine       =   -1  'True
            TabIndex        =   22
            Text            =   "FrmPC.frx":49E2
            Top             =   1440
            Width           =   4125
         End
         Begin VB.FileListBox File1 
            Height          =   810
            Left            =   210
            TabIndex        =   21
            Top             =   240
            Width           =   4125
         End
      End
      Begin VB.CommandButton CmdProjectAndSave 
         BackColor       =   &H00FF80FF&
         Caption         =   "投影存盘"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   390
         MaskColor       =   &H00E0E0E0&
         Picture         =   "FrmPC.frx":49EC
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4800
         Width           =   1785
      End
      Begin VB.CommandButton CmdInsert1V 
         Caption         =   "插入1个顶点"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2370
         TabIndex        =   16
         Top             =   2940
         Width           =   1155
      End
      Begin VB.CommandButton CmdAdjust 
         Caption         =   "调整"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2910
         Width           =   735
      End
      Begin VB.CommandButton CmdeExit 
         BackColor       =   &H00FF80FF&
         Caption         =   "退出"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   2340
         Picture         =   "FrmPC.frx":AC76
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4890
         Width           =   1725
      End
      Begin VB.CommandButton Cmdstart 
         Caption         =   "开始"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2880
         Width           =   1215
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "图形显示区"
      Height          =   9735
      Left            =   6840
      TabIndex        =   0
      Top             =   600
      Width           =   9375
      Begin VB.TextBox XTxt 
         Height          =   315
         Left            =   360
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   9360
         Width           =   3195
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFF80&
         Height          =   9045
         Left            =   210
         TabIndex        =   18
         Top             =   240
         Width           =   8865
         Begin VB.PictureBox PicC_Qc 
            BackColor       =   &H00FFFFFF&
            Height          =   8730
            Left            =   90
            ScaleHeight     =   8670
            ScaleMode       =   0  'User
            ScaleWidth      =   8670
            TabIndex        =   19
            Top             =   210
            Width           =   8730
         End
      End
   End
End
Attribute VB_Name = "FrmPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                       '检查未经声明的变量



Private Sub Form_Load()          '窗体初始化时执行

    File1.Path = "..\SRCdata"       '数据文件路径

    '-----------------------------------------------------
    Dim Iw As Double
    Dim Sw As String
    Dim X As Double, Y As Double
'    Open "..\SRCdata\Lizi2_YWC.TXT" For Output As #1       '打开文件
'    '逐行写入
'    For Iw = 0 To 6.283 Step 0.02          '0.2
'        X = Sin(Iw)  ',- Cos(0.8 * Iw) ^ 2 + 0.2 * Rnd(1)
'        Y = Cos(Iw)  '- Sin(2 * Iw) ^ 2 + 0.2 * Rnd(1)
'        Sw = Format$(X, "0.0000000000") & " " & Format$(Y, "0.0000000000")
'        Print #1, Sw             '写入1行(Sw)
'    Next Iw
'    '关闭文件
'    Close #1
'    '-----------------------------------------------------
  
  
    'PCAandTxyFileName = "..\Tdata\T_SmallNoise_S.txt"
    Cmdstart.Enabled = False           '"开始"按钮
    CmdAdjust.Enabled = False          '"调整"按钮
    CmdInsert1V.Enabled = False        '"插入1个顶点"按钮
    CmdProjectAndSave.Enabled = False  '"投影存盘"按钮
End Sub


Private Sub File1_Click()        '选择数据文件
    '----------------------------------执行的数据文件----------------------------------
    DataFileName = File1.Path & "\" & File1.FileName
    TxtFile.Text = "已经选择的数据文件为:" & "   ..\SRCdata\" & File1.FileName
    
    '----------------------------------保存的数据文件-------------------------------------
    PCAandTxyFileName = "..\Tdata\T_" & File1.FileName
    SaveFileName.Text = "[t,(x,y)]目标文件名:" & "   ..\Tdata\T_" & File1.FileName
    
    Cmdstart.Enabled = True           '"开始"按钮
    CmdAdjust.Enabled = False          '"调整"按钮
    CmdInsert1V.Enabled = False        '"插入1个顶点"按钮
    CmdProjectAndSave.Enabled = False  '"投影存盘"按钮
End Sub



'-------------------------------------------------------------------
Private Sub Cmdstart_ClIck()           '“开始”按钮
   '定义临时变量
   Dim i As Integer, j As Integer
   Dim Vpoint As xy, fPCA As xy
   Dim t1 As Double
   
   '在图片框中画坐标轴,刻度,边框
   ' -----------------------Normalization 把变量限制在(-1,-1)-(1,1)之间,中心为原点-----------------------
   Call Drawcoordinate(PicC_Qc, vbWhite, vbWhite, vbWhite)  ''在图片框中画坐标轴,坐标轴颜色,刻度颜色,边框颜色
   '打开文件,得出全局变量Nd---数据点个数, [D().x,D().y]---数据规范点值
   Call OpenTextFile(DataFileName) '打开文件,得出全局变量Nd---数据点个数, [D().x,D().y]---数据规范点值
 
   
   '画数据点[D().x,D().y]--------------------------------------不显示-------------------------------------
    'Call DrawDataPoint(D)
   '-----------------------------------------------------------------------------------------------------
   
'   '第一主成分
        'Call PCARTB(D)               '入口:(全局变量) 数据点个数Nd,数据点集合[D().x,D().y]
                                     '出口参数：第一主成分=(ta*t,tb*t)的ta,tb(全局变量)
        'Call JYSAstep1  '简易算法step1-数据规范点值D():第一主成分的ta,tb(全局变量)->顶点V(1) V(2) V(3) 投影指标tmin tmax
        'Call InsertV(3)
        '适当缩短第一主成分线段
        'tmin = tmin * 3 / 5: tmax = tmax * 3 / 5
        'V(1).X = tmin * ta: V(1).Y = tmin * tb
        'V(2)不变
        'V(3).X = tmax * ta: V(3).Y = tmax * tb
        '画第一主成分

       tmax = 0.99999
       tmin = 0.00001
       ReDim V(1 To 5)
'       V(1).X = -0.1: V(1).Y = -0.1
'       V(2).X = -0.1: V(2).Y = 0.1
'       V(3).X = 0.1: V(3).Y = 0.1
'       V(4).X = 0.1: V(4).Y = -0.1
'       V(5).X = -0.1: V(5).Y = -0.1
       
       V(1).X = -0.05: V(1).Y = -0.05
       V(2).X = -0.05: V(2).Y = 0.05
       V(3).X = 0.05: V(3).Y = 0.05
       V(4).X = 0.05: V(4).Y = -0.05
       V(5).X = -0.05: V(5).Y = -0.05
    
  
   For i = LBound(V) To UBound(V)
       Call DrawData(PicC_Qc, V(i), vbRed, "DrawCircle", 10)     '在图片框中,画点,颜色,形状,大小
   Next i
    Call SegmentExpression(V, tmin)   '求线段的uxy(),tsx()   V(1)与V(2)间是线段uxy(1)
     '用线段的表达方式画图
     For i = 1 To UBound(uxy)
          For t1 = tsx(i) To tsx(i + 1) Step 0.002
             '线段的表达方式:第i个线段(由第i到第i+1顶点构成)
             Vpoint.X = V(i).X + (t1 - tsx(i)) * uxy(i).X: Vpoint.Y = V(i).Y + (t1 - tsx(i)) * uxy(i).Y
             Call DrawData(PicC_Qc, Vpoint, vbRed, "DrawForkX", 2)     '在图片框PicC_Qc中,画V点,红色,圆点,半径为1
          Next t1
    Next i
    Cmdstart.Enabled = False           '"开始"按钮
    CmdAdjust.Enabled = True           '"调整"按钮
    CmdInsert1V.Enabled = True         '"插入1个顶点"按钮
    CmdProjectAndSave.Enabled = True   '"投影存盘"按钮
End Sub
'画数据点[D().x,D().y]
Private Sub DrawDataPoint(ByRef DataPonit() As xy)
   Dim i As Integer
   For i = 1 To UBound(D)                     '对所有数据点循环（画数据点)
      'Call DrawData(PicC_Qc, DataPonit(i), vbBlack, "DrawCircle", 10)      '在图片框PicC_Qc中,画D(i)点,黑色,圆点,半径为10
      'Call DrawData(PicC_Qc, DataPonit(i), vbRed, "DrawCircle", 10)      '在图片框PicC_Qc中,画D(i)点,黑色,圆点,半径为10
      Call DrawData(PicC_Qc, DataPonit(i), vbMagenta, "DrawCircle", 10)      '在图片框PicC_Qc中,画D(i)点,黑色,圆点,半径为10
   Next i
End Sub


Private Sub CmdAuto_Click()
   Dim CS As Byte
   For CS = 1 To 10
      Call CmdAdjust_Click
      Call CmdInsert1V_Click
   Next CS
End Sub

Private Sub CmdAdjust_Click()              '"调整"按钮
    Dim i As Integer, j As Integer
    Dim CVcmp As Double, ux As Double, uy As Double
    Dim tt As Double, t1 As Double, t0 As Double, Vpoint As xy              'Vpoint为全部点显示
    Dim Kad As Integer
    Dim AdjustNum As Integer
    '
    Dim D1V As Double
    Dim DLS1 As Double
    Dim Vtemp As xy                      '顶点
    Dim tmintemp As Double
    Dim k1 As Double
    Dim m As Integer
    '
    Call SegmentExpression(V, tmin)        '求线段的uxy(),tsx()   V(1)与V(2)间是线段uxy(1)
    Call DataProject(D(), V, uxy, tsx)     '入口:数据点,顶点,各线段单位矢量,各线段投影指标初值
    For AdjustNum = 1 To 50
        DLS1 = DistanceofDtoVSZ            'LS 数据点到曲线的总距离平方
        For Kad = 1 To UBound(V) - 1 Step 1
             DoEvents
             'If Kad = 1 Then tmin = tmin + 0.05
             'If Kad = UBound(V) Then tmax = tmax + 0.05
             
             Call Adjust1Point(Kad, 0.02)
             '
             
             TxtPara(1).Text = Kad: TxtPara(2).Text = PV(Kad)
             If Kad = 1 Then
                TxtPara(3).Text = Pi(Kad + 1) - 1
                Else
                  If Kad = UBound(V) Then
                     TxtPara(3).Text = Pi(Kad - 1) - 1
                  Else
                     TxtPara(3).Text = Pi(Kad) - 1
                  End If
             End If
             TxtPara(4).Text = DairTa(Kad): TxtPara(5).Text = GV(Kad)
             'Text1.Text = DistanceofDtoVSZ
        Next Kad
        V(UBound(V)).X = V(1).X: V(UBound(V)).Y = V(1).Y
        
        '对首尾点特殊调整(开始)---------------
         D1V = DistanceofDtoVSZ
         tmintemp = tmin
         Vtemp.X = V(1).X: Vtemp.Y = V(1).Y
         tmin = tmin + 0.01
         tsx(1) = tmin
         V(1).X = V(2).X - (tsx(2) - tsx(1)) * uxy(1).X
         V(1).Y = V(2).Y - (tsx(2) - tsx(1)) * uxy(1).Y
         Call SegmentExpression(V, tmin)        '求线段的uxy(),tsx()   V(1)与V(2)间是线段uxy(1)
         Call DataProject(D(), V, uxy, tsx)     '入口:数据点,顶点,各线段单位矢量,各线段投影指标初值
         If D1V < DistanceofDtoVSZ Then   '数据点到曲线的总距离平方
              V(1).X = Vtemp.X: V(1).Y = Vtemp.Y
              tmin = tmin + 0.01
              tsx(1) = tmin
              Call SegmentExpression(V, tmin)
              Call DataProject(D(), V, uxy, tsx)
          End If
          
          V(UBound(V)).X = V(1).X: V(UBound(V)).Y = V(1).Y
'         '
'         D1V = DistanceofDtoVSZ
'         m = UBound(V)
'         Vtemp.X = V(m).X: Vtemp.Y = V(m).Y
'         V(m).X = V(m - 1).X + (tsx(m) - tsx(m - 1) - 0.01) * uxy(m - 1).X
'         V(m).Y = V(m - 1).Y + (tsx(m) - tsx(m - 1) - 0.01) * uxy(m - 1).Y
'         Call SegmentExpression(V, tmin)        '求线段的uxy(),tsx()   V(1)与V(2)间是线段uxy(1)
'         Call DataProject(D(), V, uxy, tsx)     '入口:数据点,顶点,各线段单位矢量,各线段投影指标初值
'         If D1V < DistanceofDtoVSZ Then   '数据点到曲线的总距离平方
'              V(m).X = Vtemp.X: V(m).Y = Vtemp.Y
'              Call SegmentExpression(V, tmin)
'              Call DataProject(D(), V, uxy, tsx)
'          End If
'         '对首尾点特殊调整(结束)--------------20070621
             
        
           PicC_Qc.Cls                     '清除原始线段图像
'-----------------画数据点[D().x,D().y]--------------------------------------不显示---------------------------
           'Call DrawDataPoint(D)
'------------------------------------------------------------------------------------------------------------
           
           
'------------------------------------------------Vertices cleaning--------------------------------------------
           
             '用线段的表达方式画图
           For i = 1 To UBound(uxy)
                    Dim VpointPrevious As xy          '存储上一个点的坐标，j-1点的坐标
                    Dim VpointNext As xy              '存储下一个点的坐标，j+1点的坐标
                    Dim K As Integer                  '存储单层全部顶点数目
                    Dim Rflag As Integer              '总循环标志位，1保存，0删除
                    Dim XYflag As Integer             'X,Ysum的标志位，1为正，0为负
                    Dim radiu As Double                   '存储数据半径r
                    Dim Ysum As Double                    '存储y坐标sum
                    Dim Ymean As Double                   '存储y坐标sum的mean
                    Ysum = 0#
                 
                For t0 = tsx(i) To tsx(i + 1) Step 0.002
                    K = (tsx(i + 1) - tsx(i)) / 0.002          '存储单层全部顶点数目
                    Vpoint.Y = V(i).Y + (t1 - tsx(i)) * uxy(i).Y  '导出Vpoint点，进行判断
                    Ysum = Vpoint.Y + Ysum               '求sumY的总和
                Next t0
                
                If Ysum < 0 Then
                    XYflag = 0                           'Ysum的标志位，1为正，0为负
                    Ysum = Abs(Ysum)                     '存储y坐标sum的绝对值
                    Ymean = Ysum / K                     '存储y坐标sum的mean
                Else
                    If Ysum > 0 Then
                    XYflag = 1                           'Ysum的标志位，1为正，0为负
                    Ysum = Abs(Ysum)                     '存储y坐标sum的绝对值
                    Ymean = Ysum / K                     '存储y坐标sum的mean
                    End If
                End If
                
                VpointPrevious.X = 0
                VpointPrevious.Y = 0
                
                For t1 = tsx(i) To tsx(i + 1) Step 0.002
                    Rflag = 0
                    K = (tsx(i + 1) - tsx(i)) / 0.002          '存储单层全部顶点数目
                    Vpoint.X = V(i).X + (t1 - tsx(i)) * uxy(i).X: Vpoint.Y = V(i).Y + (t1 - tsx(i)) * uxy(i).Y  '导出Vpoint点，进行判断
                    
                    If Vpoint.X < 0 And XYflag = 0 Then                             'X负，Y负
                        radiu = (Abs(Ymean) - Abs(Vpoint.X)) * (Abs(Ymean) - Abs(Vpoint.X))
                        If Sqr((Vpoint.X - VpointPrevious.X) * (Vpoint.X - VpointPrevious.X) + (Vpoint.Y - VpointPrevious.Y) * (Vpoint.Y - VpointPrevious.Y)) > radiu Then
                            Rflag = 1
                        End If
                    Else
                        If Vpoint.X > 0 And XYflag = 1 Then                         'X正，Y正
                            radiu = (-Abs(Ymean) + Abs(Vpoint.X)) * (-Abs(Ymean) + Abs(Vpoint.X))
                            If Sqr((Vpoint.X - VpointPrevious.X) * (Vpoint.X - VpointPrevious.X) + (Vpoint.Y - VpointPrevious.Y) * (Vpoint.Y - VpointPrevious.Y)) > radiu Then
                                Rflag = 1
                            End If
                        Else
                            If Vpoint.X < 0 And XYflag = 1 Then                     'X负，Y正
                                radiu = (Abs(Ymean) + Abs(Vpoint.X)) * (Abs(Ymean) + Abs(Vpoint.X))
                                If Sqr((Vpoint.X - VpointPrevious.X) * (Vpoint.X - VpointPrevious.X) + (Vpoint.Y - VpointPrevious.Y) * (Vpoint.Y - VpointPrevious.Y)) > radiu Then
                                    Rflag = 1
                                End If
                            Else
                                If Vpoint.X > 0 And XYflag = 0 Then                 'X正，Y负
                                    radiu = (Abs(Ymean) + Abs(Vpoint.X)) * (Abs(Ymean) + Abs(Vpoint.X))
                                    If Sqr((Vpoint.X - VpointPrevious.X) * (Vpoint.X - VpointPrevious.X) + (Vpoint.Y - VpointPrevious.Y) * (Vpoint.Y - VpointPrevious.Y)) > radiu Then
                                        Rflag = 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                    

                    
                    
                    VpointPrevious.X = Vpoint.X: VpointPrevious.Y = Vpoint.Y      '存储上一个点的坐标
                    
                    If Rflag = 1 Then
                          '线段的表达方式:第i个线段(由第i到第i+1顶点构成)
                        Vpoint.X = V(i).X + (t1 - tsx(i)) * uxy(i).X: Vpoint.Y = V(i).Y + (t1 - tsx(i)) * uxy(i).Y
                        Call DrawData(PicC_Qc, Vpoint, vbBlue, "DrawForkX", 2)       '在图片框中,画点,颜色,形式,半径
                    End If
                    
                Next t1
                 

                
                
           Next i
           
'             '用线段的表达方式画图
'           For i = 1 To UBound(uxy)
'                 For t1 = tsx(i) To tsx(i + 1) Step 0.002
'                      '线段的表达方式:第i个线段(由第i到第i+1顶点构成)
'                    Vpoint.X = V(i).X + (t1 - tsx(i)) * uxy(i).X: Vpoint.Y = V(i).Y + (t1 - tsx(i)) * uxy(i).Y
'                    Call DrawData(PicC_Qc, Vpoint, vbBlue, "DrawForkX", 2)       '在图片框中,画点,颜色,形式,半径
'                 Next t1
'           Next i
           
           
'             '显示插入点的过程（只显示插入点）
'           For i = 1 To UBound(V)
'                V为插入点
'                V.X = V(i).X + (t1 - tsx(i)) * uxy(i).X: Vpoint.Y = V(i).Y + (t1 - tsx(i)) * uxy(i).Y
'                Call DrawData(PicC_Qc, V(i), vbBlue, "DrawForkX", 2)       '在图片框中,画点,颜色,形式,半径
'           Next i
                     
           If Abs(DLS1 - DistanceofDtoVSZ) < 0.001 Then Exit For
    Next AdjustNum
 
    '-----------------------------------------------vertex merging--------------------------------------------
    Dim Vleft As xy, Vright As xy   '定义左点和右点
    Dim Disleft As Double, Disright As Double, Distheshold As Double, Diswhole As Double    '定义左点到中点，右点到中点的临时距离，和距离阈值，总距离
    Dim mL As Integer, mU As Integer   '记录vertex的个数,mL下限，mU上限
    Dim Vxmax As Double, Vxmaxnumber As Integer   '记录vertex中最大的x，和对应顶点编号
    Dim kpara As Double, aupara As Double, alpara As Double, wpara As Double    '定义配置参数
    Dim angle As Double    '定义配置参数
    
    mL = LBound(V)
    mU = UBound(V)
    If mU > 5 Then   '数据点到曲线的总距离平方
        Diswhole = 0
        For i = mL To mU - 1
            Diswhole = Diswhole + Sqr((V(i + 1).X - V(i).X) * (V(i + 1).X - V(i).X) + (V(i + 1).Y - V(i).Y) * (V(i + 1).Y - V(i).Y))
        Next i
        Distheshold = Diswhole / mU                '求出Euclidean distance阈值
        
        For i = mL To mU
            If V(i).X > Vxmax Then
                Vxmax = V(i).X                '记录vertex中最大的x
                Vxmaxnumber = i               '记录vertex中最大x所对应的顶点编号
            End If
        Next i
        
        For i = mL + 1 To Vxmaxnumber - 1        '例如从2到15（放出1，16的空间）
            Disleft = Sqr((V(i - 1).X - V(i).X) * (V(i - 1).X - V(i).X) + (V(i - 1).Y - V(i).Y) * (V(i - 1).Y - V(i).Y))    '左点到中点距离
            Disright = Sqr((V(i + 1).X - V(i).X) * (V(i + 1).X - V(i).X) + (V(i + 1).Y - V(i).Y) * (V(i + 1).Y - V(i).Y))   '右点到中点距离
            If Disleft + Disright < Distheshold * 3 Then
                 '删除点Vi
                 Call UpdateArray(V(), i, V(i).X, V(i).Y)      '更新数组V
                 GoTo ExitLoop              '添加GOTO函数进行跳转，退出循环
            End If
            If Disleft / Disright > aupara Or Disleft / Disright < alpara Then
                 '删除点Vi
                 Call UpdateArray(V(), i, V(i).X, V(i).Y)      '更新数组V
                 GoTo ExitLoop              '添加GOTO函数进行跳转，退出循环
            End If
            
'            tmin = Atn(tmin)                   'vb6只有atn(x)，也就是arctan(x)  'arccos(x)可以用Atn(Sqr(1-x^2)/x)表示。
            Call CalculateAngle(V(i - 1), V(i), V(i + 1), Disleft, Disright, angle)
            If angle > 180 Then                  '角度判断
                 angle = 360 - angle
            End If

            If angle > 180 Then                  '角度判断
                 angle = 360 - angle
            End If
            wpara = 60
            If 180 - angle > wpara Then                '角度判断
                 Call UpdateArray(V(), i, V(i).X, V(i).Y)      '更新数组V
                 GoTo ExitLoop              '添加GOTO函数进行跳转，退出循环
            End If
        Next i
    End If

ExitLoop:                                   'GOTO函数跳转的地方
    
End Sub

Private Sub CmdInsert1V_Click()   '"插入1个顶点"按钮
   Call Insert1V
   Call DrawData(PicC_Qc, V(VnewSerialNumber), vbRed, "DrawCircle", 30)     '在图片框中,画点,颜色,形状,大小
   'Call SegmentExpression(V, tmin)        '求线段的uxy(),tsx()   V(1)与V(2)间是线段uxy(1)
   Call DataProject(D(), V, uxy, tsx)     '入口:数据点,顶点,各线段单位矢量,各线段投影指标初值
End Sub

Private Sub CmdProjectAndSave_Click()    '"投影存盘"按钮
    Dim i As Double, i1 As Double, i2 As Double
    Dim j As Double, j1 As Double, j2 As Double
    Dim t1 As Double
    Dim d1 As Double, d2 As Double
    Dim b1 As Double
    Dim t10 As Double, t11 As Double
    Dim m As Long
    ReDim Tfx(1 To Nd) As Double
  'GoTo AAA
   '(1)计算总的曲线(由线段组成)总长度
   
   d1 = 0#
   For i = LBound(V) To UBound(V) - 1    '对线段循环(开始)
       d1 = d1 + Sqr((V(i + 1).X - V(i).X) ^ 2 + (V(i + 1).Y - V(i).Y) ^ 2)   '计算线段长度
   Next i
   
   j1 = 0
   '(2)计算10000个点的坐标
   For i = LBound(V) To UBound(V) - 1    '对线段循环(开始)
       d2 = Sqr((V(i + 1).X - V(i).X) ^ 2 + (V(i + 1).Y - V(i).Y) ^ 2)   '计算线段长度
       b1 = (d2 / d1) * 10000#                                            '线段内点数
       
       For j = 0 To b1 - 1
         CurcvsPoint(j1).X = V(i).X + (V(i + 1).X - V(i).X) * (j / b1)   '点坐标
         CurcvsPoint(j1).Y = V(i).Y + (V(i + 1).Y - V(i).Y) * (j / b1)   '点坐标
         j1 = j1 + 1
       Next j
   Next i
   '(3)画10000个点(作为检验)
   For j1 = 0 To 9999
    Call DrawData(PicC_Qc, CurcvsPoint(j1), vbRed, "DrawCircle", 1)     '在图片框中,画点,颜色,形状,大小
   Next j1
   '(4)求Tfx(i1)
 
   For i1 = 1 To Nd       '对数据点循环
     j1 = 0
     d1 = (D(i1).X - CurcvsPoint(0).X) ^ 2 + (D(i1).Y - CurcvsPoint(0).Y) ^ 2
     For j2 = 1 To 9999
       d2 = (D(i1).X - CurcvsPoint(j2).X) ^ 2 + (D(i1).Y - CurcvsPoint(j2).Y) ^ 2
       If d2 <= d1 Then d1 = d2: j1 = j2
     Next j2
     Tfx(i1) = j1 / 10000#
   Next i1
'   GoTo BBB
'AAA:
   
   '求数据点在第一主成分线上的投影值
'    For i = 1 To UBound(Tfx)
'        Tfx(i) = D(i).x * ta + D(i).y * tb
'    Next i
  
    '==================================

   
'BBB:
   
   
    '(3)对TFX(I)从小到大排序(相应数据点也变)
        t10 = Tfx(1)
        For i = 1 To UBound(Tfx) - 1 Step 1
          For j = i + 1 To UBound(Tfx) Step 1
             If Tfx(i) > Tfx(j) Then
                t11 = Tfx(i): Tfx(i) = Tfx(j): Tfx(j) = t11
                t11 = D(i).X: D(i).X = D(j).X: D(j).X = t11
                t11 = D(i).Y: D(i).Y = D(j).Y: D(j).Y = t11
            End If
          Next j
        Next i
       
   
  '(5)把数据按比例并平移到[0-1]之间,以便后续运算
   For i = LBound(Tfx) To UBound(Tfx)
      'Tfx(i) = (Tfx(i) + 1) / 2
      D(i).X = (D(i).X + 1) / 2
      D(i).Y = (D(i).Y + 1) / 2
   Next i
   
   
   Call WriteFile(Tfx, D, PCAandTxyFileName)
   Cmdstart.Enabled = False           '"开始"按钮
   CmdAdjust.Enabled = False          '"调整"按钮
   CmdInsert1V.Enabled = False        '"插入1个顶点"按钮
   CmdProjectAndSave.Enabled = False  '"投影存盘"按钮

End Sub


 
Private Sub CmdeExit_Click()          '“退出”按钮
   End
End Sub



Private Sub PicC_Qc_MouseMove(button As Integer, shift As Integer, X As Single, Y As Single)
  'PicC_Qc.ScaleMode = vbPixels '图象按像素分辨（监视器的最小分辨单位）
  Dim i1 As Integer
  Dim t11 As Double
  Dim Vpt As xy
  XTxt.Text = "鼠标指向点(" & Format$(X, "###0") & "," & Format$(Y, "###0") & ") 颜色值="
  XTxt.Text = XTxt.Text & Hex$(PicC_Qc.Point(X, Y))
  If button = 1 Or button = 2 Or button = 4 Then
     If button = 1 Then
       V(1).X = (X - Px0) / Kx: V(1).Y = (Y - Py0) / Ky
       Call Adjust1Point(1, 0.02)
     End If
     If button = 4 Then
        V(2).X = (X - Px0) / Kx: V(2).Y = (Y - Py0) / Ky
        Call Adjust1Point(2, 0.02)
     End If
     If button = 2 Then
       V(3).X = (X - Px0) / Kx: V(3).Y = (Y - Py0) / Ky
        Call Adjust1Point(3, 0.02)
     End If
       Call SegmentExpression(V, tmin)
       Call DataProject(D(), V, uxy, tsx)
       PicC_Qc.Cls
              '画数据点[D().x,D().y]
            Call DrawDataPoint(D)
             '用线段的表达方式画图
             For i1 = 1 To UBound(uxy)
                   For t11 = tsx(i1) To tsx(i1 + 1) Step 0.002
                      '线段的表达方式:第i个线段(由第i到第i+1顶点构成)
                      Vpt.X = V(i1).X + (t11 - tsx(i1)) * uxy(i1).X: Vpt.Y = V(i1).Y + (t11 - tsx(i1)) * uxy(i1).Y
                      Call DrawData(PicC_Qc, Vpt, vbBlue, "DrawForkX", 2)     '在图片框中,画点,颜色,形式,半径
                   Next t11
             Next i1
  End If
  TxtState(1).Text = "button =" & button: TxtState(1).Refresh
  
  
  
End Sub




