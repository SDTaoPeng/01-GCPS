VERSION 5.00
Begin VB.Form FrmPC 
   BackColor       =   &H00E0E0E0&
   Caption         =   "�պ������߳���"
   ClientHeight    =   10350
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   16320
   Icon            =   "FrmPC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10350
   ScaleWidth      =   16320
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox TxtState 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "����"
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
      Text            =   "�����̵ó�[t,(x,y)]�Ĺ�ϵ,��Ϊ�мලBNNM��ѧϰ����(����)"
      Top             =   9240
      Width           =   6855
   End
   Begin VB.TextBox TxtState 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "����"
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
      Text            =   "���ýǶȳͷ�������պ�������(�÷ֶ�ֱ�߱��,������ݵ��ͶӰָ�����)"
      Top             =   150
      Width           =   8385
   End
   Begin VB.Frame Frame2 
      Caption         =   "������ʾ"
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
         Text            =   "��ʾ����"
         Top             =   1500
         Width           =   3705
      End
      Begin VB.TextBox TxtPara 
         Height          =   270
         Index           =   4
         Left            =   840
         TabIndex        =   13
         Text            =   "��ʾ����"
         Top             =   1200
         Width           =   3705
      End
      Begin VB.TextBox TxtPara 
         Height          =   270
         Index           =   3
         Left            =   870
         TabIndex        =   11
         Text            =   "��ʾ����"
         Top             =   540
         Width           =   3705
      End
      Begin VB.TextBox TxtPara 
         Height          =   270
         Index           =   2
         Left            =   870
         TabIndex        =   9
         Text            =   "��ʾ����"
         Top             =   900
         Width           =   3705
      End
      Begin VB.TextBox TxtPara 
         Height          =   270
         Index           =   1
         Left            =   870
         TabIndex        =   6
         Text            =   "��ʾ����"
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
         Caption         =   "�ܳͷ�"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1530
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "����ͷ�"
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   10
         Top             =   1170
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "�Ƕȳͷ�"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   8
         Top             =   870
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "�������"
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
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "�Զ�"
         Height          =   495
         Left            =   3630
         TabIndex        =   28
         Top             =   2970
         Width           =   765
      End
      Begin VB.Frame Frame5 
         Caption         =   "[t,(x,y)]����·�����ļ���"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "ѡ�������ļ�(SRCdataĿ¼��)"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "ͶӰ����"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����1������"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�˳�"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��ʼ"
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   "ͼ����ʾ��"
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
Option Explicit                       '���δ�������ı���



Private Sub Form_Load()          '�����ʼ��ʱִ��

    File1.Path = "..\SRCdata"       '�����ļ�·��

    '-----------------------------------------------------
    Dim Iw As Double
    Dim Sw As String
    Dim X As Double, Y As Double
'    Open "..\SRCdata\Lizi2_YWC.TXT" For Output As #1       '���ļ�
'    '����д��
'    For Iw = 0 To 6.283 Step 0.02          '0.2
'        X = Sin(Iw)  ',- Cos(0.8 * Iw) ^ 2 + 0.2 * Rnd(1)
'        Y = Cos(Iw)  '- Sin(2 * Iw) ^ 2 + 0.2 * Rnd(1)
'        Sw = Format$(X, "0.0000000000") & " " & Format$(Y, "0.0000000000")
'        Print #1, Sw             'д��1��(Sw)
'    Next Iw
'    '�ر��ļ�
'    Close #1
'    '-----------------------------------------------------
  
  
    'PCAandTxyFileName = "..\Tdata\T_SmallNoise_S.txt"
    Cmdstart.Enabled = False           '"��ʼ"��ť
    CmdAdjust.Enabled = False          '"����"��ť
    CmdInsert1V.Enabled = False        '"����1������"��ť
    CmdProjectAndSave.Enabled = False  '"ͶӰ����"��ť
End Sub


Private Sub File1_Click()        'ѡ�������ļ�
    '----------------------------------ִ�е������ļ�----------------------------------
    DataFileName = File1.Path & "\" & File1.FileName
    TxtFile.Text = "�Ѿ�ѡ��������ļ�Ϊ:" & "   ..\SRCdata\" & File1.FileName
    
    '----------------------------------����������ļ�-------------------------------------
    PCAandTxyFileName = "..\Tdata\T_" & File1.FileName
    SaveFileName.Text = "[t,(x,y)]Ŀ���ļ���:" & "   ..\Tdata\T_" & File1.FileName
    
    Cmdstart.Enabled = True           '"��ʼ"��ť
    CmdAdjust.Enabled = False          '"����"��ť
    CmdInsert1V.Enabled = False        '"����1������"��ť
    CmdProjectAndSave.Enabled = False  '"ͶӰ����"��ť
End Sub



'-------------------------------------------------------------------
Private Sub Cmdstart_ClIck()           '����ʼ����ť
   '������ʱ����
   Dim i As Integer, j As Integer
   Dim Vpoint As xy, fPCA As xy
   Dim t1 As Double
   
   '��ͼƬ���л�������,�̶�,�߿�
   ' -----------------------Normalization �ѱ���������(-1,-1)-(1,1)֮��,����Ϊԭ��-----------------------
   Call Drawcoordinate(PicC_Qc, vbWhite, vbWhite, vbWhite)  ''��ͼƬ���л�������,��������ɫ,�̶���ɫ,�߿���ɫ
   '���ļ�,�ó�ȫ�ֱ���Nd---���ݵ����, [D().x,D().y]---���ݹ淶��ֵ
   Call OpenTextFile(DataFileName) '���ļ�,�ó�ȫ�ֱ���Nd---���ݵ����, [D().x,D().y]---���ݹ淶��ֵ
 
   
   '�����ݵ�[D().x,D().y]--------------------------------------����ʾ-------------------------------------
    'Call DrawDataPoint(D)
   '-----------------------------------------------------------------------------------------------------
   
'   '��һ���ɷ�
        'Call PCARTB(D)               '���:(ȫ�ֱ���) ���ݵ����Nd,���ݵ㼯��[D().x,D().y]
                                     '���ڲ�������һ���ɷ�=(ta*t,tb*t)��ta,tb(ȫ�ֱ���)
        'Call JYSAstep1  '�����㷨step1-���ݹ淶��ֵD():��һ���ɷֵ�ta,tb(ȫ�ֱ���)->����V(1) V(2) V(3) ͶӰָ��tmin tmax
        'Call InsertV(3)
        '�ʵ����̵�һ���ɷ��߶�
        'tmin = tmin * 3 / 5: tmax = tmax * 3 / 5
        'V(1).X = tmin * ta: V(1).Y = tmin * tb
        'V(2)����
        'V(3).X = tmax * ta: V(3).Y = tmax * tb
        '����һ���ɷ�

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
       Call DrawData(PicC_Qc, V(i), vbRed, "DrawCircle", 10)     '��ͼƬ����,����,��ɫ,��״,��С
   Next i
    Call SegmentExpression(V, tmin)   '���߶ε�uxy(),tsx()   V(1)��V(2)�����߶�uxy(1)
     '���߶εı�﷽ʽ��ͼ
     For i = 1 To UBound(uxy)
          For t1 = tsx(i) To tsx(i + 1) Step 0.002
             '�߶εı�﷽ʽ:��i���߶�(�ɵ�i����i+1���㹹��)
             Vpoint.X = V(i).X + (t1 - tsx(i)) * uxy(i).X: Vpoint.Y = V(i).Y + (t1 - tsx(i)) * uxy(i).Y
             Call DrawData(PicC_Qc, Vpoint, vbRed, "DrawForkX", 2)     '��ͼƬ��PicC_Qc��,��V��,��ɫ,Բ��,�뾶Ϊ1
          Next t1
    Next i
    Cmdstart.Enabled = False           '"��ʼ"��ť
    CmdAdjust.Enabled = True           '"����"��ť
    CmdInsert1V.Enabled = True         '"����1������"��ť
    CmdProjectAndSave.Enabled = True   '"ͶӰ����"��ť
End Sub
'�����ݵ�[D().x,D().y]
Private Sub DrawDataPoint(ByRef DataPonit() As xy)
   Dim i As Integer
   For i = 1 To UBound(D)                     '���������ݵ�ѭ���������ݵ�)
      'Call DrawData(PicC_Qc, DataPonit(i), vbBlack, "DrawCircle", 10)      '��ͼƬ��PicC_Qc��,��D(i)��,��ɫ,Բ��,�뾶Ϊ10
      'Call DrawData(PicC_Qc, DataPonit(i), vbRed, "DrawCircle", 10)      '��ͼƬ��PicC_Qc��,��D(i)��,��ɫ,Բ��,�뾶Ϊ10
      Call DrawData(PicC_Qc, DataPonit(i), vbMagenta, "DrawCircle", 10)      '��ͼƬ��PicC_Qc��,��D(i)��,��ɫ,Բ��,�뾶Ϊ10
   Next i
End Sub


Private Sub CmdAuto_Click()
   Dim CS As Byte
   For CS = 1 To 10
      Call CmdAdjust_Click
      Call CmdInsert1V_Click
   Next CS
End Sub

Private Sub CmdAdjust_Click()              '"����"��ť
    Dim i As Integer, j As Integer
    Dim CVcmp As Double, ux As Double, uy As Double
    Dim tt As Double, t1 As Double, t0 As Double, Vpoint As xy              'VpointΪȫ������ʾ
    Dim Kad As Integer
    Dim AdjustNum As Integer
    '
    Dim D1V As Double
    Dim DLS1 As Double
    Dim Vtemp As xy                      '����
    Dim tmintemp As Double
    Dim k1 As Double
    Dim m As Integer
    '
    Call SegmentExpression(V, tmin)        '���߶ε�uxy(),tsx()   V(1)��V(2)�����߶�uxy(1)
    Call DataProject(D(), V, uxy, tsx)     '���:���ݵ�,����,���߶ε�λʸ��,���߶�ͶӰָ���ֵ
    For AdjustNum = 1 To 50
        DLS1 = DistanceofDtoVSZ            'LS ���ݵ㵽���ߵ��ܾ���ƽ��
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
        
        '����β���������(��ʼ)---------------
         D1V = DistanceofDtoVSZ
         tmintemp = tmin
         Vtemp.X = V(1).X: Vtemp.Y = V(1).Y
         tmin = tmin + 0.01
         tsx(1) = tmin
         V(1).X = V(2).X - (tsx(2) - tsx(1)) * uxy(1).X
         V(1).Y = V(2).Y - (tsx(2) - tsx(1)) * uxy(1).Y
         Call SegmentExpression(V, tmin)        '���߶ε�uxy(),tsx()   V(1)��V(2)�����߶�uxy(1)
         Call DataProject(D(), V, uxy, tsx)     '���:���ݵ�,����,���߶ε�λʸ��,���߶�ͶӰָ���ֵ
         If D1V < DistanceofDtoVSZ Then   '���ݵ㵽���ߵ��ܾ���ƽ��
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
'         Call SegmentExpression(V, tmin)        '���߶ε�uxy(),tsx()   V(1)��V(2)�����߶�uxy(1)
'         Call DataProject(D(), V, uxy, tsx)     '���:���ݵ�,����,���߶ε�λʸ��,���߶�ͶӰָ���ֵ
'         If D1V < DistanceofDtoVSZ Then   '���ݵ㵽���ߵ��ܾ���ƽ��
'              V(m).X = Vtemp.X: V(m).Y = Vtemp.Y
'              Call SegmentExpression(V, tmin)
'              Call DataProject(D(), V, uxy, tsx)
'          End If
'         '����β���������(����)--------------20070621
             
        
           PicC_Qc.Cls                     '���ԭʼ�߶�ͼ��
'-----------------�����ݵ�[D().x,D().y]--------------------------------------����ʾ---------------------------
           'Call DrawDataPoint(D)
'------------------------------------------------------------------------------------------------------------
           
           
'------------------------------------------------Vertices cleaning--------------------------------------------
           
             '���߶εı�﷽ʽ��ͼ
           For i = 1 To UBound(uxy)
                    Dim VpointPrevious As xy          '�洢��һ��������꣬j-1�������
                    Dim VpointNext As xy              '�洢��һ��������꣬j+1�������
                    Dim K As Integer                  '�洢����ȫ��������Ŀ
                    Dim Rflag As Integer              '��ѭ����־λ��1���棬0ɾ��
                    Dim XYflag As Integer             'X,Ysum�ı�־λ��1Ϊ����0Ϊ��
                    Dim radiu As Double                   '�洢���ݰ뾶r
                    Dim Ysum As Double                    '�洢y����sum
                    Dim Ymean As Double                   '�洢y����sum��mean
                    Ysum = 0#
                 
                For t0 = tsx(i) To tsx(i + 1) Step 0.002
                    K = (tsx(i + 1) - tsx(i)) / 0.002          '�洢����ȫ��������Ŀ
                    Vpoint.Y = V(i).Y + (t1 - tsx(i)) * uxy(i).Y  '����Vpoint�㣬�����ж�
                    Ysum = Vpoint.Y + Ysum               '��sumY���ܺ�
                Next t0
                
                If Ysum < 0 Then
                    XYflag = 0                           'Ysum�ı�־λ��1Ϊ����0Ϊ��
                    Ysum = Abs(Ysum)                     '�洢y����sum�ľ���ֵ
                    Ymean = Ysum / K                     '�洢y����sum��mean
                Else
                    If Ysum > 0 Then
                    XYflag = 1                           'Ysum�ı�־λ��1Ϊ����0Ϊ��
                    Ysum = Abs(Ysum)                     '�洢y����sum�ľ���ֵ
                    Ymean = Ysum / K                     '�洢y����sum��mean
                    End If
                End If
                
                VpointPrevious.X = 0
                VpointPrevious.Y = 0
                
                For t1 = tsx(i) To tsx(i + 1) Step 0.002
                    Rflag = 0
                    K = (tsx(i + 1) - tsx(i)) / 0.002          '�洢����ȫ��������Ŀ
                    Vpoint.X = V(i).X + (t1 - tsx(i)) * uxy(i).X: Vpoint.Y = V(i).Y + (t1 - tsx(i)) * uxy(i).Y  '����Vpoint�㣬�����ж�
                    
                    If Vpoint.X < 0 And XYflag = 0 Then                             'X����Y��
                        radiu = (Abs(Ymean) - Abs(Vpoint.X)) * (Abs(Ymean) - Abs(Vpoint.X))
                        If Sqr((Vpoint.X - VpointPrevious.X) * (Vpoint.X - VpointPrevious.X) + (Vpoint.Y - VpointPrevious.Y) * (Vpoint.Y - VpointPrevious.Y)) > radiu Then
                            Rflag = 1
                        End If
                    Else
                        If Vpoint.X > 0 And XYflag = 1 Then                         'X����Y��
                            radiu = (-Abs(Ymean) + Abs(Vpoint.X)) * (-Abs(Ymean) + Abs(Vpoint.X))
                            If Sqr((Vpoint.X - VpointPrevious.X) * (Vpoint.X - VpointPrevious.X) + (Vpoint.Y - VpointPrevious.Y) * (Vpoint.Y - VpointPrevious.Y)) > radiu Then
                                Rflag = 1
                            End If
                        Else
                            If Vpoint.X < 0 And XYflag = 1 Then                     'X����Y��
                                radiu = (Abs(Ymean) + Abs(Vpoint.X)) * (Abs(Ymean) + Abs(Vpoint.X))
                                If Sqr((Vpoint.X - VpointPrevious.X) * (Vpoint.X - VpointPrevious.X) + (Vpoint.Y - VpointPrevious.Y) * (Vpoint.Y - VpointPrevious.Y)) > radiu Then
                                    Rflag = 1
                                End If
                            Else
                                If Vpoint.X > 0 And XYflag = 0 Then                 'X����Y��
                                    radiu = (Abs(Ymean) + Abs(Vpoint.X)) * (Abs(Ymean) + Abs(Vpoint.X))
                                    If Sqr((Vpoint.X - VpointPrevious.X) * (Vpoint.X - VpointPrevious.X) + (Vpoint.Y - VpointPrevious.Y) * (Vpoint.Y - VpointPrevious.Y)) > radiu Then
                                        Rflag = 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                    

                    
                    
                    VpointPrevious.X = Vpoint.X: VpointPrevious.Y = Vpoint.Y      '�洢��һ���������
                    
                    If Rflag = 1 Then
                          '�߶εı�﷽ʽ:��i���߶�(�ɵ�i����i+1���㹹��)
                        Vpoint.X = V(i).X + (t1 - tsx(i)) * uxy(i).X: Vpoint.Y = V(i).Y + (t1 - tsx(i)) * uxy(i).Y
                        Call DrawData(PicC_Qc, Vpoint, vbBlue, "DrawForkX", 2)       '��ͼƬ����,����,��ɫ,��ʽ,�뾶
                    End If
                    
                Next t1
                 

                
                
           Next i
           
'             '���߶εı�﷽ʽ��ͼ
'           For i = 1 To UBound(uxy)
'                 For t1 = tsx(i) To tsx(i + 1) Step 0.002
'                      '�߶εı�﷽ʽ:��i���߶�(�ɵ�i����i+1���㹹��)
'                    Vpoint.X = V(i).X + (t1 - tsx(i)) * uxy(i).X: Vpoint.Y = V(i).Y + (t1 - tsx(i)) * uxy(i).Y
'                    Call DrawData(PicC_Qc, Vpoint, vbBlue, "DrawForkX", 2)       '��ͼƬ����,����,��ɫ,��ʽ,�뾶
'                 Next t1
'           Next i
           
           
'             '��ʾ�����Ĺ��̣�ֻ��ʾ����㣩
'           For i = 1 To UBound(V)
'                VΪ�����
'                V.X = V(i).X + (t1 - tsx(i)) * uxy(i).X: Vpoint.Y = V(i).Y + (t1 - tsx(i)) * uxy(i).Y
'                Call DrawData(PicC_Qc, V(i), vbBlue, "DrawForkX", 2)       '��ͼƬ����,����,��ɫ,��ʽ,�뾶
'           Next i
                     
           If Abs(DLS1 - DistanceofDtoVSZ) < 0.001 Then Exit For
    Next AdjustNum
 
    '-----------------------------------------------vertex merging--------------------------------------------
    Dim Vleft As xy, Vright As xy   '���������ҵ�
    Dim Disleft As Double, Disright As Double, Distheshold As Double, Diswhole As Double    '������㵽�е㣬�ҵ㵽�е����ʱ���룬�;�����ֵ���ܾ���
    Dim mL As Integer, mU As Integer   '��¼vertex�ĸ���,mL���ޣ�mU����
    Dim Vxmax As Double, Vxmaxnumber As Integer   '��¼vertex������x���Ͷ�Ӧ������
    Dim kpara As Double, aupara As Double, alpara As Double, wpara As Double    '�������ò���
    Dim angle As Double    '�������ò���
    
    mL = LBound(V)
    mU = UBound(V)
    If mU > 5 Then   '���ݵ㵽���ߵ��ܾ���ƽ��
        Diswhole = 0
        For i = mL To mU - 1
            Diswhole = Diswhole + Sqr((V(i + 1).X - V(i).X) * (V(i + 1).X - V(i).X) + (V(i + 1).Y - V(i).Y) * (V(i + 1).Y - V(i).Y))
        Next i
        Distheshold = Diswhole / mU                '���Euclidean distance��ֵ
        
        For i = mL To mU
            If V(i).X > Vxmax Then
                Vxmax = V(i).X                '��¼vertex������x
                Vxmaxnumber = i               '��¼vertex�����x����Ӧ�Ķ�����
            End If
        Next i
        
        For i = mL + 1 To Vxmaxnumber - 1        '�����2��15���ų�1��16�Ŀռ䣩
            Disleft = Sqr((V(i - 1).X - V(i).X) * (V(i - 1).X - V(i).X) + (V(i - 1).Y - V(i).Y) * (V(i - 1).Y - V(i).Y))    '��㵽�е����
            Disright = Sqr((V(i + 1).X - V(i).X) * (V(i + 1).X - V(i).X) + (V(i + 1).Y - V(i).Y) * (V(i + 1).Y - V(i).Y))   '�ҵ㵽�е����
            If Disleft + Disright < Distheshold * 3 Then
                 'ɾ����Vi
                 Call UpdateArray(V(), i, V(i).X, V(i).Y)      '��������V
                 GoTo ExitLoop              '���GOTO����������ת���˳�ѭ��
            End If
            If Disleft / Disright > aupara Or Disleft / Disright < alpara Then
                 'ɾ����Vi
                 Call UpdateArray(V(), i, V(i).X, V(i).Y)      '��������V
                 GoTo ExitLoop              '���GOTO����������ת���˳�ѭ��
            End If
            
'            tmin = Atn(tmin)                   'vb6ֻ��atn(x)��Ҳ����arctan(x)  'arccos(x)������Atn(Sqr(1-x^2)/x)��ʾ��
            Call CalculateAngle(V(i - 1), V(i), V(i + 1), Disleft, Disright, angle)
            If angle > 180 Then                  '�Ƕ��ж�
                 angle = 360 - angle
            End If

            If angle > 180 Then                  '�Ƕ��ж�
                 angle = 360 - angle
            End If
            wpara = 60
            If 180 - angle > wpara Then                '�Ƕ��ж�
                 Call UpdateArray(V(), i, V(i).X, V(i).Y)      '��������V
                 GoTo ExitLoop              '���GOTO����������ת���˳�ѭ��
            End If
        Next i
    End If

ExitLoop:                                   'GOTO������ת�ĵط�
    
End Sub

Private Sub CmdInsert1V_Click()   '"����1������"��ť
   Call Insert1V
   Call DrawData(PicC_Qc, V(VnewSerialNumber), vbRed, "DrawCircle", 30)     '��ͼƬ����,����,��ɫ,��״,��С
   'Call SegmentExpression(V, tmin)        '���߶ε�uxy(),tsx()   V(1)��V(2)�����߶�uxy(1)
   Call DataProject(D(), V, uxy, tsx)     '���:���ݵ�,����,���߶ε�λʸ��,���߶�ͶӰָ���ֵ
End Sub

Private Sub CmdProjectAndSave_Click()    '"ͶӰ����"��ť
    Dim i As Double, i1 As Double, i2 As Double
    Dim j As Double, j1 As Double, j2 As Double
    Dim t1 As Double
    Dim d1 As Double, d2 As Double
    Dim b1 As Double
    Dim t10 As Double, t11 As Double
    Dim m As Long
    ReDim Tfx(1 To Nd) As Double
  'GoTo AAA
   '(1)�����ܵ�����(���߶����)�ܳ���
   
   d1 = 0#
   For i = LBound(V) To UBound(V) - 1    '���߶�ѭ��(��ʼ)
       d1 = d1 + Sqr((V(i + 1).X - V(i).X) ^ 2 + (V(i + 1).Y - V(i).Y) ^ 2)   '�����߶γ���
   Next i
   
   j1 = 0
   '(2)����10000���������
   For i = LBound(V) To UBound(V) - 1    '���߶�ѭ��(��ʼ)
       d2 = Sqr((V(i + 1).X - V(i).X) ^ 2 + (V(i + 1).Y - V(i).Y) ^ 2)   '�����߶γ���
       b1 = (d2 / d1) * 10000#                                            '�߶��ڵ���
       
       For j = 0 To b1 - 1
         CurcvsPoint(j1).X = V(i).X + (V(i + 1).X - V(i).X) * (j / b1)   '������
         CurcvsPoint(j1).Y = V(i).Y + (V(i + 1).Y - V(i).Y) * (j / b1)   '������
         j1 = j1 + 1
       Next j
   Next i
   '(3)��10000����(��Ϊ����)
   For j1 = 0 To 9999
    Call DrawData(PicC_Qc, CurcvsPoint(j1), vbRed, "DrawCircle", 1)     '��ͼƬ����,����,��ɫ,��״,��С
   Next j1
   '(4)��Tfx(i1)
 
   For i1 = 1 To Nd       '�����ݵ�ѭ��
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
   
   '�����ݵ��ڵ�һ���ɷ����ϵ�ͶӰֵ
'    For i = 1 To UBound(Tfx)
'        Tfx(i) = D(i).x * ta + D(i).y * tb
'    Next i
  
    '==================================

   
'BBB:
   
   
    '(3)��TFX(I)��С��������(��Ӧ���ݵ�Ҳ��)
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
       
   
  '(5)�����ݰ�������ƽ�Ƶ�[0-1]֮��,�Ա��������
   For i = LBound(Tfx) To UBound(Tfx)
      'Tfx(i) = (Tfx(i) + 1) / 2
      D(i).X = (D(i).X + 1) / 2
      D(i).Y = (D(i).Y + 1) / 2
   Next i
   
   
   Call WriteFile(Tfx, D, PCAandTxyFileName)
   Cmdstart.Enabled = False           '"��ʼ"��ť
   CmdAdjust.Enabled = False          '"����"��ť
   CmdInsert1V.Enabled = False        '"����1������"��ť
   CmdProjectAndSave.Enabled = False  '"ͶӰ����"��ť

End Sub


 
Private Sub CmdeExit_Click()          '���˳�����ť
   End
End Sub



Private Sub PicC_Qc_MouseMove(button As Integer, shift As Integer, X As Single, Y As Single)
  'PicC_Qc.ScaleMode = vbPixels 'ͼ�����طֱ棨����������С�ֱ浥λ��
  Dim i1 As Integer
  Dim t11 As Double
  Dim Vpt As xy
  XTxt.Text = "���ָ���(" & Format$(X, "###0") & "," & Format$(Y, "###0") & ") ��ɫֵ="
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
              '�����ݵ�[D().x,D().y]
            Call DrawDataPoint(D)
             '���߶εı�﷽ʽ��ͼ
             For i1 = 1 To UBound(uxy)
                   For t11 = tsx(i1) To tsx(i1 + 1) Step 0.002
                      '�߶εı�﷽ʽ:��i���߶�(�ɵ�i����i+1���㹹��)
                      Vpt.X = V(i1).X + (t11 - tsx(i1)) * uxy(i1).X: Vpt.Y = V(i1).Y + (t11 - tsx(i1)) * uxy(i1).Y
                      Call DrawData(PicC_Qc, Vpt, vbBlue, "DrawForkX", 2)     '��ͼƬ����,����,��ɫ,��ʽ,�뾶
                   Next t11
             Next i1
  End If
  TxtState(1).Text = "button =" & button: TxtState(1).Refresh
  
  
  
End Sub




