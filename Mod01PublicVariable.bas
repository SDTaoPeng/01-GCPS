Attribute VB_Name = "Mod01PublicVariable"
'��ͼʱʹ�õı�����Drawcoordinateģ�����ó���
Public Px0 As Integer, Py0 As Integer '(0��λ��)
Public Kx As Double, Ky As Double     '(��λ����������Ļ���ȶ�����λ)



Public DataFileName As String
Public PCAandTxyFileName As String



'ͶӰ����ʹ�õ�����
'ƽ�����ݵ����Ͷ���
Public Type xy
     X As Double
     Y As Double
End Type
Public CurcvsPoint(0 To 9999) As xy    '���ߵ�
'ȫ�ֱ�������
Public Tfx() As Double    'ͶӰָ��   (����,�Ը����ݵ�ΪԪ��)
Public Fxy() As Double    '���ݵ���� (��1�±�Ϊ���ݵ����,��2�±�Ϊ����ά��)
                          'ÿ������Ϊ������(Fxy(i,1),Fxy(i,2))
                          
Public xymax As Double, Sumx As Double, Sumy As Double

'���ļ��ж�����
Public Nd As Integer            'ԭʼ���ݸ���
Public D() As xy                '���ݵ�(D().x,D().y)
'PCA-RTB
Public t() As Double               'ͶӰָ��(Projection Index)
Public ta As Double, tb As Double  'PCA-RTB����
'
Public V() As xy                      '����
Public tmax As Double, tmin As Double 'ͶӰָ����С�����ֵ
'
Public uxy() As xy                       'uxy()���߶ε�λʸ������
Public tsx() As Double                   'tsx()���߶�ͶӰָ���ֵ����
Public DataNumberofSegment() As Integer  '�����ڸ��߶ε����ݸ���

'
Public DtoVS() As Integer          '���ݵ�ͶӰ��ʶ 1-20000 Ϊ���ڶ��� 20000����Ϊ�����߶�
Public DistanceofDtoVS() As Double '���ݵ㵽ͶӰ���ľ���ƽ��
'�����Ż�������
Public Cgm() As Double             '�������߶����ݵ����߶�Si�ľ���ƽ��
Public VV() As Double              '�����ڶ�������ݵ��ö���Vi�ľ���ƽ��
Public u2() As Double              '���߶γ���ƽ��

Public Pi() As Double              '����ĽǶȳͷ�
Public PV() As Double              '����ĽǶȳͷ��ܺ�
Public DairTa() As Double          '����ľ���Լ���ܺ�
Public GV() As Double              '����ľ���Լ��+�Ƕȳͷ� �ܺ�
Public DistanceofDtoVSZ As Double  '���ݵ㵽���ߵ��ܾ���ƽ��

Public MoveDirectionDistance(1 To 8) As Double  '���ݵ㵽���ߵ��ܾ���ƽ������ǰ��
Public MoveDirectionV(1 To 8) As xy             '8���¶���


'
Public VnewSerialNumber  As Integer




   


