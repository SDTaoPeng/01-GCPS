Attribute VB_Name = "Mod08"
'-----------------------------------------------------------------------------------------------
'�ӳ�������WriteFile(t() As Double, xy() As Double,FileName As String)
'���ܣ�����t���顢��xy���鰴��д������FileName�ļ�FileName��
'      (��tΪһά����.xyΪ2Ϊ����(��1�±���t�����Ӧ,��1�±�Ϊ1 to 2)
'      ת���ַ���ʽ����
'������:
'       ReDim tfx(1 To 2) As Double
'       ReDim fxy(1 To 2, 1 To 2) As Double
'       tfx(1) = -0.111
'       tfx(2) = 1.22111
'       fxy(1, 1) = 0.11222: fxy(1, 2) = 0.113333333333
'       fxy(2, 1) = 0.33222: fxy(2, 2) = 0.223333333333
'       Call WriteFile(tfx, fxy,"LS.TXT")    д��"LS.TXT"�ļ�
'
Public Sub WriteFile(t() As Double, data() As xy, FileName As String)
    '�������ӳ���ʹ�õ���ʱ����
    Dim Iw As Integer
    Dim Sw As String
    '���ļ���д��(��ԭ������,�ᱻ���)
    Open FileName For Output As #1      '���ļ�
    '����д��
    For Iw = LBound(t) To UBound(t)
        Sw = Format$(t(Iw), "0.0000000000") & " "
        
        If data(Iw).X >= 0 Then
           Sw = Sw & "+" & Format$(data(Iw).X, "0.0000000000") & " "
        Else
           Sw = Sw & Format$(data(Iw).X, "#0.0000000000") & " "
        End If
        
        If data(Iw).Y >= 0 Then
           Sw = Sw & "+" & Format$(data(Iw).Y, "0.0000000000")
        Else
           Sw = Sw & Format$(data(Iw).Y, "#0.0000000000")
        End If
        
        Print #1, Sw             'д��1��(Sw)
    Next Iw
    'д��xymax,Sumx,Sumy,
        Sw = Format$(xymax, "0.0000000000") & " "
        
        If Sumx >= 0 Then
           Sw = Sw & "+" & Format$(Sumx, "0.0000000000") & " "
        Else
           Sw = Sw & Format$(Sumx, "#0.0000000000") & " "
        End If
        
        If Sumy >= 0 Then
           Sw = Sw & "+" & Format$(Sumy, "0.0000000000") & " "
        Else
           Sw = Sw & Format$(Sumy, "#0.0000000000") & " "
        End If
    Print #1, Sw              'д��1��(Sw)
    '�ر��ļ�
    Close #1
End Sub
