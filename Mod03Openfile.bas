Attribute VB_Name = "Mod03Openfile"
'Name:OpenTextFile(�����ļ���)
'Function:(1)���ļ����й�һ��������E()=0
'         (2)������(-1,1)
'���ڲ����������ȫ�ֱ���Nd---���ݵ����, [D().x,D().y]---���ݹ淶��ֵ
'��������Call OpenTextFile("halfCircle.txt") '���ļ�,�ó�ȫ�ֱ���Nd---���ݵ����, [D().x,D().y]---���ݹ淶��ֵ


Public Sub OpenTextFile(FileName As String)   '���ļ�����x() y()��ֵ
    Dim i As Integer, n As Byte
    'Dim xymax As Double, Sumx As Double, Sumy As Double
    Dim s1 As String
    Dim fileline() As String
    Dim SQRxy() As Double
    '
    On Error GoTo OpenTextFileError0        '���ļ��ܵĴ���
    '(1)���ļ��ж����ַ�������fileline��
    Nd = 0
    Open FileName For Input As #1          '����ȷ���ļ���,���ļ�
    Nd = 0                              '�ļ���������ֵ=0
    Do Until EOF(1)
       Nd = Nd + 1
       ReDim Preserve fileline(1 To Nd) '���¶����ַ�������fileline������±�
       Line Input #1, fileline(Nd)      '��һ�С�>������
    Loop
    Close #1                               '�ر��ļ�
    i = 0
    Do
       i = i + 1
    Loop Until (Len(Trim(fileline(i))) <= 2 Or i = Nd)
    If i < Nd Then Nd = i - 1
    If Nd <= 2 Then GoTo OpenTextFileError0
    '���¶���D()����,�����Ǹ�ֵ
    ReDim D(1 To Nd)                    '���¶�������±�
    ReDim SQRxy(1 To Nd)
    '===================================================================
    Sumx = 0#: Sumy = 0#
    For i = 1 To Nd
        fileline(i) = LTrim$(RTrim$(fileline(i)))  '
        n = InStr(fileline(i), " ")
        D(i).X = Left$(fileline(i), n): Sumx = Sumx + D(i).X
        D(i).Y = Right$(fileline(i), Len(fileline(i)) - n): Sumy = Sumy + D(i).Y
    Next i
    'ƽ��ʹE?=0
    Sumx = Sumx / Nd: Sumy = Sumy / Nd
     For i = 1 To Nd
         D(i).X = D(i).X - Sumx
         D(i).Y = D(i).Y - Sumy
         SQRxy(i) = Sqr(D(i).X * D(i).X + D(i).Y * (D(i).Y))
     Next i
     
    '�����ֵ(����ֵ)
    'xymax = SQRxy(1)
    'For i = 1 To Nd - 1
    '    If xymax < SQRxy(i + 1) Then xymax = SQRxy(i + 1)
    'Next i
    '����������ֵ(������(-1,1)֮��
    'For i = 1 To Nd: D(i).X = D(i).X / xymax: D(i).Y = D(i).Y / xymax: Next i
    
    xymax = SQRxy(1)
    For i = 1 To Nd - 1
        If xymax < SQRxy(i + 1) Then xymax = SQRxy(i + 1)
    Next i
    '����������ֵ(������(-1,1)֮��
    For i = 1 To Nd:
        D(i).X = (D(i).X) / xymax
        D(i).Y = (D(i).Y) / xymax
    Next i
    
    
    
    
    
    '----------------------------------------------------------------------------
    
    'Sw = Format$(X1, "0000.000000") & " " & Format$(Y1, "0000.000000")
    
    FrmPC.TxtState(1).Text = "Sumxp=" & Format$(Sumx, "000.00000") & "   " & "Sumyp=" & Format$(Sumy, "000.00000") & "   " & "xymax=" & Format$(xymax, "000.00000")
    'FrmPC.TxtState(1).Text = "Sumxp=" & Sumx & "Sumyp=" & Sumy & "xymax=" & xymax
    '-----------------------------------------------------------------------------
    GoTo OpenTextFileExit
OpenTextFileError0:            '���ļ��ܵĴ���
    MsgBox ("���ļ�ģ��:���ļ��ܵĴ���")
OpenTextFileExit:
End Sub

