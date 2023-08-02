Attribute VB_Name = "Mod03Openfile"
'Name:OpenTextFile(数据文件名)
'Function:(1)打开文件进行归一化处理且E()=0
'         (2)限制在(-1,1)
'出口参数：计算出全局变量Nd---数据点个数, [D().x,D().y]---数据规范点值
'调用例：Call OpenTextFile("halfCircle.txt") '打开文件,得出全局变量Nd---数据点个数, [D().x,D().y]---数据规范点值


Public Sub OpenTextFile(FileName As String)   '打开文件并给x() y()赋值
    Dim i As Integer, n As Byte
    'Dim xymax As Double, Sumx As Double, Sumy As Double
    Dim s1 As String
    Dim fileline() As String
    Dim SQRxy() As Double
    '
    On Error GoTo OpenTextFileError0        '打开文件总的错误
    '(1)从文件中读到字符串数组fileline中
    Nd = 0
    Open FileName For Input As #1          '有正确的文件名,打开文件
    Nd = 0                              '文件总行数初值=0
    Do Until EOF(1)
       Nd = Nd + 1
       ReDim Preserve fileline(1 To Nd) '重新定义字符串数组fileline的最大下标
       Line Input #1, fileline(Nd)      '读一行—>最新行
    Loop
    Close #1                               '关闭文件
    i = 0
    Do
       i = i + 1
    Loop Until (Len(Trim(fileline(i))) <= 2 Or i = Nd)
    If i < Nd Then Nd = i - 1
    If Nd <= 2 Then GoTo OpenTextFileError0
    '重新定义D()数组,给它们赋值
    ReDim D(1 To Nd)                    '重新定义最大下标
    ReDim SQRxy(1 To Nd)
    '===================================================================
    Sumx = 0#: Sumy = 0#
    For i = 1 To Nd
        fileline(i) = LTrim$(RTrim$(fileline(i)))  '
        n = InStr(fileline(i), " ")
        D(i).X = Left$(fileline(i), n): Sumx = Sumx + D(i).X
        D(i).Y = Right$(fileline(i), Len(fileline(i)) - n): Sumy = Sumy + D(i).Y
    Next i
    '平移使E?=0
    Sumx = Sumx / Nd: Sumy = Sumy / Nd
     For i = 1 To Nd
         D(i).X = D(i).X - Sumx
         D(i).Y = D(i).Y - Sumy
         SQRxy(i) = Sqr(D(i).X * D(i).X + D(i).Y * (D(i).Y))
     Next i
     
    '求最大值(绝对值)
    'xymax = SQRxy(1)
    'For i = 1 To Nd - 1
    '    If xymax < SQRxy(i + 1) Then xymax = SQRxy(i + 1)
    'Next i
    '除以最大绝对值(限制在(-1,1)之内
    'For i = 1 To Nd: D(i).X = D(i).X / xymax: D(i).Y = D(i).Y / xymax: Next i
    
    xymax = SQRxy(1)
    For i = 1 To Nd - 1
        If xymax < SQRxy(i + 1) Then xymax = SQRxy(i + 1)
    Next i
    '除以最大绝对值(限制在(-1,1)之内
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
OpenTextFileError0:            '打开文件总的错误
    MsgBox ("读文件模块:打开文件总的错误")
OpenTextFileExit:
End Sub

