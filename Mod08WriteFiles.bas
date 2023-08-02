Attribute VB_Name = "Mod08"
'-----------------------------------------------------------------------------------------------
'子程序名：WriteFile(t() As Double, xy() As Double,FileName As String)
'功能：将第t数组、第xy数组按行写入数据FileName文件FileName中
'      (第t为一维数组.xy为2为数组(第1下标与t数组对应,第1下标为1 to 2)
'      转成字符形式存盘
'调用例:
'       ReDim tfx(1 To 2) As Double
'       ReDim fxy(1 To 2, 1 To 2) As Double
'       tfx(1) = -0.111
'       tfx(2) = 1.22111
'       fxy(1, 1) = 0.11222: fxy(1, 2) = 0.113333333333
'       fxy(2, 1) = 0.33222: fxy(2, 2) = 0.223333333333
'       Call WriteFile(tfx, fxy,"LS.TXT")    写入"LS.TXT"文件
'
Public Sub WriteFile(t() As Double, data() As xy, FileName As String)
    '声明本子程序使用的临时变量
    Dim Iw As Integer
    Dim Sw As String
    '打开文件供写入(若原有数据,会被清空)
    Open FileName For Output As #1      '打开文件
    '逐行写入
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
        
        Print #1, Sw             '写入1行(Sw)
    Next Iw
    '写入xymax,Sumx,Sumy,
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
    Print #1, Sw              '写入1行(Sw)
    '关闭文件
    Close #1
End Sub
