'====================================================================
'                            使用说明:
'制作 VBA 文件时, 请只保留一个Sub 用于执行,  其他过程请用 Function 关键字。
'Sub 过程名推荐用 用户ID_中文名称 的方式命名,避免代码冲突。
'VBA 文件请使用ANSI(简体中文GB2312)编码保存, 微软的VBA解释器不支持UTF-8这类编码,会导致中文乱码。
'不正确的代码可能导致word崩溃、文档丢失或损坏。
'请务必保证vba代码来源安全可靠，插件作者不承担vba文件造成的任何损失！
'分享 VBA 文件时，请尊重作者版权，注明来源以示感谢。
'====================================================================
'原作者: https://blog.csdn.net/qilei2010/article/details/85253970
'小恐龙略作修改
'用法: 选中数字后执行本脚本即可''

Sub 转换成人民币()

'将所选数字转为人民币大写
    Dim rmb As String
    rmb = mychange(Selection.Text)
    
    Selection.Text = rmb

End Sub

Function mychange(ByVal Myinput)
    'MyinputA 去除空白符且变成整数（去掉小数）后的数字串
    'MyinputB 翻转后的数字串
    'MyinputC 转换为大写的金额
    Dim Temp, TempA, MyinputA, MyinputB, MyinputC
    Dim Place As String
    Dim J As Integer
    Place = "分角元拾佰仟万拾佰仟亿拾佰仟万"
    shuzi1 = "壹贰叁肆伍陆柒捌玖"
    shuzi2 = "整零元零零零万零零零亿零零零万"
    
    qianzhui = ""
    If Val(Myinput) = 0 Then Myinput = 0
    If Myinput = "" Then Myinput = 0
    If Myinput < 0 Then qianzhui = "负"
    
    '将小数转为整数，去掉小数点, 123.45 -> 12345
    Myinput = Int(Abs(Myinput) * 100 + 0.5)
    If Myinput > 99999999999# Then
      mychange = "输入有误：数字过大"
      Exit Function
    End If
    If Myinput = 0 Then
      mychange = "零元零分"
      Exit Function
    End If
    
    MyinputA = Trim(Str(Myinput))
    shuzilong = Len(MyinputA)
    
    '翻转金额，12345->54321
    For J = 1 To shuzilong
    MyinputB = Mid(MyinputA, J, 1) & MyinputB
    Next
    
    '1把阿拉伯数字转为大写， 54321， 5->伍
    '2将数字和对应位置的单位拼接，伍肆叁贰壹，伍->伍分
    '3拼接时翻转回来， 肆角伍分
    '注意0：从 shuzi2 得到单位，而不是从 Place
    '       12.10->1210->0121->  整 壹角 贰元 壹拾
    '       10.88->1088->8801->捌分 捌角   元 壹拾
    '       30800.25->3080025->5200803->..贰角 元 零 捌佰 零 叁万
    '               ->叁万 零 捌佰 零 元 贰角...
    For J = 1 To shuzilong
      Temp = Val(Mid(MyinputB, J, 1))
      If Temp = 0 Then
         MyinputC = Mid(shuzi2, J, 1) & MyinputC
      Else
         MyinputC = Mid(shuzi1, Temp, 1) & Mid(Place, J, 1) & MyinputC
      End If
    Next
    
    '细节：处理零
    '10.46          壹拾零元... -> 壹拾元
    '10 1234.56     壹拾零万... -> 壹拾万
    '10 1234 5678.56壹拾零亿... -> 壹拾亿
    '30800.25       上一步得到：叁万 零   捌佰 零     元 贰角伍分
    '               注意并不是：叁万 零仟 捌佰 零拾 零元 贰角伍分
    '30800.25       叁万零捌佰(零)元.. ->  叁万零捌佰 元..
    shuzilong = Len(MyinputC)
    For J = 1 To shuzilong - 1
      If Mid(MyinputC, J, 1) = "零" Then
         Select Case Mid(MyinputC, J + 1, 1)
            Case "零", "元", "万", "亿", "整":
            MyinputC = Left(MyinputC, J - 1) & Mid(MyinputC, J + 1, 30)
            J = J - 1
         End Select
      End If
    Next
    
    '贰亿万... -> 贰亿...
    shuzilong = Len(MyinputC)
    For J = 1 To shuzilong - 1
       If Mid(MyinputC, J, 1) = "亿" And Mid(MyinputC, J + 1, 1) = "万" Then
         MyinputC = Left(MyinputC, J) & Mid(MyinputC, J + 2, 30)
         Exit For
       End If
    Next
    
    mychange = qianzhui & Trim(MyinputC)
End Function
