'====================================================================
'                            使用说明:
'制作 VBA 文件时, 请只保留一个Sub 用于执行,  其他过程请用 Function 关键字。
'Sub 过程名推荐用 用户ID_中文名称 的方式命名,避免代码冲突。
'VBA 文件请使用ANSI(简体中文GB2312)编码保存, 微软的VBA解释器不支持UTF-8这类编码,会导致中文乱码。
'不正确的代码可能导致word崩溃、文档丢失或损坏。
'请务必保证vba代码来源安全可靠，插件作者不承担vba文件造成的任何损失！
'分享 VBA 文件时，请尊重作者版权，注明来源以示感谢。
'====================================================================
'作者:小恐龙
'日期:2023年2月8日
'

Option Explicit    '强制声明变量是个好习惯, 可以提高代码速度,减少bug


Sub 批量调整图片大小()
    On Error Resume Next
    
    Dim w As single
    Dim h As single
    w = InputBox("请输入批量调整图片的宽度(厘米),0为不调整宽度,只调整高度", , 0)
    h = InputBox("请输入批量调整图片的高度(厘米),0为不调整高度,只调整宽度", , 0)
    
    '如果高宽都为0,则退出'
    If w = 0 And h = 0 Then
        MsgBox ("未输入宽度,也未输入高度")
        Exit Sub
    End If
    
    
    '本段代码批量调整普通图片
    Dim myPic As Shape
    For Each myPic In ActiveDocument.Shapes
            If myPic.Type = msoPicture Then   '只调整图片,避免调整形状,公式,图表等类型'
                myPic.Select
                If w > 0 Then
                    myPic.Width = 28.345 * w
                ElseIf h > 0 Then
                    myPic.Height = 28.345 * h
                End If
            End If
    Next
    
    '本段代码批量调整嵌入式图片
    Dim myinPic As InlineShape
    For Each myinPic In ActiveDocument.InlineShapes
            myinPic.Select
            If w > 0 Then
                myinPic.Width = 28.345 * w
            ElseIf h > 0 Then
                myinPic.Height = 28.345 * h
            End If
    Next
End Sub
