'====================================================================
'                            使用说明:
'制作 VBA 文件时, 请只保留一个Sub 用于执行,  其他过程请用 Function 关键字。
'Sub 过程名推荐用 用户ID_中文名称 的方式命名,避免代码冲突。
'VBA 文件请使用ANSI(简体中文GB2312)编码保存, 微软的VBA解释器不支持UTF-8这类编码,会导致中文乱码。
'不正确的代码可能导致word崩溃、文档丢失或损坏。
'请务必保证vba代码来源安全可靠，插件作者不承担vba文件造成的任何损失！
'分享 VBA 文件时，请尊重作者版权，注明来源以示感谢。
'====================================================================

'  作者: 小恐龙
'  日期: 2023年2月8日

Option Explicit    '强制声明变量是个好习惯, 可以提高代码速度,减少bug


Sub 图片批量加边框()
    Application.ScreenUpdating = False
    

'以下代码调整嵌入式图片边框
    Dim inshape As InlineShape

    For Each inshape In ActiveDocument.InlineShapes
        inshape.Select
        With inshape.Borders
            .OutsideLineStyle = wdLineStyleSingle  '边框风格'
            .OutsideColorIndex = wdColorAutomatic  '颜色为自动, 通常为黑色
            .OutsideLineWidth = wdLineWidth025pt   '边框粗细,设置边框为0.25pt
        End With
    Next
    
    
 '以下代码调整普通图片边框, 个别版本 wps 无效.很奇怪
    Dim myPic As shape

    For Each myPic In ActiveDocument.Shapes
        If myPic.Type = msoPicture Then   '如果shape类型为图片
            myPic.Select
            Selection.ShapeRange.Line.Weight = 0.25   '设置边框为0.25pt
        End If
    Next
    
    Application.ScreenUpdating = True
End Sub