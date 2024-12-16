'====================================================================
'                            使用说明:
'制作 VBA 文件时, 请只保留一个Sub 用于执行,  其他过程请用 Function 关键字。
'Sub 过程名推荐用 用户ID_中文名称 的方式命名,避免代码冲突。
'VBA 文件请使用ANSI(简体中文GB2312)编码保存, 微软的VBA解释器不支持UTF-8这类编码,会导致中文乱码。
'不正确的代码可能导致word崩溃、文档丢失或损坏。
'请务必保证vba代码来源安全可靠，插件作者不承担vba文件造成的任何损失！
'分享 VBA 文件时，请尊重作者版权，注明来源以示感谢。
'====================================================================
'本代码参考自https://zhuanlan.zhihu.com/p/63921449
'原作者知乎ID: koko可可
'小恐龙室略作修改


Sub 手写体()

    Dim R_Character As Range

	'字体大小在5个值之间进行波动，可以改写
    Dim FontSize(5)
    FontSize(1) = "16"
    FontSize(2) = "18"
    FontSize(3) = "20"
    FontSize(4) = "17"
    FontSize(5) = "19"


    '括号中的数字指的是你要控制的字体波动的种数，我这里是在3种之间波动，注意如果需要变更下面也有一处要改
    Dim FontName(3)
    '字体名称这几种字体之间波动，注意要提前安装这些字体(字体自己上网找手写体,不一定是这几个)
    FontName(1) = "李国夫手写体"
    FontName(2) = "华阳手写"
    FontName(3) = "恐龙手写2"

    Dim ParagraphSpace(5)
    '行间距 在一定以下值中均等分布，可改写
    ParagraphSpace(1) = "26"
    ParagraphSpace(2) = "28"
    ParagraphSpace(3) = "30"
    ParagraphSpace(4) = "32"
    ParagraphSpace(5) = "24"

    '不懂原理的话，不建议修改下列代码
    For Each R_Character In ActiveDocument.Characters
        VBA.Randomize
        '下面这一行的2页代表字体的种类数，如果上面的更改了，这里也要改
        R_Character.Font.Name = FontName(Int(VBA.Rnd * 3) + 1)

        R_Character.Font.Size = FontSize(Int(VBA.Rnd * 5) + 1)

        R_Character.Font.Position = Int(VBA.Rnd * 3) + 1

        R_Character.Font.Spacing = 0
    Next
    Application.ScreenUpdating = True


	'此段为随机行间距, 可根据需要启用或注释掉
	'For Each Cur_Paragraph In ActiveDocument.Paragraphs
	'   Cur_Paragraph.LineSpacing = ParagraphSpace(Int(VBA.Rnd * 5) + 1)
	'Next
     Application.ScreenUpdating = True

End Sub