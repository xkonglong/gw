'====================================================================
'                            使用说明:
'制作 VBA 文件时, 请只保留一个Sub 用于执行,  其他过程请用 Function 关键字。
'Sub 过程名推荐用 用户ID_中文名称 的方式命名,避免代码冲突。
'VBA 文件请使用ANSI(简体中文GB2312)编码保存, 微软的VBA解释器不支持UTF-8这类编码,会导致中文乱码。
'不正确的代码可能导致word崩溃、文档丢失或损坏。
'请务必保证vba代码来源安全可靠，插件作者不承担vba文件造成的任何损失！
'分享 VBA 文件时，请尊重作者版权，注明来源以示感谢。
'====================================================================

'  原作者: chinablank
'  源码地址: https://club.excelhome.net/thread-716702-1-3.html
'  修改人: 小恐龙
'  修改日期: 2023年1月31日

Option Explicit    '强制声明变量是个好习惯, 可以提高代码速度,减少bug
Sub xkonglong_删除页眉页脚()
    On Error Resume Next
    Dim oSec As Section
    Dim mydoc As Document
    Dim i As Long
    
    Application.ScreenUpdating = False
    With ActiveDocument
        .ActiveWindow.View.Type = wdPrintView
        For Each oSec In mydoc.Sections '文档的节中循环
            For i = 9 To 10
                .ActiveWindow.View.SeekView = i '9-wdSeekCurrentPageHeader,10-wdSeekCurrentPageFooter
                .Application.Selection.WholeStory
                .Application.Selection.Delete
                .ActiveWindow.View.SeekView = 0 ' wdSeekMainDocument
                .Styles("页眉").ParagraphFormat.Borders(wdBorderBottom).LineStyle = wdLineStyleNone  '删除页眉横线
            Next
        Next
    End With
    Application.ScreenUpdating = True
    MsgBox "页眉页脚删除完毕！"
End Sub
