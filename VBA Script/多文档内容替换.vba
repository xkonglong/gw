'====================================================================
'                            使用说明:
'制作 VBA 文件时, 请只保留一个Sub 用于执行,  其他过程请用 Function 关键字。
'Sub 过程名推荐用 用户ID_中文名称 的方式命名,避免代码冲突。
'VBA 文件请使用ANSI(简体中文GB2312)编码保存, 微软的VBA解释器不支持UTF-8这类编码,会导致中文乱码。
'不正确的代码可能导致word崩溃、文档丢失或损坏。
'请务必保证vba代码来源安全可靠，插件作者不承担vba文件造成的任何损失！
'分享 VBA 文件时，请尊重作者版权，注明来源以示感谢。
'====================================================================
'作者: 小恐龙'
'日期:2023年2月2日'

'---------------------------------------------------------------------
'本示例主要演示了几个特性:
'   1. VBA脚本的批量文件处理能力
'   2. 变量既可以用英文名称,也可以用中文, 使用中文变量会很大程度上帮助新手理解代码

'---------------------------------------------------------------------

Option Explicit    '强制声明变量是个好习惯, 可以提高代码速度,减少bug
Dim strF As String
Dim strR As String

Sub xkonglong_批量替换()

    Dim 筛选器 As FileDialog, 文件名, 文件 As Document, pd, 计数 As Integer
    Set 筛选器 = Application.FileDialog(msoFileDialogFilePicker)
    With 筛选器
        .AllowMultiSelect = True
        .Title = "请选择要批量替换的文档,可多选"
        .Show
    End With
    
    
    strF = InputBox("请输入要查找的内容,支持通配符,默认是查找下划线_", , "([_]{1,})")
    strR = InputBox("请输入替换文本,支持通配符,默认是给查找内容加()", , "(\1)")
    
    For Each 文件名 In 筛选器.SelectedItems
        If Not Right(文件名, Len(文件名) - InStrRev(文件名, ".")) Like "doc*" Then GoTo 结束
        Set 文件 = Documents.Open(文件名)
        With 文件.Content.find
            .ClearFormatting
            .MatchWildcards = True    '支持通配符
            .Wrap = wdFindStop
            .Text = strF           '要查找的内容'
            With .Replacement
                .ClearFormatting
                .Text = strR        '替换后的内容'
            End With
            .Execute Replace:=wdReplaceAll   '替换所有'
        End With
        文件.Close wdSaveChanges     '保存文件'
        Debug.Print 文件名 & " 已处理完成！"
        计数 = 计数 + 1
        Set 文件 = Nothing
结束:  Next
    Set 筛选器 = Nothing
    MsgBox "已完成！共处理了" & 计数 & "个文件。"
End Sub



