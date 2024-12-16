'====================================================================
'                            使用说明:  
'制作 VBA 文件时, 请只保留一个Sub 用于执行,  其他过程请用 Function 关键字。
'Sub 过程名推荐用 用户ID_中文名称 的方式命名,避免代码冲突。
'VBA 文件请使用ANSI(简体中文GB2312)编码保存, 微软的VBA解释器不支持UTF-8这类编码,会导致中文乱码。
'不正确的代码可能导致word崩溃、文档丢失或损坏。
'请务必保证vba代码来源安全可靠，插件作者不承担vba文件造成的任何损失！
'分享 VBA 文件时，请尊重作者版权，注明来源以示感谢。
'====================================================================

'  原作者: 413191246se
'  源码地址: https://club.excelhome.net/thread-1637945-1-1.html
'  修改人: 小恐龙

Option Explicit    '强制声明变量是个好习惯, 可以提高代码速度,减少bug
Sub xkonglong_页面纵横转换()
'指定页面纵横转换
    Dim p$, q$
    q = ActiveDocument.Content.Information(wdActiveEndPageNumber)
    p = InputBox("本文档共有 " & q & " 页！", "请输入页码！", q)
    If p = "" Then End
    If p > q Then MsgBox "超出页码范围！", 0 + 16: End
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, Name:=p
    ActiveDocument.Bookmarks("\Page").Range.Select
    If p = q Then Selection.MoveEnd
    Vertical
End Sub


Function Vertical()
'纵横转换
    With Selection
        If .Type = wdSelectionIP Then End
        If .Start <> 0 Then
            ActiveDocument.Range(Start:=.Start, End:=.Start).InsertBreak Type:=wdSectionBreakNextPage
            .Start = .Start + 1
        End If
        If .End <> ActiveDocument.Content.End Then
            ActiveDocument.Range(Start:=.End, End:=.End).InsertBreak Type:=wdSectionBreakNextPage
        End If
        With .PageSetup
            If .Orientation = wdOrientPortrait Then .Orientation = wdOrientLandscape Else .Orientation = wdOrientPortrait
        End With
    End With
End Function
