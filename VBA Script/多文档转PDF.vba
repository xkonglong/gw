'====================================================================
'                            使用说明:
'制作 VBA 文件时, 请只保留一个Sub 用于执行,  其他过程请用 Function 关键字。
'Sub 过程名推荐用 用户ID_中文名称 的方式命名,避免代码冲突。
'VBA 文件请使用ANSI(简体中文GB2312)编码保存, 微软的VBA解释器不支持UTF-8这类编码,会导致中文乱码。
'不正确的代码可能导致word崩溃、文档丢失或损坏。
'请务必保证vba代码来源安全可靠，插件作者不承担vba文件造成的任何损失！
'分享 VBA 文件时，请尊重作者版权，注明来源以示感谢。
'====================================================================

'  原作者: Bingo 260961242
'  源码地址: https://club.excelhome.net/thread-1125035-1-2.html
'  修改人: 小恐龙
'  修改日期: 2023年2月7日

Option Explicit    '强制声明变量是个好习惯, 可以提高代码速度,减少bug



Sub 多文档转PDF()

Application.DisplayAlerts = True
Application.ScreenUpdating = False


Dim fDialog As FileDialog
Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
Dim vrtSelectedItem As Variant
Dim wdDoc As Document
Dim i As Long

With fDialog
    .Filters.Add "Word文件", "*.doc;*.docx;*.docm", 1
    If .Show = -1 Then
        For Each vrtSelectedItem In .SelectedItems
            '如果选择了本文档则跳过
            If InStrRev(vrtSelectedItem, ThisDocument.Name) = 0 Then
                On Error Resume Next
                Set wdDoc = Application.Documents.Open(vrtSelectedItem, ReadOnly:=True)
                If Right(vrtSelectedItem, 4) = "docx" Then
                    wdDoc.SaveAs Left(vrtSelectedItem, Len(vrtSelectedItem) - 5), wdFormatPDF
                    i = i + 1
                 Else
                    wdDoc.SaveAs Left(vrtSelectedItem, Len(vrtSelectedItem) - 4), wdFormatPDF
                    i = i + 1
                End If
                'Debug.Print Left(vrtSelectedItem, Len(vrtSelectedItem) - 4)
                wdDoc.Close False
            End If
        Next vrtSelectedItem

    End If
End With
Set fDialog = Nothing
Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox (i & "个文档已转换!")


End Sub