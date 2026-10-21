'====================================================================
'                            使用说明:
'制作 VBA 文件时, 请只保留一个Sub 用于执行,  其他过程请用 Function 关键字。
'Sub 过程名推荐用 用户ID_中文名称 的方式命名,避免代码冲突。
'VBA 文件请使用ANSI(简体中文GB2312)编码保存, 微软的VBA解释器不支持UTF-8这类编码,会导致中文乱码。
'不正确的代码可能导致word崩溃、文档丢失或损坏。
'请务必保证vba代码来源安全可靠，插件作者不承担vba文件造成的任何损失！
'分享 VBA 文件时，请尊重作者版权，注明来源以示感谢。
'====================================================================

'作者：小恐龙
'日期：2025年8月31日
'说明：重写代码，解决了有时候生成的pdf不含扩展名的情况。

Option Explicit

Sub 多文档转PDF()
    ' 关闭弹窗和屏幕刷新，提升效率
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim fDialog As FileDialog
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    Dim vrtSelectedItem As Variant  ' 选中的文件路径
    Dim wdDoc As Document           ' 打开的Word文档
    Dim i As Long                   ' 转换成功计数
    Dim fullPath As String          ' 完整文件路径
    Dim fileNameWithoutExt As String ' 不含扩展名的文件名
    Dim lastDotPos As Long          ' 最后一个"."的位置
    
    ' 设置文件选择框：仅选Word文件
    With fDialog
        .Filters.Clear
        .Filters.Add "Word文件", "*.doc;*.docx;*.docm", 1
        .Title = "选择需转换为PDF的Word文档"
        
        If .Show = -1 Then ' 用户确认选择
            For Each vrtSelectedItem In .SelectedItems
                fullPath = CStr(vrtSelectedItem)
                
                ' 跳过当前VBA所在文档（避免自己转自己）
                If InStr(LCase(fullPath), LCase(ThisDocument.Name)) = 0 Then
                    On Error Resume Next ' 捕获文件占用等错误
                    Set wdDoc = Application.Documents.Open(fullPath, ReadOnly:=True)
                    
                    If Not wdDoc Is Nothing Then ' 成功打开文档
                        ' 关键：找到最后一个"."的位置，截取左侧为“不含扩展名的文件名”
                        lastDotPos = InStrRev(fullPath, ".")
                        If lastDotPos > 0 Then ' 确保文件有扩展名（避免无后缀文件报错）
                            fileNameWithoutExt = Left(fullPath, lastDotPos - 1)
                        Else
                            fileNameWithoutExt = fullPath ' 无后缀时直接用原路径
                        End If
                        
                        ' 直接拼接.pdf，强制带扩展名
                        wdDoc.SaveAs _
                            FileName:=fileNameWithoutExt & ".pdf", _
                            FileFormat:=wdFormatPDF
                        i = i + 1 ' 计数+1
                        
                        wdDoc.Close SaveChanges:=False ' 关闭文档不保存
                        Set wdDoc = Nothing ' 释放内存
                    End If
                    On Error GoTo 0 ' 恢复默认错误处理
                End If
            Next
        End If
    End With
    
    ' 清理并恢复Word设置
    Set fDialog = Nothing
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' 提示结果
    MsgBox i & "个文档已转换为PDF", vbInformation
End Sub

