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
'  更新日期: 2024年7月5日


Sub 复制文件本地路径()

    Dim sContent
    Dim oDataObject
    
    sContent = LocalPath()
    
    Set oDataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    With oDataObject
        '给DataObject对象写入内容
        .SetText sContent
        '将DataObject对象的内容存入剪贴板
        .PutInClipboard
        '将剪贴板中的内容存入DataObject对象中
        
    End With
    
End Sub
    
Public Function LocalPath() As String
    Dim ShortName As String
    Dim i%
    ShortName = ActiveDocument.FullName
    If VBA.InStr(1, ActiveDocument.FullName, "http") >= 1 Then '如果这个文件是网盘文件
        'MsgBox ShortName
        ShortName = VBA.Replace(ShortName, "/", "\")
        For i = 1 To 4 '删除跟Onedrive路径有关的前缀:
            ShortName = VBA.Mid(ShortName, InStr(1, ShortName, "\", vbTextCompare) + 1)
    '        Debug.Print i & ":   " & ShortName
        Next i
          LocalPath = VBA.Environ("OneDrive") & "\" & ShortName
    Else
          LocalPath = ActiveDocument.FullName
    End If
	'MsgBox LocalPath
    
End Function