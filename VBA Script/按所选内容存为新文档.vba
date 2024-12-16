'====================================================================
'                            使用说明:  
'制作 VBA 文件时, 请只保留一个Sub 用于执行,  其他过程请用 Function 关键字。
'Sub 过程名推荐用 用户ID_中文名称 的方式命名,避免代码冲突。
'VBA 文件请使用ANSI(简体中文GB2312)编码保存, 微软的VBA解释器不支持UTF-8这类编码,会导致中文乱码。
'不正确的代码可能导致word崩溃、文档丢失或损坏。
'请务必保证vba代码来源安全可靠，插件作者不承担vba文件造成的任何损失！
'分享 VBA 文件时，请尊重作者版权，注明来源以示感谢。
'====================================================================
'作者: 小恐龙 
'将所选内容存为新文档

Sub SaveSelected()

    Dim doc As PageSetup
    Set doc = ActiveDocument.PageSetup
    Selection.Copy
    
    Dim newdoc As Document
    
    Set newdoc = Documents.Add(, , wdNewBlankDocument)
    newdoc.PageSetup = doc
    Selection.Paste
    newdoc.Save

End Sub