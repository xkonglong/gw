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
'  日期: 2024年3月22日

Option Explicit    '强制声明变量是个好习惯, 可以提高代码速度,减少bug

Sub 复制文件路径并打开微信()
    
    On Error Resume Next
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    If Not doc.Saved Then
        doc.Save  '如果文件未保存, 先保存 '
    End If
    
    Dim path As String
    
    path = doc.FullName  '获取文件路径'
    
    '剪贴板操作
    Dim RR
    Set RR = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")   'Forms 2.0 Object Library
    RR.SetText path
    RR.PutInClipboard  '复制文件路径到剪贴板
    
    
    '打开微信, 路径以你自己的微信路径为准
    Shell "D:\Tools\WeChat\WeChat.exe"   '
    
    
End Sub
