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
'  修改日期: 2023年2月14日

'  本功能用于移除工作簿保护的密码.

Sub xkonglong_移除xls写保护()
    
    Dim xls As Object
    Set xls = CreateObject("excel.application")
    
    Dim dia As FileDialog
    Set dia = Application.FileDialog(msoFileDialogFilePicker)
    With dia
        .AllowMultiSelect = False
        .Title = "打开一个excel文件"
        .Filters.Add "Excel 文件", "*.xls*", 1
        .Show
    End With
    
    Dim fname As String
    fname = dia.SelectedItems(1)
    

    Dim book As Object
    
    Set book = xls.Workbooks.Open(fname)
    
    Dim sht As Object
    
    '循环清理sheet写保护

    For Each sht In book.Worksheets

        sht.Protect AllowFiltering:=True

        sht.Unprotect

    Next
    book.Close (True)  'xls文件保存并关闭'
    Set xls = Nothing  '关闭excel后台程序'
    MsgBox (fname & "写保护已清除")

End Sub

