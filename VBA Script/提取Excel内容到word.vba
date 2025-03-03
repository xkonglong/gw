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
'  日期: 2023年2月8日

Option Explicit    '强制声明变量是个好习惯, 可以提高代码速度,减少bug

Sub 提取Excel内容()

On Error Resume Next
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

If dia.SelectedItems.Count <= 0 Then Exit Sub   '取消退出'
    
Dim book As Object
Set book = xls.Workbooks.Open(dia.SelectedItems(1))



Dim sheet As Object
Set sheet = book.Worksheets(1)



Dim i As Long

Dim doc As Document
Set doc = Application.ActiveDocument


Dim cell As String

Dim r As String
r = InputBox("输入要提取的范围,比如 A1:A10", "提取Excel内容", "A1:A10")

Dim c As Object

For Each c In sheet.Range(r).Cells
    cell = c.Value & Chr(9)
    doc.Range.InsertAfter cell
Next

xls.Workbooks.Close
Set xls = Nothing


End Sub
