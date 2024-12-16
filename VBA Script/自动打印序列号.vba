'自动打印序列号
'源代码地址:https://blog.csdn.net/yipiantian/article/details/124245361
'小恐龙根据源码修改.

'使用方法:   
'光标移动到需要写入编号的地方,运行本脚本即可.     默认格式为: 第N份
'你可以修改编号前缀和后缀.  ""可留空
'如果对编号有格式要求, 可以先调整好编号所在位置的文字格式再运行脚本, 也可以直接修改本脚本.

Sub 自动打印序列号()
    Dim i As Long
    Dim lngStart    '开始编号
    Dim lngCount    '结束编号
    Dim leftWord As String
    Dim rightWord As String
    leftWord = "第"          '序列号前缀
    rightWord = "份"              '序列号后缀
    lngStart = InputBox("开始打印编号", "请输入开始打印编号！", 1)
    If lngStart = "" Then
        Exit Sub    '开始编号为空退出
    End If
    lngCount = InputBox("结束打印编号", "请输入结束打印编号！", 1)
    If lngCount = "" Then
        Exit Sub    '结束编号为空退出
    End If
    For i = lngStart To lngCount
        Selection.Text = leftWord & Format(i, "00") & rightWord
        '00的格式, 表示 01 02,09,10...99,  你可以根据需求自己调整.  如果为一个0, 则为 1,2,9,10...
        Application.PrintOut FileName:="", Range:=wdPrintAllDocument, Item:= _
        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
        ManualDuplexPrint:=False, Collate:=True, Background:=True, PrintToFile:= _
        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
        PrintZoomPaperHeight:=0
        Selection.Text = ""
    Next
End Sub