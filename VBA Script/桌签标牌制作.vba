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
'  源码地址: https://club.excelhome.net/thread-1649038-1-1.html
'  修改人: 小恐龙
'  使用方法:  这是一个桌签标牌制作代码,  新建word文档内输入参会人员姓名,每行一个姓名.运行本文件即可


Option Explicit    '强制声明变量是个好习惯, 可以提高代码速度,减少bug
Sub xkonglong_桌签标牌()
'标牌
    Dim c As Cell, i&, j&
    DataInit
    With ActiveDocument
        .PageSetup.Orientation = wdOrientLandscape
        ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitFullPage
        Selection.InsertColumnsRight
        .Tables(1).AutoFitBehavior (wdAutoFitWindow)
        j = .Tables(1).Rows.Count
        With .Tables(1)
            For i = 1 To j
                .Cell(i, 2).Range.Text = .Cell(i, 1).Range.Text
            Next i
        End With
        .Content.Find.Execute "^p", , , , , , , , , "", 2
        With .Tables(1).Range
            .Style = "普通表格"
            .Rows.HeightRule = wdRowHeightExactly
            .Rows.Height = CentimetersToPoints(14.6)
            .Cells.VerticalAlignment = wdCellAlignVerticalCenter
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            With .Font
                .NameFarEast = "黑体"
                .Size = 72
                .Bold = True
            End With
            .Columns(1).Select
            .Orientation = wdTextOrientationDownward
            For Each c In .Columns(2).Cells
                c.Range.Orientation = wdTextOrientationUpward
            Next
        End With
    End With
    LastPound
    ActiveWindow.View.TableGridlines = True
    Selection.HomeKey 6
End Sub


Function DataInit()
    Dim c As Cell
    With ActiveDocument
        .Content.Find.Execute "^l", , , 0, , , , , , "^p", 2
        .Select
        Selection.ClearFormatting
        DeleteBlankSpace
        DeleteBlankLines
        If .Tables.Count = 0 Then
            .Content.ConvertToTable 0, 1
        ElseIf .Tables.Count = 1 Then
            .Content.Find.Execute "^p", , , 0, , , , , , "", 2
        Else
            MsgBox "仅限一表！", 0 + 16: End
        End If
        With .Tables(1)
            .Select
            TableDeleteBlankRows
            For Each c In .Range.Cells
                With c.Range
                    If .Text Like "????" Then .Characters(1).InsertAfter Text:="  "
                End With
            Next
            .Select
        End With
    End With
End Function


Function LastPound()
'最后一磅
    With ActiveDocument.Paragraphs
        With .Last.Range
            If .Text = vbCr Then .Delete
        End With
        With .Last.Range
            If .Text = vbCr Then
                With .Font
                    .Size = 1
                    .Kerning = 0
                    .DisableCharacterSpaceGrid = True
                End With
                With .ParagraphFormat
                    .LineSpacing = LinesToPoints(0.06)
                    .AutoAdjustRightIndent = False
                    .DisableLineHeightGrid = True
                End With
            End If
        End With
    End With
End Function


Function DeleteBlankLines()
'删除空行
    Dim i As Paragraph
    For Each i In ActiveDocument.Paragraphs
        With i.Range
            If Not .Information(12) Then
                If Asc(.Text) = 13 Then .Delete
            End If
        End With
    Next
End Function

Function DeleteBlankSpace()
'删除空格
    ActiveDocument.Content.Find.Execute "[ 　^s^t]", , , 1, , , , , , "", 2
End Function

Function TableDeleteBlankRows()
'表格删除空行
    Dim r As Row
    For Each r In Selection.Tables(1).Rows
        If Len(Replace(Replace(r.Range, vbCr, ""), Chr(7), "")) = 0 Then r.Delete
    Next
End Function