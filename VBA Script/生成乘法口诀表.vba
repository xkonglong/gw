Sub 乘法口诀表()
    Dim i&, j&
    With Selection
        For j = 1 To 9
            For i = 1 To 9
                .TypeText Text:=i & " × " & j & " = " & i * j & Chr(-24159)
                If i = j Then .TypeParagraph: Exit For
            Next
        Next
        .HomeKey 6
    End With

    With ActiveDocument
        With .Content
            .Find.Execute "(?)(^13)", , , 1, , , , , , "\2", 2
            With .Font
                .NameAscii = "Times New Roman"
                .Size = 15
                .Bold = True
            End With
            .InsertBefore Text:="乘法口诀表" & vbCr
        End With

        With .Paragraphs(1)
            .Style = wdStyleHeading1
            .SpaceBefore = 0
            .SpaceAfter = 0
            .Range.Underline = wdUnderlineDouble
        End With

        With .PageSetup
            .Orientation = wdOrientLandscape
'            .TopMargin = CentimetersToPoints(0.3)
'            .BottomMargin = CentimetersToPoints(2.5)
'            .LeftMargin = CentimetersToPoints(2.54)
'            .RightMargin = CentimetersToPoints(2.54)
            .PageWidth = CentimetersToPoints(29.7)
            .PageHeight = CentimetersToPoints(21)
        End With
    End With

    With ActiveWindow.ActivePane.View
        .Zoom.PageFit = wdPageFitFullPage
        .Zoom.PageFit = wdPageFitBestFit
        .ShowAll = False
    End With

    Selection.MoveUp Unit:=wdScreen, Count:=1
End Sub