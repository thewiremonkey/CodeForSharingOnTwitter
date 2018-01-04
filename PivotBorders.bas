Attribute VB_Name = "PivotBorders"
Sub Pivot_Borders()
    Dim oTable As Table
    Dim oRow As Row
    Dim oCol As Column
    Dim sCellText As String
    Dim iColumnToTest As Integer
     
    iColumnToTest = 1
    For Each oTable In ActiveDocument.Tables
    oTable.Select
        With Selection.Find
        .Text = "NA"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
         'this will fail if you have vertically merged cells
        For Each oRow In oTable.Rows
             'get the cell text
            sCellText = oTable.Cell(oRow.Index, iColumnToTest).Range.Text
             'remove the end of cell character
            sCellText = Left(sCellText, Len(sCellText) - 2)
           
            
            'Debug.Print sCellText & Len(sCellText)
            If Len(sCellText) > 0 Then
            oRow.Select
            With Selection.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
            End If
        Next
        'Select the first row and add gray shading, set the header row to repeat
        oTable.Rows(1).Select
        With Selection
            Selection.Shading.Texture = wdTextureNone
            Selection.Shading.ForegroundPatternColor = wdColorAutomatic
            Selection.Shading.BackgroundPatternColor = -704577741
            Selection.Rows.HeadingFormat = True
        End With
        
        'Select the table and add a border around the whole thing.
        oTable.Select
        With Selection.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderLeft)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderBottom)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderRight)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    Next
End Sub


