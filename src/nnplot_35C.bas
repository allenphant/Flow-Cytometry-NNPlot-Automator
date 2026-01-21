Attribute VB_Name = "nnplot_35C"
Sub nnplot()
    Dim s As Slide
    Dim shp As Shape
    Dim i As Integer
    
    For Each s In ActivePresentation.Slides
        Count = s.Shapes.Count
        pic_count = 1
        'MsgBox (Count)
        lSlideHeight = s.Parent.PageSetup.SlideHeight
        lSlideWidth = s.Parent.PageSetup.SlideWidth
                        
        If Count < 10 Then
            'MsgBox Count & " < 10"
            For Each shp In s.Shapes
                'MsgBox shp.Type
                '目前圖片的shp.Type是13、文字方塊是17
                If shp.Type = 17 Then
                    shp.Top = 0
                    shp.Left = (lSlideWidth - shp.Width) / 2
                    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                End If
                If shp.Type = msoPicture Then
                    'MsgBox "This picture is msoPicture. shp.Left = " & shp.Left & "; shp.Top = " & shp.Top & "; shp.Width = " & shp.Width & "; shp.Height = " & shp.Height
                    shp.LockAspectRatio = msoTrue
                    shp.Height = 240
                    
                    If shp.Width > shp.Height * 1.5 Then
                        shp.Left = (lSlideWidth - shp.Width) / 2
                        shp.Top = 60
                    End If
                    
                    If shp.Width < shp.Height * 1.5 Then
                        shp.Left = 250 * pic_count
                        shp.Top = 300
                        pic_count = pic_count + 0.95
                    End If
                End If
            Next shp
        End If
        
        If Count > 10 Then
            i = 0
            For Each shp In s.Shapes
                If shp.Type = msoPicture Then
                    shp.LockAspectRatio = msoTrue
                    shp.Height = 112
                    'MsgBox (1 + (i Mod 7))
                    'MsgBox (1 + (i \ 7))
                    shp.Left = (s.Parent.PageSetup.SlideWidth / 7) * ((i Mod 7)) + 12
                    shp.Top = (s.Parent.PageSetup.SlideHeight / 5) * ((i \ 7)) + 2.5
                    i = i + 1
                End If
            Next shp
        End If

    Next s
End Sub



