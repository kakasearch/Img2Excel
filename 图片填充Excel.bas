Attribute VB_Name = "Excel±³¾°É«×Ô¶¨Òå"
Function SelectFile()
    'Ñ¡Ôñµ¥Ò»ÎÄ¼þ
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False   'µ¥Ñ¡Ôñ
        .Filters.Clear   'Çå³ýÎÄ¼þ¹ýÂËÆ÷
        .Filters.Add "ÕÕÆ¬", "*.png,*.jpg,*.jpeg"
        .Filters.Add "ËùÓÐÎÄ¼þ", "*.*"          'ÉèÖÃÁ½¸öÎÄ¼þ¹ýÂËÆ÷
        If .Show = -1 Then    'FileDialog ¶ÔÏóµÄ Show ·½·¨ÏÔÊ¾¶Ô»°¿ò£¬²¢ÇÒ·µ»Ø -1£¨Èç¹ûÄú°´ OK£©ºÍ 0£¨Èç¹ûÄú°´ Cancel£©¡£
           ' MsgBox "ÄúÑ¡ÔñµÄÎÄ¼þÊÇ£º" & .SelectedItems(1), vbOKOnly + vbInformation, "ÖÇÄÜExcel"
            SelectFile = .SelectedItems(1)
        Else
           SelectFile = ""
        End If
    End With
End Function
Function HEX2d(val)
    Rem 16½øÖÆ×ª10½øÖÆ
   HEX2d = 1 * ("&H" & [val])
End Function
Function get_color(val)
    Rem ×ª»¯ÎªrgbÐÎÊ½Êä³ö
    Dim s As String
    Dim r As String
    Dim g As String
    Dim b As String
    s = Hex(val)
    s = "0" & s
    'Debug.Print s
    s = Right(s, 6)
    r = Left(s, 2)
    g = Mid(s, 3, 2)
    b = Right(s, 2)
    get_color = rgb(HEX2d(r), HEX2d(g), HEX2d(b))
End Function

Function set_img(path As String)
'¸ø¶¨ÕÕÆ¬ÊäÈëÂ·¾¶£¬½«ÕÕÆ¬ÉèÖÃÎªexcel±³¾°
Dim img
Set img = CreateObject("WIA.ImageFile")
img.LoadFIle (path) '¼ÓÔØÕÕÆ¬
Dim rgb
Set rgb_v = img.ARGBData   '»ñÈ¡ÏñËØÑÕÉ«£¬ÀàÐÍÎªÒ»Î¬ÏòÁ¿
'MsgBox rgb_v.Count
Dim color
For i = 1 To img.Height
    For j = 1 To img.Width
        Index = (i - 1) * img.Width + j 'ÑÕÉ«Êý×é´Ó1¿ªÊ¼
        color = get_color(rgb_v(Index))
        With ActiveWindow.ActiveSheet.Cells(i, j) 'ÉèÖÃÑÕÉ«ºÍ±ß¿ò
            .Interior.color = color
            .Borders.LineStyle = xlContinuous
        End With
    Next
    Rows(i).RowHeight = 50 'µ÷ÕûÐÐ¸ß£¬ÈÃÃ¿¸öµ¥Ôª¸ñ±ä³ÉÕý·½ÐÎ
Next
End Function

Sub replaceBg()
Rem ½«ÕÕÆ¬ÏñËØÌî³äÈëExcelµ¥Ôª¸ñ
Dim a As String
a = SelectFile()
If a Then
    set_img (a)
    MsgBox "ÒÑ¾­Ìæ»»"
Else
    MsgBox "ÇëÑ¡ÔñÎÄ¼þ"
End If

End Sub
