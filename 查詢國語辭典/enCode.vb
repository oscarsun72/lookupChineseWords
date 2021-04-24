Module enCode
    Function 查詢字串轉換_百度碼(w As String) As String '不成！有空再研究20210420 用網路碼即可
        Dim i As Integer, result As String = ""
        For index = 1 To w.Length
            result += "%" + Conversion.Hex(Convert.ToInt32(w(i)))
        Next
        Return result
#Region "'百度搜索链接中的汉字转码:"
        'function getEncodeStr(src: string): string;
        '        var i: Integer;
        'begin
        '        result := '';
        '    For i := 1 To length(src) Do begin
        '            //Dec2Hex用于返回十进制数的十六进制编码字符串
        '        result := result + '%' + Dec2Hex(ord(src[i]));
        '    End;
        'End;
        '————————————————
        '版权声明：       本文为CSDN博主「alvin_2005」的原创文章，遵循CC 4.0 BY-SA版权协议，转载请附上原文出处链接及本声明。
        '原文链接：       https : //blog.csdn.net/alvin_2005/article/details/2076174
        '
#End Region
    End Function
    Function 查詢字串轉換_國語會碼(w As String) 'Big5碼
        Dim u8 As System.Text.Encoding = System.Text.Encoding.GetEncoding("big5") 'https://msdn.microsoft.com/zh-tw/library/system.text.encoding(v=vs.110).aspx
        Dim bytes As Byte() = u8.GetBytes(w)
        查詢字串轉換_國語會碼 = PrintHexBytes(bytes)
    End Function
    Function 查詢字串轉換_網路碼(w As String) 'UTF8碼
        Dim u8 As System.Text.Encoding = System.Text.Encoding.UTF8 'System.Text.Encoding.GetEncoding("UTF16") 'https://msdn.microsoft.com/zh-tw/library/system.text.encoding(v=vs.110).aspx
        Dim bytes As Byte() = u8.GetBytes(w)
        查詢字串轉換_網路碼 = PrintHexBytes(bytes)
    End Function

    Public Function PrintHexBytes(bytes() As Byte) As String ''https://msdn.microsoft.com/zh-tw/library/system.text.encoding.utf8(v=vs.110).aspx
        PrintHexBytes = ""
        If bytes Is Nothing OrElse bytes.Length = 0 Then
            'Console.WriteLine("<none>")
            MsgBox("<none>")
        Else
            Dim i As Integer
            For i = 0 To bytes.Length - 1
                PrintHexBytes &= "%" & Hex(bytes(i))
                'PrintHexBytes &= Hex(bytes(i))
                'Console.Write("{0:X2} ", bytes(i))
            Next i
            'Console.WriteLine()
        End If
    End Function 'PrintHexBytes 

End Module