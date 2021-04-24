Module LookUP
    Dim wx As String,
        browserApp As String = DefaultBrowser()
    'Replace(DefaultBrowser, " -- ""%1""", "").Trim()

    Sub lookupDicts()
        Dim qx As String = Clipboard.GetText
        If qx = "" Then End 'Exit Sub 'exit sub 會跑出表單來
        wx = 查詢字串轉換_網路碼(qx)
        If Not browserApp.IndexOf("iexplore") Then
            Dim bChrom As New BrowserChrome
            browserApp = bChrom.ChromeAppFileName
        End If
        '檢索《百度漢語》，可連至《百度百科》
        Process.Start(browserApp, "https://dict.baidu.com/s?wd=" + wx)

        If Len(qx) > 1 Then 查詢國學大師汉语字典()
        查詢漢典()

        查詢國語辭典()
        End

    End Sub
    Sub 查詢國語辭典()
        Const url = "http://dict.revised.moe.edu.tw/cbdic/search.htm"
        Process.Start(browserApp, url)
#Region "oldCode"
        'On Error GoTo Error_GetUserAddress
        '顯示隱藏漢語大詞典()
        'If Screen.ActiveControl.Sellength > 0 Then DoCmd.RunCommand(acCmdCopy)
        'Clipboard.SetText(x)
        'If Len(x) > 1 Then
        '    Shell(Replace(browserApp, """%1", "http://dict.revised.moe.edu.tw/cgi-bin/newDict/dict.sh?idx=dict.idx&cond=" & 查詢字串轉換_國語會碼(x) & "&pieceLen=50&fld=1&cat=&imgFont=1"))
        'Else
        '    Shell(Replace(browserApp, """%1", "http://dict.revised.moe.edu.tw/cgi-bin/newDict/dict.sh?idx=dict.idx&cond=^" & 查詢字串轉換_國語會碼(x) & "$&pieceLen=50&fld=1&cat=&imgFont=1"))
        'End If

        'Shell(Replace(browserApp, """%1", "http://dict.revised.moe.edu.tw/cbdic/search.htm"))


        'If GetUserAddress(x) = True Then
        '    ''        MsgBox "成功的跟隨超連結。"
        '    '        DoEvents
        '    '        SendKeys "{tab 5}" '如果Yahoo.MSN.Google工具列全開的話
        'Else
        '    MsgBox("無法跟隨超連結。")
        'End If
#End Region
    End Sub
    Sub 查詢漢典()
        Dim url = "https://www.google.com.tw/search?q=site:http://www.zdic.net/+" & wx &
                " " & "http://www.zdic.net/search/?q=" & wx
        Process.Start(browserApp, url)
#Region "舊碼"
        'Shell(Replace(browserApp, """%1", "https://www.google.com.tw/search?q=site:http://www.zdic.net/+" & wx))
        'Shell(Replace(browserApp, """%1", "http://www.zdic.net/search/?q=" & wx)) 'http://www.zdic.net/search/?q=%E8%AD%A6%E7%9B%AE&c=2
#End Region
    End Sub
    Sub 查詢國學大師汉语字典() 'http://www.guoxuedashi.net/zidian/93F5.html
        Dim url As String = "http://www.guoxuedashi.net/so.php?sokeytm=" & wx & "&ka=100&submit=" &
            " " & "http://tw.ichacha.net/zaoju/" & wx & ".html"
        Try 'https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/try-catch-finally-statement?f1url=%3FappId%3DDev16IDEF1%26l%3DEN-US%26k%3Dk(vb.Try);k(TargetFrameworkMoniker-.NETFramework,Version%253Dv4.0);k(DevLang-VB)%26rd%3Dtrue
            Process.Start(browserApp, url)
        Catch ex As Exception
            MsgBox(browserApp + ex.Message)
        End Try
#Region "舊碼"
        'Sub 查詢國學大師汉语字典(x As String)
        'Dim u8 As System.Text.Encoding = System.Text.Encoding.Unicode
        'Dim bytes As Byte() = u8.GetBytes(x)
        'Const HDurl As String = "http://www.guoxuedashi.net/so.php?sokeytm="
        'Shell(Replace(browserApp, """%1", HDurl & wx & "&ka=100&submit="))
        'Shell(Replace(browserApp, """%1", "http://tw.ichacha.net/zaoju/" & wx & ".html")) '查詢查查網造句 2016/10/11
#End Region
    End Sub

End Module
