Imports Microsoft.Win32

Module BrowserOps 'VB 的Module 應就類似 C的 Struct 所預設都是 Public 權限
#Region "預設瀏覽器"
    'https://ithelp.ithome.com.tw/questions/10197561
    'Property Statement:https://docs.microsoft.com/zh-tw/dotnet/visual-basic/language-reference/statements/property-statement
    'Auto-Implemented Properties:https://docs.microsoft.com/zh-tw/dotnet/visual-basic/programming-guide/language-features/procedures/auto-implemented-properties
    ReadOnly Property DefaultBrowser As String
        Get
            Const defaultBrowserkey As String =
                    "HKEY_CLASSES_ROOT\http\shell\open\command"
            DefaultBrowser = Registry.GetValue(
                defaultBrowserkey, "", Nothing)
#Region "其他參考&舊碼"
            '： https://vimsky.com/zh-tw/examples/detail/vbnet-method-microsoft.win32.registry.getvalue.html
            'https://docs.microsoft.com/zh-tw/dotnet/visual-basic/developing-apps/programming/computer-resources/how-to-read-a-value-from-a-registry-key

            '以下為舊碼'2010/10/18由http://chijanzen.net/wp/?p=156#comment-1303(取得預設瀏覽器(default web browser)的名稱? chijanzen 雜貨舖)而來.
            'Dim objShell
            'objShell = CreateObject("WScript.Shell")
            ''HKEY_CLASSES_ROOT\HTTP\shell\open\ddeexec\Application
            ''取得註冊表中的值
            'browserApp = objShell.RegRead _
            '        ("HKCR\http\shell\open\command\")
#End Region
#Region "區塊註解,多行註解"
            'https://docs.microsoft.com/zh-tw/dotnet/visual-basic/programming-guide/program-structure/comments-in-code
            'https://docs.microsoft.com/zh-tw/dotnet/visual-basic/programming-guide/program-structure/how-to-collapse-and-hide-sections-of-code
            'http://www.blueshop.com.tw/board/FUM20041006161839LRJ/BRD20030212094337TJV.html

#End Region
        End Get
    End Property
#End Region
End Module
