Imports Microsoft.Win32

Module BrowserOps 'VB 的Module 應就類似 C的 Struct 所預設都是 Public 權限
#Region "預設瀏覽器"
    ReadOnly Property DefaultBrowser As String '2010/10/18由http://chijanzen.net/wp/?p=156#comment-1303(取得預設瀏覽器(default web browser)的名稱? chijanzen 雜貨舖)而來.
        Get
            Const defaultBrowserkey As String =
                    "HKEY_CLASSES_ROOT\http\shell\open\command"
            DefaultBrowser = Registry.GetValue(
                defaultBrowserkey, "", Nothing)
            '以下為舊碼
            'Dim objShell
            'objShell = CreateObject("WScript.Shell")
            ''HKEY_CLASSES_ROOT\HTTP\shell\open\ddeexec\Application
            ''取得註冊表中的值
            'browserApp = objShell.RegRead _
            '        ("HKCR\http\shell\open\command\")
        End Get
    End Property
#End Region
End Module
