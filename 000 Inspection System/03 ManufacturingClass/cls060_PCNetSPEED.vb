
Public Class cls060_PCNetSPEED

    'dim RegTree as string = "SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}\0014"     ’USB-LANｱﾀﾞﾌﾟﾀ
    Dim RegTree As String = "SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}\0007"      'PC内臓LANｱﾀﾞﾌﾟﾀ
    
    Public Function VBSetRegistryString(ByVal Value As String, ByVal sKEY As String) As Boolean
        '==========================================================================================
        ' F.Name   : VBSetRegistryString
        ' Function : Set Resistry String
        ' Data     : KeyHead,Key,Value
        ' Return   :
        ' Remarks  :
        '==========================================================================================
        VBSetRegistryString = False

        Dim regkey As Microsoft.Win32.RegistryKey = _
        Microsoft.Win32.Registry.LocalMachine.OpenSubKey(RegTree, True)
        If regkey Is Nothing Then
            MsgBox("なし！")
        End If
        regkey.SetValue(sKEY, Value)

        Dim strings As String = DirectCast(regkey.GetValue(sKEY), String)
        regkey.Close()

        VBSetRegistryString = True

    End Function

    Public Function DisEnableNetAdapter() As Boolean
        '==========================================================================================
        ' F.Name   : Dis->EnableNetAdapter
        ' Function : Net Adapter Disable->Enable
        ' Data     : RegKey USING devcon.exe
        ' Return   : true/false
        ' Remarks  :
        '==========================================================================================
        DisEnableNetAdapter = False

        Dim regkey As Microsoft.Win32.RegistryKey = _
        Microsoft.Win32.Registry.LocalMachine.OpenSubKey(RegTree, False)
        If regkey Is Nothing Then
            MsgBox("なし！")
        End If

        Dim NetCfgInstanceID As String = DirectCast(regkey.GetValue("NetCfgInstanceID"), String)
        regkey.Close()

        regkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Control\NetWork\{4D36E972-E325-11CE-BFC1-08002BE10318}\" & NetCfgInstanceID & "\Connection", False)
        Dim PnpInstanceID As String = DirectCast(regkey.GetValue("PnpInstanceID"), String)
        regkey.Close()

        Dim results As String
        Dim psi As New System.Diagnostics.ProcessStartInfo()

        Dim ID As String = Chr(34) & "@" & PnpInstanceID & Chr(34)
        'ComSpecのパスを取得する
        psi.FileName = System.Environment.GetEnvironmentVariable("ComSpec")
        '出力を読み取れるようにする
        psi.RedirectStandardInput = False
        psi.RedirectStandardOutput = True
        psi.UseShellExecute = False
        'ウィンドウを表示しないようにする
        psi.CreateNoWindow = True
        'コマンドラインを指定（"/c"は実行後閉じるために必要）
        psi.Arguments = "/c devcon.exe disable " & ID & " /w"
        '起動
        Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start(psi)
        '出力を読み取る
        results = p.StandardOutput.ReadToEnd
        Debug.Print(results)
        If InStr(results, "1 device(s) disabled.") = 0 Then
            MsgBox("ﾈｯﾄﾜｰｸｱﾀﾞﾌﾟﾀdisableｴﾗｰ")
            GoTo NGtrap
        End If

        psi.Arguments = "/c devcon.exe enable " & ID & " /w"
        p = System.Diagnostics.Process.Start(psi)
        results = p.StandardOutput.ReadToEnd
        Debug.Print(results)
        If InStr(results, "1 device(s) enabled.") = 0 Then
            MsgBox("ﾈｯﾄﾜｰｸｱﾀﾞﾌﾟﾀEnableｴﾗｰ")
            GoTo NGtrap
        End If

        DisEnableNetAdapter = True
        'WaitForExitはReadToEndの後である必要がある
        '(親プロセス、子プロセスでブロック防止のため)
NGtrap:
        p.WaitForExit()

        psi = Nothing
        regkey = Nothing

    End Function

    Public Function DisableNetAdapter() As Boolean
        '==========================================================================================
        ' F.Name   : DisableNetAdapter
        ' Function : Net Adapter Disable
        ' Data     : RegKey USING devcon.exe
        ' Return   : true/false
        ' Remarks  :
        '==========================================================================================
        DisableNetAdapter = False

        Dim regkey As Microsoft.Win32.RegistryKey = _
        Microsoft.Win32.Registry.LocalMachine.OpenSubKey(RegTree, False)
        If regkey Is Nothing Then
            MsgBox("なし！")
        End If

        Dim NetCfgInstanceID As String = DirectCast(regkey.GetValue("NetCfgInstanceID"), String)
        regkey.Close()

        regkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Control\NetWork\{4D36E972-E325-11CE-BFC1-08002BE10318}\" & NetCfgInstanceID & "\Connection", False)
        Dim PnpInstanceID As String = DirectCast(regkey.GetValue("PnpInstanceID"), String)
        regkey.Close()

        Dim results As String
        Dim psi As New System.Diagnostics.ProcessStartInfo()

        Dim ID As String = Chr(34) & "@" & PnpInstanceID & Chr(34)
        'ComSpecのパスを取得する
        psi.FileName = System.Environment.GetEnvironmentVariable("ComSpec")
        '出力を読み取れるようにする
        psi.RedirectStandardInput = False
        psi.RedirectStandardOutput = True
        psi.UseShellExecute = False
        'ウィンドウを表示しないようにする
        psi.CreateNoWindow = True
        'コマンドラインを指定（"/c"は実行後閉じるために必要）
        psi.Arguments = "/c devcon.exe disable " & ID & " /w"
        '起動
        Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start(psi)
        '出力を読み取る
        results = p.StandardOutput.ReadToEnd
        Debug.Print(results)
        If InStr(results, "1 device(s) disabled.") = 0 Then
            MsgBox("ﾈｯﾄﾜｰｸｱﾀﾞﾌﾟﾀdisableｴﾗｰ")
            GoTo NGtrap
        End If

        DisableNetAdapter = True
        'WaitForExitはReadToEndの後である必要がある
        '(親プロセス、子プロセスでブロック防止のため)
NGtrap:
        p.WaitForExit()

        psi = Nothing
        regkey = Nothing

    End Function

    Public Function EnableNetAdapter() As Boolean
        '==========================================================================================
        ' F.Name   : EnableNetAdapter
        ' Function : Net Adapter Disable
        ' Data     : RegKey USING devcon.exe
        ' Return   : true/false
        ' Remarks  :
        '==========================================================================================
        EnableNetAdapter = False

        Dim regkey As Microsoft.Win32.RegistryKey = _
        Microsoft.Win32.Registry.LocalMachine.OpenSubKey(RegTree, False)
        If regkey Is Nothing Then
            MsgBox("なし！")
        End If

        Dim NetCfgInstanceID As String = DirectCast(regkey.GetValue("NetCfgInstanceID"), String)
        regkey.Close()

        regkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Control\NetWork\{4D36E972-E325-11CE-BFC1-08002BE10318}\" & NetCfgInstanceID & "\Connection", False)
        Dim PnpInstanceID As String = DirectCast(regkey.GetValue("PnpInstanceID"), String)
        regkey.Close()

        Dim results As String
        Dim psi As New System.Diagnostics.ProcessStartInfo()

        Dim ID As String = Chr(34) & "@" & PnpInstanceID & Chr(34)
        'ComSpecのパスを取得する
        psi.FileName = System.Environment.GetEnvironmentVariable("ComSpec")
        '出力を読み取れるようにする
        psi.RedirectStandardInput = False
        psi.RedirectStandardOutput = True
        psi.UseShellExecute = False
        'ウィンドウを表示しないようにする
        psi.CreateNoWindow = True
        'コマンドラインを指定（"/c"は実行後閉じるために必要）
        psi.Arguments = "/c devcon.exe enable " & ID & " /w"
        '起動
        Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start(psi)
        '出力を読み取る
        results = p.StandardOutput.ReadToEnd
        Debug.Print(results)
        If InStr(results, "1 device(s) enabled.") = 0 Then
            MsgBox("ﾈｯﾄﾜｰｸｱﾀﾞﾌﾟﾀEnableｴﾗｰ")
            GoTo NGtrap
        End If

        EnableNetAdapter = True
        'WaitForExitはReadToEndの後である必要がある
        '(親プロセス、子プロセスでブロック防止のため)
NGtrap:
        p.WaitForExit()

        psi = Nothing
        regkey = Nothing

    End Function

End Class
