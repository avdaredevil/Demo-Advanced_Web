function Send-Key($k) {[System.Windows.Forms.SendKeys]::SendWait($k)}
clean
Write-AP "x*Welcome to the ","n!Full Stack Development Demo"
Write-AP ">*Setting up your stack"
$a = cc bower,git,mongo -RawOutput
$p = $null;$SY = "[\+\-\*\!\#]"
Flatten $a | % {$o = $(if(!$p -or $p -match "Configured" -or $p -match "^n?$SY"){$_ -replace "^(.*?)($SY)",'$1>>$2'}else{$_});$p=$o;$o} | % {Write-AP $o}
Write-AP ">*Entering Demo Directory"
pushd "$PSHell\Web-Dev-Demo"
Write-AP "x>+Opening the ","n!powerpoint!"
$p = New-Object -com PowerPoint.application
$doc = $p.Presentations.Open("$PWD/PPT.pptx")
$p.visible=1
sleep 3
Write-AP ">>*Starting Presentation..."
$doc.SlideShowSettings.Run() | Out-Null
sleep 3
Write-AP ">>*Moving to Starting Directory"
pushd Demo-Start
#$rcWindow = New-Object RECT
#$rcClient = New-Object RECT
#[Win32]::GetWindowRect($h,[ref]$rcWindow);[Win32]::GetClientRect($h,[ref]$rcClient)
#$res = [System.Windows.Forms.Screen]::AllScreens.WorkingArea
#$width,$height = $res.width,$res.height
#$dx = ($rcWindow.Right - $rcWindow.Left) - $rcClient.Right
#$dy = ($rcWindow.Bottom - $rcWindow.Top) - $rcClient.Bottom
#[Win32]::MoveWindow($h, $width/2, 0, $width/2, $height, $true)
