' Author: ppxxcc@www.52pojie.cn

Set objShell = CreateObject("WScript.Shell")

strPSCode = _
"Add-Type -AssemblyName System.Windows.Forms;" & _
"Add-Type -AssemblyName System.Drawing;" & _
"$notifyIcon = New-Object System.Windows.Forms.NotifyIcon;" & _
"$notifyIcon.Icon = [System.Drawing.SystemIcons]::Information;" & _
"$notifyIcon.Text = 'Numlock Monitor';" & _
"$notifyIcon.Visible = $true;" & _
"$contextMenu = New-Object System.Windows.Forms.ContextMenu;" & _
"$menuItemToggle = New-Object System.Windows.Forms.MenuItem;" & _
"$menuItemToggle.Text = 'Disable';" & _
"$isMonitoringEnabled = $true;" & _
"$menuItemToggle.Add_Click({" & _
"    $script:isMonitoringEnabled = -not $script:isMonitoringEnabled;" & _
"    $menuItemToggle.Text = if ($script:isMonitoringEnabled) { 'Disable' } else { 'Enable' };" & _
"});" & _
"$menuItemExit = New-Object System.Windows.Forms.MenuItem;" & _
"$menuItemExit.Text = 'Exit';" & _
"$menuItemExit.Add_Click({ $notifyIcon.Dispose(); [System.Environment]::Exit(0) });" & _
"$contextMenu.MenuItems.Add($menuItemToggle);" & _
"$contextMenu.MenuItems.Add($menuItemExit);" & _
"$notifyIcon.ContextMenu = $contextMenu;" & _
"$numlockTimer = New-Object System.Windows.Forms.Timer;" & _
"$numlockTimer.Interval = 500;" & _
"$numlockTimer.Add_Tick({" & _
"    if ($script:isMonitoringEnabled -and -not [console]::NumberLock) {" & _
"        $w = New-Object -ComObject WScript.Shell;" & _
"        $w.SendKeys('{NUMLOCK}')" & _
"    }" & _
"});" & _
"$numlockTimer.Start();" & _
"[System.Windows.Forms.Application]::Run();"

Set objFSO = CreateObject("Scripting.FileSystemObject")
tempPSFile = objFSO.GetSpecialFolder(2) & "\tempScript.ps1"
Set objFile = objFSO.CreateTextFile(tempPSFile, True)
objFile.WriteLine strPSCode
objFile.Close

objShell.Run "powershell.exe -ExecutionPolicy Bypass -NoProfile -File """ & tempPSFile & """", 0, True

objFSO.DeleteFile(tempPSFile)