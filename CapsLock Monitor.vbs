' Function: Monitors Caps Lock and turns it off automatically.

Set objShell = CreateObject("WScript.Shell")

' PowerShell script content starts here
strPSCode = _
"Add-Type -AssemblyName System.Windows.Forms;" & _
"Add-Type -AssemblyName System.Drawing;" & _
"$notifyIcon = New-Object System.Windows.Forms.NotifyIcon;" & _
"$notifyIcon.Icon = [System.Drawing.SystemIcons]::Warning;" & _
"$notifyIcon.Text = 'Capslock Monitor';" & _
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
"$capslockTimer = New-Object System.Windows.Forms.Timer;" & _
"$capslockTimer.Interval = 1500;" & _
"$capslockTimer.Add_Tick({" & _
"    if ($script:isMonitoringEnabled -and [console]::CapsLock) {" & _
"        $w = New-Object -ComObject WScript.Shell;" & _
"        $w.SendKeys('{CAPSLOCK}')" & _
"    }" & _
"});" & _
"$capslockTimer.Start();" & _
"[System.Windows.Forms.Application]::Run();"

' VBScript wrapper logic to run the PowerShell script silently
Set objFSO = CreateObject("Scripting.FileSystemObject")
tempPSFile = objFSO.GetSpecialFolder(2) & "\tempCapsLockScript.ps1"
Set objFile = objFSO.CreateTextFile(tempPSFile, True)
objFile.WriteLine strPSCode
objFile.Close

objShell.Run "powershell.exe -ExecutionPolicy Bypass -NoProfile -File """ & tempPSFile & """", 0, True

objFSO.DeleteFile(tempPSFile)