#region Functions

function Execute-ProvisionJob ($progressStep, $number, $provisionJob) {
	$ProgressBar = New-Object System.Windows.Forms.ProgressBar
	$ProgressBar.Location = New-Object System.Drawing.Point(10, 50)
	$ProgressBar.Size = New-Object System.Drawing.Size(370, 20)
	$ProgressBar.Style = "Marquee"
	$ProgressBar.MarqueeAnimationSpeed = 2
	$ProgressBar.AutoSize = $true
	
	$main_form.Controls.Add($ProgressBar);
	
	$Label.ForeColor = 'black'
	$Label.Text = "$($progressStep) (Step: $($number)/10)"
	$ProgressBar.visible
	
	$job = Start-Job -ScriptBlock $provisionJob
	do { [System.Windows.Forms.Application]::DoEvents() }
	until ($job.State -eq "Completed")
	Remove-Job -Job $job -Force
	
	$Label.Text = "Provisioning Complete"
	$ProgressBar.Hide()
	$EndButton.Visible
}

#Logging
function Write-Log {
	param
	(
		[parameter(Mandatory = $true)]
		[String]$logMessage
	)
	
	Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | $($logMessage)"
}

#endregion Functions

#region Dependecies

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

#endregion Dependecies

#region Build Form

$main_form = New-Object System.Windows.Forms.Form
$main_form.BackColor = 'WhiteSmoke'
$main_form.Width = 400
$main_form.Height = 95
$main_form.ControlBox = $false
$main_form.TopMost = $true
$main_form.FormBorderStyle = 'Fixed3D'
$main_form.AutoScaleDimensions = '6, 13'
$main_form.AutoScaleMode = 'None'

$header = New-Object System.Drawing.Font("Verdana", 13, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$procFont = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$Label = New-Object System.Windows.Forms.Label
$Label.ForeColor = 'black'
$Label.Text = "Creation in process. Please wait.."
$Label.Location = New-Object System.Drawing.Point(10, 10)
$Label.Width = 480
$Label.Height = 20

$main_form.Controls.AddRange(($Label))

$main_form.StartPosition = "CenterScreen"
$main_form.Location = New-Object System.Drawing.Size(500, 300)
$result = $main_form.Show()

#endregion Build Form

#region Job ScriptBlocks
$downloadFiles =
{
	#log
	Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | INFO: Starting installation files download"
	
	# Create installation files directory
	$installationDir = "C:\APPS"
	New-Item -Path $installationDir -ItemType Directory
	
	# URL of OneDrive share
	$url = "https://stagwell-my.sharepoint.com/:u:/g/personal/joe_moscho_stagwellgtg_com/EcZTwC_LvoFEqCnNCU76vFgBesZ4lBqxLyWS-MNG7NIj4Q?e=VbZ48L&download=1"
	
	# Hide progress bar to speed up web request
	$ProgressPreference = 'SilentlyContinue'
	
	# Download installation files
	Invoke-WebRequest -Uri $url -OutFile $installationDir\installerFiles.zip
	
	# Extract installation files
	Expand-Archive -Path C:\APPS\installerFiles.zip -DestinationPath $installationDir
	
	#log
	Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | INFO: Finished installation files download"
}

$computerName =
{
	Rename-Computer -ComputerName $env:COMPUTERNAME -NewName "APR-$((get-ciminstance win32_bios).SerialNumber)" -Force
}

$JoinDomain =
{
	$azureVerify = dsregcmd /status | Select-String -Pattern AzureAdJoined
	if ($azureVerify -like "*YES*") {
		Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | SUCCESS: Joined Azure AD Domain"
	}
	else {
		Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | ERROR: Failed to Join Azure AD Domain"
	}
	
	Start-Sleep -Seconds 5
}

$adobeReader =
{
	#Install Adobe Reader
	Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | INFO: Installing Adobe Reader"
	cmd /c "C:\APPS\adobereader.exe" /qn EULA_ACCEPT=YES AgreeToLicense=Yes RebootYesNo=No /sAll
	
	#Verify Adobe Reader
	if ((Test-Path -Path "C:\ProgramData\Adobe") -eq "True") {
		Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | SUCCESS: Adobe Reader installed successfully"
	}
	else {
		Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | ERROR: Adobe Reader failed to install. ERROR: $($Error[0])"
	}
}

$zoomClient = 
{
	#Install Zoom
	Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | INFO: Installing Zoom"
	Start-Process C:\Windows\System32\msiexec.exe -Wait -ArgumentList "/i C:\APPS\ZoomInstallerFull.msi /qn"

	#Verify Chrome
	if ((Test-Path -Path "C:\Program Files\Zoom\bin\Zoom.exe") -eq "True") {
		Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | SUCCESS: Zoom installed successfully"
	}
	else {
		Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | ERROR: Zoom failed to install"
	}
}

$googleChrome =
{
	#Install Google Chrome
	Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | INFO: Installing Google Chrome"
	Start-Process C:\Windows\System32\msiexec.exe -Wait -ArgumentList "/i C:\APPS\chrome.msi /qn /norestart"
	
	#Verify Chrome
	if ((Test-Path -Path "C:\Program Files\Google\Chrome") -eq "True") {
		Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | SUCCESS: Google Chrome installed successfully"
	}
	else {
		Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | ERROR: Google Chrome failed to install"
	}
}

$kaseyaAgent =
{
	#Install Kaseya
	Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | INFO: Installing Kasyea Agent"
	cmd /c "C:\APPS\kcsSetup.exe"
	
	#Verify Kaseya Agent
	if ((Test-Path -Path "C:\ProgramData\Kaseya") -eq "True") {
		Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | SUCCESS: Kaseya Agent installed successfully"
	}
	else {
		Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | ERROR: Kaseya Agent failed to install"
	}
}

$ESETAgent =
{
	# URL of OneDrive share
	$url = "https://redirector.eset.systems/li-handler/?uuid=epi_win-08a2408d-df3c-4084-8815-5ab7d758e5fc"

	# Hide progress bar to speed up web request
	$ProgressPreference = 'SilentlyContinue'

	# Download installation files
	Invoke-WebRequest -Uri $url -OutFile "C:\APPS\ESET_Installer.exe"

	# Install ESET
	Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | INFO: Installing ESET Agent"
	Start-Process -FilePath "C:\APPS\ESET_Installer.exe" -Wait -ArgumentList "--silent --accepteula"
	
	#Verify ESET
	if ((Test-Path -Path "C:\ProgramData\ESET\RemoteAdministrator\Agent\SetupData\Installer\Agent_x64.msi") -eq "True") {
		Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | SUCCESS: ESET Agent installed successfully"
	}
	else {
		Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | ERROR: ESET Agent failed to install"
	}
}

$office365 =
{
	#Install Office 365
	Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | INFO: Installing Office 365"
	cmd /c "C:\APPS\Setup.exe" SETUP /Configure "C:\APPS\Configuration.xml"
	
	#Wait for GUI to launch
	Start-Sleep 5
	
	#Kill GUI
	Start-Process .\taskkill.exe -wait -ArgumentList "/F /IM OfficeC2RClient.exe /T"
	
	#Verify Office 365
	if ((Test-Path -Path "C:\Program Files\Microsoft Office") -eq "True") {
		Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | SUCCESS: Microsoft Office installed successfully"
	}
	else {
		Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | ERROR: Microsoft Office failed to install"
	}
}

$displayResults =
{
	#Display Provisioning Results
	Add-Content -Path C:\DevProvisionLog.txt -Value "$(get-date -Format "MM/dd/yyyy HH:mm:ss") | INFO: Displaying Results"
}

$test =
{
	Start-Sleep 15
}

#endregion Job ScriptBlocks

#region Main Process
$num = 0

#Installation Files
Execute-ProvisionJob -progressStep "Downloading Installation Files" -number ($num = $num + 1) -provisionJob $downloadFiles

#Updating computer name
Execute-ProvisionJob -progressStep "Setting Computer Name" -number ($num = $num + 1) -provisionJob $computerName

#Join Domain
Execute-ProvisionJob -progressStep "Verifying Device is Joined to Azure AD" -number ($num = $num + 1) -provisionJob $JoinDomain

#Adobe Reader
Execute-ProvisionJob -progressStep "Installing Adobe Reader" -number ($num = $num + 1) -provisionJob $adobeReader

#Zoom Client
Execute-ProvisionJob -progressStep "Installing Zoom Desktop Client" -number ($num = $num + 1) -provisionJob $zoomClient

#Google Chrome
Execute-ProvisionJob -progressStep "Installing Google Chrome" -number ($num = $num + 1) -provisionJob $googleChrome

#Kaseya
Execute-ProvisionJob -progressStep "Installing Kaseya Agent" -number ($num = $num + 1) -provisionJob $kaseyaAgent

#Sophos
Execute-ProvisionJob -progressStep "Installing ESET Agent" -number ($num = $num + 1) -provisionJob $ESETAgent

#Office 365
Execute-ProvisionJob -progressStep "Installing Office 365" -number ($num = $num + 1) -provisionJob $office365

#Display Results
Execute-ProvisionJob -progressStep "Displaying Results" -number ($num = $num + 1) -provisionJob $displayResults

$main_form.Close()

#endregion Main Process