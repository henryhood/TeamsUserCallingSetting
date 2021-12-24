########################################################################
# Name: Teams User Calling Setting 
# Version: v0.01 (21/12/2021)
# Date: 21/12/2021
# Created By: Enrico Gualandi
# 
# Notes: This is a PowerShell tool. To run the tool, open it from the PowerShell command line on a PC that has the MicrosoftTeams PowerShell module installed. Get it by opening a PowerShell window using Run as Administrator and running "Install-Module MicrosoftTeams -AllowClobber"
#		 
#
#
########################################################################

[cmdletbinding()]
Param()

$theVersion = $PSVersionTable.PSVersion
$MajorVersion = $theVersion.Major

Write-Host ""
Write-Host "--------------------------------------------------------------"
Write-Host "PowerShell Version Check..." -foreground "yellow"
if($MajorVersion -eq  "1")
{
	Write-Host "This machine only has Version 1 PowerShell installed.  This version of PowerShell is not supported." -foreground "red"
}
elseif($MajorVersion -eq  "2")
{
	Write-Host "This machine has Version 2 PowerShell installed. This version of PowerShell is not supported." -foreground "red"
}
elseif($MajorVersion -eq  "3")
{
	Write-Host "This machine has version 3 PowerShell installed. This version of PowerShell is not supported." -foreground "red"
}
elseif($MajorVersion -eq  "4")
{
	Write-Host "This machine has version 4 PowerShell installed. This version of PowerShell is not supported." -foreground "red"
}
elseif($MajorVersion -eq  "5")
{
	Write-Host "This machine has version 5 PowerShell installed. CHECK PASSED!" -foreground "green"
}
elseif(([int]$MajorVersion) -ge  6)
{
	Write-Host "This machine has version $MajorVersion PowerShell installed. This version uses .NET Core which doesn't support Windows Forms. Please use PowerShell 5 instead." -foreground "red"
	exit
}
else
{
	Write-Host "This machine has version $MajorVersion PowerShell installed. Unknown level of support for this version." -foreground "yellow"
}
Write-Host "--------------------------------------------------------------"
Write-Host ""


$script:OnlineUsername = ""
if($OnlineUsernameInput -ne $null -and $OnlineUsernameInput -ne "")
{
	Write-Host "INFO: Using command line AdminPasswordInput setting = $OnlineUsernameInput" -foreground "Yellow"
	$script:OnlineUsername = $OnlineUsernameInput
}

$script:OnlinePassword = ""
if($OnlinePasswordInput -ne $null -and $OnlinePasswordInput -ne "")
{
	Write-Host "INFO: Using command line OnlinePasswordInput setting = $OnlinePasswordInput" -foreground "Yellow"
	$script:OnlinePassword = $OnlinePasswordInput
}

Function Get-MyModule 
{ 
Param([string]$name) 
	
	if(-not(Get-Module -name $name)) 
	{ 
		if(Get-Module -ListAvailable | Where-Object { $_.name -eq $name }) 
		{ 
			Import-Module -Name $name 
			return $true 
		} #end if module available then import 
		else 
		{ 
			return $false 
		} #module not available 
	} # end if not module 
	else 
	{ 
		return $true 
	} #module already loaded 
} #end function get-MyModule

$Script:TeamsModuleAvailable = $false

Write-Host "--------------------------------------------------------------" -foreground "green"
Write-Host "Checking for PowerShell Modules..." -foreground "green"
#Import MicrosoftTeams Module
if(Get-MyModule "MicrosoftTeams")
{
	#Invoke-Expression "Import-Module Lync"
	Write-Host "INFO: Teams module should be at least 3.0.1" -foreground "yellow"
	$version = (Get-Module -name "MicrosoftTeams").Version
	Write-Host "INFO: Your current version of Teams Module: $version" -foreground "yellow"
	if([System.Version]$version -ge [System.Version]"3.0.1")
	{
		Write-Host "Congratulations, your version is acceptable!" -foreground "green"
	}
	else
	{
		Write-Host "ERROR: You need to update your Teams Version to higher than 3.0.1. Use the command Update-Module MicrosoftTeams -AllowPrerelease" -foreground "red"
		exit
	}
	Write-Host "Found MicrosoftTeams Module..." -foreground "green"
	$Script:TeamsModuleAvailable = $true
}
else
{
	Write-Host "ERROR: You do not have the Microsoft Teams Module installed. Get it by opening a PowerShell window using `"Run as Administrator`" and running `"Install-Module MicrosoftTeams -AllowClobber -AllowPrerelease`"" -foreground "red"
	#Can't find module so exit
	exit
}

Write-Host "--------------------------------------------------------------" -foreground "green"
Import-Module MicrosoftTeams
Connect-MicrosoftTeams


# Set up the form  ============================================================

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$mainForm = New-Object System.Windows.Forms.Form 
$mainForm.Text = "Teams User Calling Setting"
$mainForm.Size = New-Object System.Drawing.Size(525,680) 
$mainForm.MinimumSize = New-Object System.Drawing.Size(520,450) 
$mainForm.StartPosition = "CenterScreen"
[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
$mainForm.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
$mainForm.KeyPreview = $True
$mainForm.TabStop = $false


$global:SFBOsession = $null
<#
#ConnectButton
$ConnectOnlineButton = New-Object System.Windows.Forms.Button
$ConnectOnlineButton.Location = New-Object System.Drawing.Size(20,7)
$ConnectOnlineButton.Size = New-Object System.Drawing.Size(110,20)
$ConnectOnlineButton.Text = "Connect Teams"
$ConnectTooltip = New-Object System.Windows.Forms.ToolTip
$ConnectToolTip.SetToolTip($ConnectOnlineButton, "Connect/Disconnect from Teams")
#$ConnectButton.tabIndex = 1
$ConnectOnlineButton.Enabled = $true
$ConnectOnlineButton.Add_Click({	

	$ConnectOnlineButton.Enabled = $false
	
	$StatusLabel.Text = "STATUS: Connecting to Teams..."
	
	if($ConnectOnlineButton.Text -eq "Connect Teams")
	{
		ConnectTeamsModule
		[System.Windows.Forms.Application]::DoEvents()
		CheckTeamsOnline
	}
	elseif($ConnectOnlineButton.Text -eq "Disconnect Teams")
	{	
		$ConnectOnlineButton.Text = "Disconnecting..."
		$StatusLabel.Text = "STATUS: Disconnecting from Teams..."
		DisconnectTeams
		CheckTeamsOnline
	}
	
	$ConnectOnlineButton.Enabled = $true
	
	$StatusLabel.Text = ""
})
$mainForm.Controls.Add($ConnectOnlineButton)
#>

$ConnectedOnlineLabel = New-Object System.Windows.Forms.Label
#$ConnectedOnlineLabel.Location = New-Object System.Drawing.Size(135,10) 
$ConnectedOnlineLabel.Location = New-Object System.Drawing.Size(20,10) 
$ConnectedOnlineLabel.Size = New-Object System.Drawing.Size(100,15) 
$ConnectedOnlineLabel.Text = "Connected"
$ConnectedOnlineLabel.TabStop = $false
$ConnectedOnlineLabel.forecolor = "green"
$ConnectedOnlineLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::left
$ConnectedOnlineLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$mainForm.Controls.Add($ConnectedOnlineLabel)
$ConnectedOnlineLabel.Visible = $false

#User Label ============================================================
$UserLabel = New-Object System.Windows.Forms.Label
$UserLabel.Location = New-Object System.Drawing.Size(22,43) 
$UserLabel.Size = New-Object System.Drawing.Size(58,15) 
$UserLabel.Text = "User: "
$UserLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$UserLabel.TabStop = $false
$mainForm.Controls.Add($UserLabel)


# Add User Dropdown box ============================================================
$UserDropDownBox = New-Object System.Windows.Forms.ComboBox 
$UserDropDownBox.Location = New-Object System.Drawing.Size(80,40) 
$UserDropDownBox.Size = New-Object System.Drawing.Size(255,20) 
$UserDropDownBox.DropDownHeight = 200 
$UserDropDownBox.DropDownWidth = 300
$UserDropDownBox.tabIndex = 1
$UserDropDownBox.DropDownStyle = "DropDownList"
$UserDropDownBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$mainForm.Controls.Add($UserDropDownBox) 


$UserDropDownBox.add_SelectedValueChanged(
{
	$StatusLabel.Text = "STATUS: Getting users list..."
	[System.Windows.Forms.Application]::DoEvents()
	GetNormalisationPolicy
	#UpdateListViewSettings
	
	$StatusLabel.Text = ""
})



#Apply button
$ApplyButton = New-Object System.Windows.Forms.Button
$ApplyButton.Location = New-Object System.Drawing.Size(392,40)
$ApplyButton.Size = New-Object System.Drawing.Size(50,50)
$ApplyButton.Text = "Apply"
$ApplyButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$ApplyButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Apply change..."
	[System.Windows.Forms.Application]::DoEvents()
    ApplyUserChanges

	GetNormalisationPolicy
	$StatusLabel.Text = ""
	
})
$mainForm.Controls.Add($ApplyButton)

#Cancel button
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(392,90)
$CancelButton.Size = New-Object System.Drawing.Size(50,50)
$CancelButton.Text = "Cancel"
$CancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$CancelButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Cancel change..."
	[System.Windows.Forms.Application]::DoEvents()
    
	GetNormalisationPolicy
	$StatusLabel.Text = ""
	
})
$mainForm.Controls.Add($CancelButton)

#Reset button
$ResetButton = New-Object System.Windows.Forms.Button
$ResetButton.Location = New-Object System.Drawing.Size(392,140)
$ResetButton.Size = New-Object System.Drawing.Size(50,50)
$ResetButton.Text = "Reset"
$ResetButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$ResetButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Reset User Parameters..."
	[System.Windows.Forms.Application]::DoEvents()
    ResetUserParameter
	GetNormalisationPolicy
	$StatusLabel.Text = ""
	
})
$mainForm.Controls.Add($ResetButton)

# SipUri
$SipUriLabel = New-Object System.Windows.Forms.Label
$SipUriLabel.Location = New-Object System.Drawing.Size(22,70) 
$SipUriLabel.Size = New-Object System.Drawing.Size(160,20)
$SipUriLabel.AutoSize = $true
$SipUriLabel.TabStop = $false
$SipUriLabel.Text = "SipUri:"
$mainForm.Controls.Add($SipUriLabel)
	
$SipUriLabelField = New-Object System.Windows.Forms.TextBox
$SipUriLabelField.Location = New-Object System.Drawing.Size(185,67) 
$SipUriLabelField.Size = New-Object System.Drawing.Size(150,20) 
$SipUriLabelField.Text = ""
#$SipUriLabelField.tabIndex = 3
$SipUriLabelField.TabStop = $false
$SipUriLabelField.Enabled = $false
$mainForm.Controls.Add($SipUriLabelField)


# IsForwardingEnabled
$IsForwardingEnabledLabel = New-Object System.Windows.Forms.Label
$IsForwardingEnabledLabel.Location = New-Object System.Drawing.Size(22,90) 
$IsForwardingEnabledLabel.Size = New-Object System.Drawing.Size(160,20)
$IsForwardingEnabledLabel.AutoSize = $true
$IsForwardingEnabledLabel.TabStop = $false
$IsForwardingEnabledLabel.Text = "IsForwardingEnabled:"
$mainForm.Controls.Add($IsForwardingEnabledLabel)

$IsForwardingEnabledDropDownBox = New-Object System.Windows.Forms.ComboBox
$IsForwardingEnabledDropDownBox.Location = New-Object System.Drawing.Size(185,87) 
$IsForwardingEnabledDropDownBox.Size = New-Object System.Drawing.Size(150,20) 
$IsForwardingEnabledDropDownBox.Text = ""
$IsForwardingEnabledDropDownBox.DropDownHeight = 200
$IsForwardingEnabledDropDownBox.DropDownWidth = 300
$IsForwardingEnabledDropDownBox.DropDownStyle = "DropDownList"
$IsForwardingEnabledDropDownBox.Items.Add("False") | Out-Null
$IsForwardingEnabledDropDownBox.Items.Add("True") | Out-Null
#$SipUriLabelField.tabIndex = 3
$IsForwardingEnabledDropDownBox.TabStop = $false
$IsForwardingEnabledDropDownBox.Enabled = $false
$mainForm.Controls.Add($IsForwardingEnabledDropDownBox)

# ForwardingType
$ForwardingTypeLabel = New-Object System.Windows.Forms.Label
$ForwardingTypeLabel.Location = New-Object System.Drawing.Size(22,110) 
$ForwardingTypeLabel.Size = New-Object System.Drawing.Size(160,20)
$ForwardingTypeLabel.AutoSize = $true
$ForwardingTypeLabel.TabStop = $false
$ForwardingTypeLabel.Text = "ForwardingType:"
$mainForm.Controls.Add($ForwardingTypeLabel)

$ForwardingTypeDropDownBox = New-Object System.Windows.Forms.ComboBox
$ForwardingTypeDropDownBox.Location = New-Object System.Drawing.Size(185,107) 
$ForwardingTypeDropDownBox.Size = New-Object System.Drawing.Size(150,20) 
$ForwardingTypeDropDownBox.Text = ""
$ForwardingTypeDropDownBox.DropDownHeight = 200
$ForwardingTypeDropDownBox.DropDownWidth = 300
$ForwardingTypeDropDownBox.DropDownStyle = "DropDownList"
$ForwardingTypeDropDownBox.Items.Add("Immediate") | Out-Null
$ForwardingTypeDropDownBox.Items.Add("Simultaneous") | Out-Null
#$SipUriLabelField.tabIndex = 3
$ForwardingTypeDropDownBox.TabStop = $false
$ForwardingTypeDropDownBox.Enabled = $false
$mainForm.Controls.Add($ForwardingTypeDropDownBox)

# ForwardingTarget
$ForwardingTargetLabel = New-Object System.Windows.Forms.Label
$ForwardingTargetLabel.Location = New-Object System.Drawing.Size(22,130) 
$ForwardingTargetLabel.Size = New-Object System.Drawing.Size(160,20)
$ForwardingTargetLabel.AutoSize = $true
$ForwardingTargetLabel.TabStop = $false
$ForwardingTargetLabel.Text = "ForwardingTarget:"
$mainForm.Controls.Add($ForwardingTargetLabel)
	
$ForwardingTargetTextBox = New-Object System.Windows.Forms.TextBox
$ForwardingTargetTextBox.Location = New-Object System.Drawing.Size(185,127) 
$ForwardingTargetTextBox.Size = New-Object System.Drawing.Size(150,20) 
$ForwardingTargetTextBox.Text = ""
#$ForwardingTargetTextBox.tabIndex = 3
$ForwardingTargetTextBox.TabStop = $false
$ForwardingTargetTextBox.Enabled = $false
$mainForm.Controls.Add($ForwardingTargetTextBox)

# ForwardingTargetType
$ForwardingTargetTypeLabel = New-Object System.Windows.Forms.Label
$ForwardingTargetTypeLabel.Location = New-Object System.Drawing.Size(22,150) 
$ForwardingTargetTypeLabel.Size = New-Object System.Drawing.Size(160,20)
$ForwardingTargetTypeLabel.AutoSize = $true
$ForwardingTargetTypeLabel.TabStop = $false
$ForwardingTargetTypeLabel.Text = "ForwardingTargetType:"
$mainForm.Controls.Add($ForwardingTargetTypeLabel)

$ForwardingTargetTypeDropDownBox = New-Object System.Windows.Forms.ComboBox
$ForwardingTargetTypeDropDownBox.Location = New-Object System.Drawing.Size(185,147) 
$ForwardingTargetTypeDropDownBox.Size = New-Object System.Drawing.Size(150,20) 
$ForwardingTargetTypeDropDownBox.Text = ""
$ForwardingTargetTypeDropDownBox.DropDownHeight = 200
$ForwardingTargetTypeDropDownBox.DropDownWidth = 300
$ForwardingTargetTypeDropDownBox.DropDownStyle = "DropDownList"
$ForwardingTargetTypeDropDownBox.Items.Add("Voicemail") | Out-Null
$ForwardingTargetTypeDropDownBox.Items.Add("SingleTarget") | Out-Null
$ForwardingTargetTypeDropDownBox.Items.Add("MyDelegates") | Out-Null
$ForwardingTargetTypeDropDownBox.Items.Add("Group") | Out-Null
#$SipUriLabelField.tabIndex = 3
$ForwardingTargetTypeDropDownBox.TabStop = $false
$ForwardingTargetTypeDropDownBox.Enabled = $false
$mainForm.Controls.Add($ForwardingTargetTypeDropDownBox)

# IsUnansweredEnabled
$IsUnansweredEnabledLabel = New-Object System.Windows.Forms.Label
$IsUnansweredEnabledLabel.Location = New-Object System.Drawing.Size(22,170) 
$IsUnansweredEnabledLabel.Size = New-Object System.Drawing.Size(160,20)
$IsUnansweredEnabledLabel.AutoSize = $true
$IsUnansweredEnabledLabel.TabStop = $false
$IsUnansweredEnabledLabel.Text = "IsUnansweredEnabled:"
$mainForm.Controls.Add($IsUnansweredEnabledLabel)

$IsUnansweredEnabledDropDownBox = New-Object System.Windows.Forms.ComboBox
$IsUnansweredEnabledDropDownBox.Location = New-Object System.Drawing.Size(185,167) 
$IsUnansweredEnabledDropDownBox.Size = New-Object System.Drawing.Size(150,20) 
$IsUnansweredEnabledDropDownBox.Text = ""
$IsUnansweredEnabledDropDownBox.DropDownHeight = 200
$IsUnansweredEnabledDropDownBox.DropDownWidth = 300
$IsUnansweredEnabledDropDownBox.DropDownStyle = "DropDownList"
$IsUnansweredEnabledDropDownBox.Items.Add("False") | Out-Null
$IsUnansweredEnabledDropDownBox.Items.Add("True") | Out-Null
#$SipUriLabelField.tabIndex = 3
$IsUnansweredEnabledDropDownBox.TabStop = $false
$IsUnansweredEnabledDropDownBox.Enabled = $false
$mainForm.Controls.Add($IsUnansweredEnabledDropDownBox)

# UnansweredTarget
$UnansweredTargetLabel = New-Object System.Windows.Forms.Label
$UnansweredTargetLabel.Location = New-Object System.Drawing.Size(22,190) 
$UnansweredTargetLabel.Size = New-Object System.Drawing.Size(160,20)
$UnansweredTargetLabel.AutoSize = $true
$UnansweredTargetLabel.TabStop = $false
$UnansweredTargetLabel.Text = "UnansweredTarget:"
$mainForm.Controls.Add($UnansweredTargetLabel)
	
$UnansweredTargetTextBox = New-Object System.Windows.Forms.TextBox
$UnansweredTargetTextBox.Location = New-Object System.Drawing.Size(185,187) 
$UnansweredTargetTextBox.Size = New-Object System.Drawing.Size(150,20) 
$UnansweredTargetTextBox.Text = ""
#$UnansweredTargetTextBox.tabIndex = 3
$UnansweredTargetTextBox.TabStop = $false
$UnansweredTargetTextBox.Enabled = $false
$mainForm.Controls.Add($UnansweredTargetTextBox)

# UnansweredTargetType
$UnansweredTargetTypeLabel = New-Object System.Windows.Forms.Label
$UnansweredTargetTypeLabel.Location = New-Object System.Drawing.Size(22,210) 
$UnansweredTargetTypeLabel.Size = New-Object System.Drawing.Size(160,20)
$UnansweredTargetTypeLabel.AutoSize = $true
$UnansweredTargetTypeLabel.TabStop = $false
$UnansweredTargetTypeLabel.Text = "UnansweredTargetType:"
$mainForm.Controls.Add($UnansweredTargetTypeLabel)

$UnansweredTargetTypeDropDownBox = New-Object System.Windows.Forms.ComboBox
$UnansweredTargetTypeDropDownBox.Location = New-Object System.Drawing.Size(185,207) 
$UnansweredTargetTypeDropDownBox.Size = New-Object System.Drawing.Size(150,20) 
$UnansweredTargetTypeDropDownBox.Text = ""
$UnansweredTargetTypeDropDownBox.DropDownHeight = 200
$UnansweredTargetTypeDropDownBox.DropDownWidth = 300
$UnansweredTargetTypeDropDownBox.DropDownStyle = "DropDownList"
$UnansweredTargetTypeDropDownBox.Items.Add("Voicemail") | Out-Null
$UnansweredTargetTypeDropDownBox.Items.Add("SingleTarget") | Out-Null
$UnansweredTargetTypeDropDownBox.Items.Add("MyDelegates") | Out-Null
$UnansweredTargetTypeDropDownBox.Items.Add("Group") | Out-Null
#$SipUriLabelField.tabIndex = 3
$UnansweredTargetTypeDropDownBox.TabStop = $false
$UnansweredTargetTypeDropDownBox.Enabled = $false
$mainForm.Controls.Add($UnansweredTargetTypeDropDownBox)

# UnansweredDelay
$UnansweredDelayLabel = New-Object System.Windows.Forms.Label
$UnansweredDelayLabel.Location = New-Object System.Drawing.Size(22,230) 
$UnansweredDelayLabel.Size = New-Object System.Drawing.Size(160,20)
$UnansweredDelayLabel.AutoSize = $true
$UnansweredDelayLabel.TabStop = $false
$UnansweredDelayLabel.Text = "UnansweredDelay:"
$mainForm.Controls.Add($UnansweredDelayLabel)
	
$UnansweredDelayTextBox = New-Object System.Windows.Forms.TextBox
$UnansweredDelayTextBox.Location = New-Object System.Drawing.Size(185,227) 
$UnansweredDelayTextBox.Size = New-Object System.Drawing.Size(150,20) 
$UnansweredDelayTextBox.Text = ""
#$UnansweredDelayTextBox.tabIndex = 3
$UnansweredDelayTextBox.TabStop = $false
$UnansweredDelayTextBox.Enabled = $false
$mainForm.Controls.Add($UnansweredDelayTextBox)




<#

#NameTextLabel Label ============================================================
$NameTextLabel = New-Object System.Windows.Forms.Label
$NameTextLabel.Location = New-Object System.Drawing.Size(10,10) 
$NameTextLabel.Size = New-Object System.Drawing.Size(60,15) 
$NameTextLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$NameTextLabel.Text = "Name:"
$NameTextLabel.TabStop = $false
$NameTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.Controls.Add($NameTextLabel)


#Name Text box ============================================================
$NameTextBox = New-Object System.Windows.Forms.TextBox
$NameTextBox.location = new-object system.drawing.size(72,10)
$NameTextBox.size = new-object system.drawing.size(250,23)
$NameTextBox.tabIndex = 1
$NameTextBox.text = ""   
$NameTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.controls.add($NameTextBox)
$NameTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{	
		#Do Nothing
	}
})


#Description Label ============================================================
$DescriptionTextLabel = New-Object System.Windows.Forms.Label
$DescriptionTextLabel.Location = New-Object System.Drawing.Size(5,35) 
$DescriptionTextLabel.Size = New-Object System.Drawing.Size(65,15) 
$DescriptionTextLabel.Text = "Description:"
$DescriptionTextLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$DescriptionTextLabel.TabStop = $false
$DescriptionTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.Controls.Add($DescriptionTextLabel)

#$DescriptionTextBox Text box ============================================================
$DescriptionTextBox = New-Object System.Windows.Forms.TextBox
$DescriptionTextBox.location = new-object system.drawing.size(72,35)
$DescriptionTextBox.size = new-object system.drawing.size(250,23)
$DescriptionTextBox.tabIndex = 1
$DescriptionTextBox.text = ""   
$DescriptionTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.controls.add($DescriptionTextBox)
$DescriptionTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{	
		#AddSetting
	}
})


#Pattern Label ============================================================
$PatternTextLabel = New-Object System.Windows.Forms.Label
$PatternTextLabel.Location = New-Object System.Drawing.Size(5,60) 
$PatternTextLabel.Size = New-Object System.Drawing.Size(65,15) 
$PatternTextLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$PatternTextLabel.Text = "Pattern:"
$PatternTextLabel.TabStop = $false
$PatternTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.Controls.Add($PatternTextLabel)

#Pattern Text box ============================================================
$PatternTextBox = New-Object System.Windows.Forms.TextBox
$PatternTextBox.location = new-object system.drawing.size(72,60)
$PatternTextBox.size = new-object system.drawing.size(250,23)
$PatternTextBox.tabIndex = 1
$PatternTextBox.text = ""   
$PatternTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.controls.add($PatternTextBox)
$PatternTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{	
		#AddSetting
	}
})


#Translation Label ============================================================
$TranslationTextLabel = New-Object System.Windows.Forms.Label
$TranslationTextLabel.Location = New-Object System.Drawing.Size(5,85) 
$TranslationTextLabel.Size = New-Object System.Drawing.Size(65,15) 
$TranslationTextLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$TranslationTextLabel.Text = "Translation:"
$TranslationTextLabel.TabStop = $false
$TranslationTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.Controls.Add($TranslationTextLabel)

#Setting Text box ============================================================
$TranslationTextBox = New-Object System.Windows.Forms.TextBox
$TranslationTextBox.location = new-object system.drawing.size(72,85)
$TranslationTextBox.size = new-object system.drawing.size(250,23)
$TranslationTextBox.tabIndex = 1
$TranslationTextBox.text = ""   
$TranslationTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.controls.add($TranslationTextBox )
$TranslationTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{
		$StatusLabel.Text = "STATUS: Adding normalization rule..."
		[System.Windows.Forms.Application]::DoEvents()
		AddSetting
		$StatusLabel.Text = ""
	}
})

#ExtensionTextLabel Label ============================================================
$ExtensionTextLabel = New-Object System.Windows.Forms.Label
$ExtensionTextLabel.Location = New-Object System.Drawing.Size(5,110) 
$ExtensionTextLabel.Size = New-Object System.Drawing.Size(65,15)
$ExtensionTextLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight 
$ExtensionTextLabel.Text = "Extension:"
$ExtensionTextLabel.TabStop = $false
$ExtensionTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
#$mainForm.Controls.Add($ExtensionTextLabel)

$ExtensionCheckBox = New-Object System.Windows.Forms.Checkbox 
$ExtensionCheckBox.Location = New-Object System.Drawing.Size(72,110) 
$ExtensionCheckBox.Size = New-Object System.Drawing.Size(20,20)
$ExtensionCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$ExtensionCheckBox.tabIndex = 2
$ExtensionCheckBox.Add_Click(
{
	
})
#$mainForm.controls.add($ExtensionCheckBox)

#Add button
$AddButton = New-Object System.Windows.Forms.Button
$AddButton.Location = New-Object System.Drawing.Size(340,20)
$AddButton.Size = New-Object System.Drawing.Size(70,18)
$AddButton.Text = "Add / Edit"
$AddButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$AddButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Adding normalization rule..."
	[System.Windows.Forms.Application]::DoEvents()
	AddSetting
	$StatusLabel.Text = ""
})
#$mainForm.Controls.Add($AddButton)


#Delete button
$DeleteButton = New-Object System.Windows.Forms.Button
$DeleteButton.Location = New-Object System.Drawing.Size(340,45)
$DeleteButton.Size = New-Object System.Drawing.Size(70,18)
$DeleteButton.Text = "Delete"
$DeleteButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$DeleteButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Deleting normalization rule..."
	[System.Windows.Forms.Application]::DoEvents()
	DeleteSetting
	$StatusLabel.Text = ""
})
#$mainForm.Controls.Add($DeleteButton)


#Add button
$DeleteAllButton = New-Object System.Windows.Forms.Button
$DeleteAllButton.Location = New-Object System.Drawing.Size(340,70)
$DeleteAllButton.Size = New-Object System.Drawing.Size(70,18)
$DeleteAllButton.Text = "Delete All"
$DeleteAllButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$DeleteAllButton.Add_Click(
{
	$StatusLabel.Text = "STATUS: Deleting all normalization rules..."
	[System.Windows.Forms.Application]::DoEvents()
	$a = new-object -comobject wscript.shell
	$intAnswer = $a.popup("Are you sure you want to remove all the rules from this Dial Plan?",0,"Remove All Rules",4) 
	if ($intAnswer -eq 6) { 
					
		DeleteAllSettings
	}
	$StatusLabel.Text = ""
	
})
#$mainForm.Controls.Add($DeleteAllButton)


$GroupBoxNormRule = New-Object System.Windows.Forms.Panel
$GroupBoxNormRule.Location = New-Object System.Drawing.Size(20,357) 
$GroupBoxNormRule.Size = New-Object System.Drawing.Size(420,135) 
$GroupBoxNormRule.MinimumSize = New-Object System.Drawing.Size(400,80) 
$GroupBoxNormRule.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
$GroupBoxNormRule.TabStop = $False
$GroupBoxNormRule.Controls.Add($NameTextLabel)
$GroupBoxNormRule.Controls.Add($NameTextBox)
$GroupBoxNormRule.Controls.Add($DescriptionTextLabel)
$GroupBoxNormRule.Controls.Add($DescriptionTextBox)
$GroupBoxNormRule.Controls.Add($PatternTextLabel)
$GroupBoxNormRule.Controls.Add($PatternTextBox)
$GroupBoxNormRule.Controls.Add($TranslationTextLabel)
$GroupBoxNormRule.Controls.Add($TranslationTextBox)
$GroupBoxNormRule.Controls.Add($ExtensionTextLabel)
$GroupBoxNormRule.Controls.Add($ExtensionCheckBox)
$GroupBoxNormRule.Controls.Add($AddButton)
$GroupBoxNormRule.Controls.Add($DeleteButton)
$GroupBoxNormRule.Controls.Add($DeleteAllButton)
$GroupBoxNormRule.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$GroupBoxNormRule.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$mainForm.Controls.Add($GroupBoxNormRule)


#>

<#
#Import Label ============================================================
$ImportTextLabel = New-Object System.Windows.Forms.Label
$ImportTextLabel.Location = New-Object System.Drawing.Size(50,578) 
$ImportTextLabel.Size = New-Object System.Drawing.Size(80,15) 
$ImportTextLabel.Text = "Import/Export:"
$ImportTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$ImportTextLabel.TabStop = $false
$mainForm.Controls.Add($ImportTextLabel)


#Import button
$ImportButton = New-Object System.Windows.Forms.Button
$ImportButton.Location = New-Object System.Drawing.Size(130,575)
$ImportButton.Size = New-Object System.Drawing.Size(120,20)
$ImportButton.Text = "Import Config"
$ImportButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$ImportButton.Add_Click(
{
	Import-Config
	UpdateListViewSettings
	
})
$mainForm.Controls.Add($ImportButton)


#Export button
$ExportButton = New-Object System.Windows.Forms.Button
$ExportButton.Location = New-Object System.Drawing.Size(260,575)
$ExportButton.Size = New-Object System.Drawing.Size(120,20)
$ExportButton.Text = "Export Config"
$ExportButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$ExportButton.Add_Click(
{
	Export-Config

})
$mainForm.Controls.Add($ExportButton)
#>


# Add the Status Label ============================================================
$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Location = New-Object System.Drawing.Size(15,620) 
$StatusLabel.Size = New-Object System.Drawing.Size(420,15) 
$StatusLabel.Text = ""
$StatusLabel.forecolor = "DarkBlue"
$StatusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$StatusLabel.TabStop = $false
$mainForm.Controls.Add($StatusLabel)

# First compile list
<#
CompileUserList

$numberOfItems = $UserDropDownBox.Items.count
if($numberOfItems -gt 0)
{
	$UserDropDownBox.SelectedIndex = 0
}

GetNormalisationPolicy
				
if($currentIndex -ne $null)
{
	if($currentIndex -lt $dgv.Rows.Count)
	{$dgv.Rows[$currentIndex].Selected = $True}
}
				
EnableAllButtons
#>
<#
$ToolTip = New-Object System.Windows.Forms.ToolTip 
$ToolTip.BackColor = [System.Drawing.Color]::LightGoldenrodYellow 
$ToolTip.IsBalloon = $true 
$ToolTip.InitialDelay = 500 
$ToolTip.ReshowDelay = 500 
$ToolTip.AutoPopDelay = 10000
#$ToolTip.ToolTipTitle = "Help:"
$ToolTip.SetToolTip($AddButton, "If the specified Name is the same as an existing rule`r`nthen than rule will be edited. If the Name is new`r`nthen a new rule will be created.") 
$ToolTip.SetToolTip($DeleteButton, "The Delete button will delete the selected rule(s).") 
$ToolTip.SetToolTip($DeleteAllButton, "The Delete All button will delete all the rules in this Dial Plan.") 
$ToolTip.SetToolTip($OptimizeDeviceDialingLabel, "Indicates whether Access Prefix`r`nis being applied by the system.")
#>


function CheckTeamsOnlineInitial
{	
	#CHECK IF COMMANDS ARE AVAILABLE		
	$command = "Get-CsOnlineUser"
	#if($CurrentlyConnected -and (Get-Command $command -errorAction SilentlyContinue) -and ($Script:UserConnectedToTeamsOnline -eq $true))
	if((Get-Command $command -errorAction SilentlyContinue))
	{
		$isConnected = $false
		try{
			(Get-CsTenantDialplan -ErrorAction SilentlyContinue) 2> $null
			$isConnected = $true
		}
		catch
		{
			Write-Host "ERROR: " $_ -foreground "red"
			$isConnected = $false
		}
		#CHECK THAT SfB ONLINE COMMANDS WORK
		if($isConnected)
		{
			#Write-Host "Connected to Teams" -foreground "Green"
			$ConnectedOnlineLabel.Visible = $true
			$StatusLabel.Text = ""

			CompileUserList
	
			$numberOfItems = $UserDropDownBox.Items.count
			if($numberOfItems -gt 0)
			{
				$UserDropDownBox.SelectedIndex = 0
			}
			
			
			if($currentIndex -ne $null)
			{
				if($currentIndex -lt $dgv.Rows.Count)
				{$dgv.Rows[$currentIndex].Selected = $True}
			}
			
			EnableAllButtons
			
			return $true
		}
		else
		{
			Write-Host "INFO: Cannot access Teams. Please use the Connect Teams button." -foreground "Yellow"
			$ConnectedOnlineLabel.Visible = $false
			$StatusLabel.Text = "Press the `"Connect Teams`" button to get started."
			
			DisableAllButtons
		}
	}
}


function CheckTeamsOnline
{	
	
	#CHECK IF COMMANDS ARE AVAILABLE		
	$isConnected = $false
	try{
		(Get-CsOnlineUser -ResultSize 1 -ErrorAction SilentlyContinue) 2> $null
		$isConnected = $true
	}
	catch
	{
		#Write-Host "ERROR: " $_ -foreground "red"
		$isConnected = $false
	}
	#CHECK THAT SfB ONLINE COMMANDS WORK
	if($isConnected)
	{
		$ConnectedOnlineLabel.Visible = $true
		
		#$StatusLabel.Text = ""
		return $true
		
	}
	else
	{
		Write-Host "INFO: Cannot access Teams. Please use the Connect Teams button." -foreground "Yellow"
		$ConnectedOnlineLabel.Visible = $false
		
		
		
		DisableAllButtons
	}
}



function DisableAllButtons()
{
	$UserDropDownBox.Enabled = $false
    $ApplyButton.Enabled = $false
    $SipUriLabelField.Enabled = $false
    $IsForwardingEnabledDropDownBox.Enabled = $false
    $ForwardingTypeDropDownBox.Enabled = $false
    $ForwardingTargetTextBox.Enabled = $false
    $ForwardingTargetTypeDropDownBox.Enabled = $false
    $IsUnansweredEnabledDropDownBox.Enabled = $false
    $UnansweredTargetTextBox.Enabled = $false
    $UnansweredTargetTypeDropDownBox.Enabled = $false
    $UnansweredDelayTextBox.Enabled = $false
    <#
    $UserDropDownBox.Enabled = $false
	$NewPolicyButton.Enabled = $false
	$RemovePolicyButton.Enabled = $false
	$UpButton.Enabled = $false
	$DownButton.Enabled = $false
	$NameTextBox.Enabled = $false
	$DescriptionTextBox.Enabled = $false
	$PatternTextBox.Enabled = $false
	$TranslationTextBox.Enabled = $false
	$AddButton.Enabled = $false
	$DeleteButton.Enabled = $false
	$DeleteAllButton.Enabled = $false
	$TestPhoneTextBox.Enabled = $false
	$TestPhoneButton.Enabled = $false
	$EditPolicyButton.Enabled = $false
	$ExtensionCheckBox.Enabled = $false
    #>
}


function EnableAllButtons()
{
	$UserDropDownBox.Enabled = $true
    $ApplyButton.Enabled = $true
    $SipUriLabelField.Enabled = $true
    $IsForwardingEnabledDropDownBox.Enabled = $true
    $ForwardingTypeDropDownBox.Enabled = $true
    $ForwardingTargetTextBox.Enabled = $true
    $ForwardingTargetTypeDropDownBox.Enabled = $true
    $IsUnansweredEnabledDropDownBox.Enabled = $true
    $UnansweredTargetTextBox.Enabled = $true
    $UnansweredTargetTypeDropDownBox.Enabled = $true
    $UnansweredDelayTextBox.Enabled = $true
    <#
    $UserDropDownBox.Enabled = $true
	$NewPolicyButton.Enabled = $true
	$RemovePolicyButton.Enabled = $true
	$UpButton.Enabled = $true
	$DownButton.Enabled = $true
	$NameTextBox.Enabled = $true
	$DescriptionTextBox.Enabled = $true
	$PatternTextBox.Enabled = $true
	$TranslationTextBox.Enabled = $true
	$AddButton.Enabled = $true
	$DeleteButton.Enabled = $true
	$DeleteAllButton.Enabled = $true
	$TestPhoneTextBox.Enabled = $true
	$TestPhoneButton.Enabled = $true
	$EditPolicyButton.Enabled = $true
	$ExtensionCheckBox.Enabled = $true
    #>
}
<#
function New-Policy([string]$Message, [string]$WindowTitle, [string]$DefaultText)
{
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
     
    # Create the Label
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Size(10,10) 
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.AutoSize = $true
    $label.Text = $Message
     	
	$PolicyTextBox = New-Object System.Windows.Forms.TextBox
	$PolicyTextBox.Location = New-Object System.Drawing.Size(10,30) 
	$PolicyTextBox.Size = New-Object System.Drawing.Size(300,20) 
	$PolicyTextBox.Text = "<Enter Dial Plan Name>"
	$PolicyTextBox.tabIndex = 1

	$label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Size(10,60) 
    $label2.Size = New-Object System.Drawing.Size(280,20)
    $label2.AutoSize = $true
    $label2.Text = "Copy Normalization Rules from existing Dial Plan:"
	
	# PoliciesDropDownBox ============================================================
	$PoliciesDropDownBox = New-Object System.Windows.Forms.ComboBox 
	$PoliciesDropDownBox.Location = New-Object System.Drawing.Size(10,80) 
	$PoliciesDropDownBox.Size = New-Object System.Drawing.Size(280,20) 
	$PoliciesDropDownBox.DropDownHeight = 200 
	$PoliciesDropDownBox.tabIndex = 1
	$PoliciesDropDownBox.Sorted = $true
	$PoliciesDropDownBox.DropDownStyle = "DropDownList"
	$PoliciesDropDownBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		Get-CsTenantDialPlan | select-object identity | ForEach-Object {[void] $PoliciesDropDownBox.Items.Add(($_.identity).Replace("Tag:",""))}
	}
	
	if($PoliciesDropDownBox.Items.Count -ge 0)
	{
		$PoliciesDropDownBox.SelectedIndex = 0
	}
	$PoliciesDropDownBox.Enabled = $false
	
	$CopyCheckBox = New-Object System.Windows.Forms.Checkbox 
	$CopyCheckBox.Location = New-Object System.Drawing.Size(295,80) 
	$CopyCheckBox.Size = New-Object System.Drawing.Size(20,20)
	$CopyCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	$CopyCheckBox.tabIndex = 2
	$CopyCheckBox.Add_Click(
	{
		if($CopyCheckBox.Checked -eq $false)
		{
			#$PolicyTextBox.Text = "<Enter Policy Name>"
			#$PolicyTextBox.Enabled = $true
			$PoliciesDropDownBox.Enabled = $false
		}
		else
		{
			#$PolicyTextBox.Text = ""
			#$PolicyTextBox.Enabled = $false
			$PoliciesDropDownBox.Enabled = $true
		}
	})
	
	# Create the Label
    $label3 = New-Object System.Windows.Forms.Label
    $label3.Location = New-Object System.Drawing.Size(10,113) 
    $label3.Size = New-Object System.Drawing.Size(90,20)
    $label3.AutoSize = $true
    $label3.Text = "Access Prefix:"
     	
	$AccessPrefixTextBox = New-Object System.Windows.Forms.TextBox
	$AccessPrefixTextBox.Location = New-Object System.Drawing.Size(95,110) 
	$AccessPrefixTextBox.Size = New-Object System.Drawing.Size(50,20) 
	$AccessPrefixTextBox.Text = ""
	$AccessPrefixTextBox.tabIndex = 3
	
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(150,150)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
    $okButton.Add_Click({ 
	if($CopyCheckBox.Checked -eq $false)
	{
		$Result = New-Object PSObject -Property @{
		  NewPolicy = $PolicyTextBox.Text.ToString()
		  ExistingChecked = $false
		  ExistingPolicy = $PoliciesDropDownBox.Text.ToString()
		  AccessPrefix = $AccessPrefixTextBox.Text.ToString()
		}
		$form.Tag = $Result
	}
	else
	{
		$Result = New-Object PSObject -Property @{
		  NewPolicy = $PolicyTextBox.Text.ToString()
		  ExistingChecked = $true
		  ExistingPolicy = $PoliciesDropDownBox.Text.ToString()
		  AccessPrefix = $AccessPrefixTextBox.Text.ToString()
		}
		$form.Tag = $Result
	}
	$form.Close() 
	})
     
    # Create the Cancel button.
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Size(240,150)
    $cancelButton.Size = New-Object System.Drawing.Size(75,25)
    $cancelButton.Text = "Cancel"
    $cancelButton.Add_Click({ $form.Tag = $null; $form.Close() })
     
    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = $WindowTitle
    $form.Size = New-Object System.Drawing.Size(350,220)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    $form.ShowInTaskbar = $true
	[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
     
    # Add all of the controls to the form.
    $form.Controls.Add($label)
    $form.Controls.Add($PolicyTextBox)
    $form.Controls.Add($okButton)
    $form.Controls.Add($cancelButton)
	$form.Controls.Add($label2)
	$form.Controls.Add($PoliciesDropDownBox)
	$form.Controls.Add($CopyCheckBox)
	$form.Controls.Add($label3)
	$form.Controls.Add($AccessPrefixTextBox)

     
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
     
    # Return the text that the user entered.
    return $form.Tag
}
#>
<#
function Edit-Policy([string]$Message, [string]$WindowTitle, [string]$PolicyName)
{
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
     
    # Create the Label
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Size(10,10) 
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.AutoSize = $true
    $label.Text = $Message
     	
	$PolicyTextBox = New-Object System.Windows.Forms.TextBox
	$PolicyTextBox.Location = New-Object System.Drawing.Size(10,30) 
	$PolicyTextBox.Size = New-Object System.Drawing.Size(300,20) 
	$PolicyTextBox.Text = "$PolicyName"
	$PolicyTextBox.tabIndex = 1
	$PolicyTextBox.Enabled = $false
	
	
	$label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Size(10,60) 
    $label2.Size = New-Object System.Drawing.Size(280,20)
    $label2.AutoSize = $true
    $label2.Text = "Overwrite Normalization Rules from an existing Dial Plan:"
	
	# PoliciesDropDownBox ============================================================
	$PoliciesDropDownBox = New-Object System.Windows.Forms.ComboBox 
	$PoliciesDropDownBox.Location = New-Object System.Drawing.Size(10,80) 
	$PoliciesDropDownBox.Size = New-Object System.Drawing.Size(280,20) 
	$PoliciesDropDownBox.DropDownHeight = 200 
	$PoliciesDropDownBox.tabIndex = 1
	$PoliciesDropDownBox.Sorted = $true
	$PoliciesDropDownBox.DropDownStyle = "DropDownList"
	$PoliciesDropDownBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
	
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		Get-CsTenantDialPlan | Select-Object identity | ForEach-Object {[void] $PoliciesDropDownBox.Items.Add(($_.identity).Replace("Tag:",""))}
	}
	
	if($PoliciesDropDownBox.Items.Count -ge 0)
	{
		$PoliciesDropDownBox.SelectedIndex = 0
	}
	$PoliciesDropDownBox.Enabled = $false
	
	$CopyCheckBox = New-Object System.Windows.Forms.Checkbox 
	$CopyCheckBox.Location = New-Object System.Drawing.Size(295,80) 
	$CopyCheckBox.Size = New-Object System.Drawing.Size(20,20)
	$CopyCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
	$CopyCheckBox.tabIndex = 2
	$CopyCheckBox.Add_Click(
	{
		if($CopyCheckBox.Checked -eq $false)
		{
			#$PolicyTextBox.Text = "<Enter Policy Name>"
			#$PolicyTextBox.Enabled = $true
			$PoliciesDropDownBox.Enabled = $false
		}
		else
		{
			#$PolicyTextBox.Text = ""
			#$PolicyTextBox.Enabled = $false
			$PoliciesDropDownBox.Enabled = $true
		}
	})
	
	# Create the Label
    $label3 = New-Object System.Windows.Forms.Label
    $label3.Location = New-Object System.Drawing.Size(10,113) 
    $label3.Size = New-Object System.Drawing.Size(90,20)
    $label3.AutoSize = $true
    $label3.Text = "Access Prefix:"
     	
	$AccessPrefixTextBox = New-Object System.Windows.Forms.TextBox
	$AccessPrefixTextBox.Location = New-Object System.Drawing.Size(95,110) 
	$AccessPrefixTextBox.Size = New-Object System.Drawing.Size(50,20) 
	$AccessPrefixTextBox.Text = ""
	$AccessPrefixTextBox.tabIndex = 3
	
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		Get-CsTenantDialPlan -identity $PolicyName | ForEach-Object {$AccessPrefixTextBox.Text = $_.ExternalAccessPrefix; $Optimized = $_.OptimizeDeviceDialing; $OptimizeDeviceDialingLabel.Text = "OptimizeDeviceDialing: $Optimized"}
	}
	
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(150,150)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
    $okButton.Add_Click({ 
	if($CopyCheckBox.Checked -eq $false)
	{
		$Result = New-Object PSObject -Property @{
		  NewPolicy = $PolicyTextBox.Text.ToString()
		  ExistingChecked = $false
		  ExistingPolicy = $PoliciesDropDownBox.Text.ToString()
		  AccessPrefix = $AccessPrefixTextBox.Text.ToString()
		}
		$form.Tag = $Result
	}
	else
	{
		$Result = New-Object PSObject -Property @{
		  NewPolicy = $PolicyTextBox.Text.ToString()
		  ExistingChecked = $true
		  ExistingPolicy = $PoliciesDropDownBox.Text.ToString()
		  AccessPrefix = $AccessPrefixTextBox.Text.ToString()
		}
		$form.Tag = $Result
	}
	$form.Close() 
	})
     
    # Create the Cancel button.
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Size(240,150)
    $cancelButton.Size = New-Object System.Drawing.Size(75,25)
    $cancelButton.Text = "Cancel"
    $cancelButton.Add_Click({ $form.Tag = $null; $form.Close() })
     
    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = $WindowTitle
    $form.Size = New-Object System.Drawing.Size(350,220)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    $form.ShowInTaskbar = $true
	[byte[]]$WindowIcon = @(71, 73, 70, 56, 57, 97, 32, 0, 32, 0, 231, 137, 0, 0, 52, 93, 0, 52, 94, 0, 52, 95, 0, 53, 93, 0, 53, 94, 0, 53, 95, 0,53, 96, 0, 54, 94, 0, 54, 95, 0, 54, 96, 2, 54, 95, 0, 55, 95, 1, 55, 96, 1, 55, 97, 6, 55, 96, 3, 56, 98, 7, 55, 96, 8, 55, 97, 9, 56, 102, 15, 57, 98, 17, 58, 98, 27, 61, 99, 27, 61, 100, 24, 61, 116, 32, 63, 100, 36, 65, 102, 37, 66, 103, 41, 68, 104, 48, 72, 106, 52, 75, 108, 55, 77, 108, 57, 78, 109, 58, 79, 111, 59, 79, 110, 64, 83, 114, 65, 83, 114, 68, 85, 116, 69, 86, 117, 71, 88, 116, 75, 91, 120, 81, 95, 123, 86, 99, 126, 88, 101, 125, 89, 102, 126, 90, 103, 129, 92, 103, 130, 95, 107, 132, 97, 108, 132, 99, 110, 134, 100, 111, 135, 102, 113, 136, 104, 114, 137, 106, 116, 137, 106,116, 139, 107, 116, 139, 110, 119, 139, 112, 121, 143, 116, 124, 145, 120, 128, 147, 121, 129, 148, 124, 132, 150, 125,133, 151, 126, 134, 152, 127, 134, 152, 128, 135, 152, 130, 137, 154, 131, 138, 155, 133, 140, 157, 134, 141, 158, 135,141, 158, 140, 146, 161, 143, 149, 164, 147, 152, 167, 148, 153, 168, 151, 156, 171, 153, 158, 172, 153, 158, 173, 156,160, 174, 156, 161, 174, 158, 163, 176, 159, 163, 176, 160, 165, 177, 163, 167, 180, 166, 170, 182, 170, 174, 186, 171,175, 186, 173, 176, 187, 173, 177, 187, 174, 178, 189, 176, 180, 190, 177, 181, 191, 179, 182, 192, 180, 183, 193, 182,185, 196, 185, 188, 197, 188, 191, 200, 190, 193, 201, 193, 195, 203, 193, 196, 204, 196, 198, 206, 196, 199, 207, 197,200, 207, 197, 200, 208, 198, 200, 208, 199, 201, 208, 199, 201, 209, 200, 202, 209, 200, 202, 210, 202, 204, 212, 204,206, 214, 206, 208, 215, 206, 208, 216, 208, 210, 218, 209, 210, 217, 209, 210, 220, 209, 211, 218, 210, 211, 219, 210,211, 220, 210, 212, 219, 211, 212, 219, 211, 212, 220, 212, 213, 221, 214, 215, 223, 215, 216, 223, 215, 216, 224, 216,217, 224, 217, 218, 225, 218, 219, 226, 218, 220, 226, 219, 220, 226, 219, 220, 227, 220, 221, 227, 221, 223, 228, 224,225, 231, 228, 229, 234, 230, 231, 235, 251, 251, 252, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,255, 255, 255, 255, 255, 255, 255, 255, 33, 254, 17, 67, 114, 101, 97, 116, 101, 100, 32, 119, 105, 116, 104, 32, 71, 73, 77, 80, 0, 33, 249, 4, 1, 10, 0, 255, 0, 44, 0, 0, 0, 0, 32, 0, 32, 0, 0, 8, 254, 0, 255, 29, 24, 72, 176, 160, 193, 131, 8, 25, 60, 16, 120, 192, 195, 10, 132, 16, 35, 170, 248, 112, 160, 193, 64, 30, 135, 4, 68, 220, 72, 16, 128, 33, 32, 7, 22, 92, 68, 84, 132, 35, 71, 33, 136, 64, 18, 228, 81, 135, 206, 0, 147, 16, 7, 192, 145, 163, 242, 226, 26, 52, 53, 96, 34, 148, 161, 230, 76, 205, 3, 60, 214, 204, 72, 163, 243, 160, 25, 27, 62, 11, 6, 61, 96, 231, 68, 81, 130, 38, 240, 28, 72, 186, 114, 205, 129, 33, 94, 158, 14, 236, 66, 100, 234, 207, 165, 14, 254, 108, 120, 170, 193, 15, 4, 175, 74, 173, 30, 120, 50, 229, 169, 20, 40, 3, 169, 218, 28, 152, 33, 80, 2, 157, 6, 252, 100, 136, 251, 85, 237, 1, 46, 71,116, 26, 225, 66, 80, 46, 80, 191, 37, 244, 0, 48, 57, 32, 15, 137, 194, 125, 11, 150, 201, 97, 18, 7, 153, 130, 134, 151, 18, 140, 209, 198, 36, 27, 24, 152, 35, 23, 188, 147, 98, 35, 138, 56, 6, 51, 251, 29, 24, 4, 204, 198, 47, 63, 82, 139, 38, 168, 64, 80, 7, 136, 28, 250, 32, 144, 157, 246, 96, 19, 43, 16, 169, 44, 57, 168, 250, 32, 6, 66, 19, 14, 70, 248, 99, 129, 248, 236, 130, 90, 148, 28, 76, 130, 5, 97, 241, 131, 35, 254, 4, 40, 8, 128, 15, 8, 235, 207, 11, 88, 142, 233, 81, 112, 71, 24, 136, 215, 15, 190, 152, 67, 128, 224, 27, 22, 232, 195, 23, 180, 227, 98, 96, 11, 55, 17, 211, 31, 244, 49, 102, 160, 24, 29, 249, 201, 71, 80, 1, 131, 136, 16, 194, 30, 237, 197, 215, 91, 68, 76, 108, 145, 5, 18, 27, 233, 119, 80, 5, 133, 0, 66, 65, 132, 32, 73, 48, 16, 13, 87, 112, 20, 133, 19, 28, 85, 113, 195, 1, 23, 48, 164, 85, 68, 18, 148, 24, 16, 0, 59)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
     
    # Add all of the controls to the form.
    $form.Controls.Add($label)
    $form.Controls.Add($PolicyTextBox)
    $form.Controls.Add($okButton)
    $form.Controls.Add($cancelButton)
	$form.Controls.Add($label2)
	$form.Controls.Add($PoliciesDropDownBox)
	$form.Controls.Add($CopyCheckBox)
	$form.Controls.Add($label3)
	$form.Controls.Add($AccessPrefixTextBox)

     
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
     
    # Return the text that the user entered.
    return $form.Tag
}
#>
<#
function Move-Up
{
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		foreach ($lvi in $lv.SelectedItems)
		{
			#GET SETTINGS OF SELECTED ITEM
			$item = $lv.Items[$lvi.Index]
			$itemValue = $item.SubItems

			[string]$Scope = $UserDropDownBox.SelectedItem.ToString()
			[string]$Name = $item.Text
			#[string]$Priority = $itemValue[1].Text
			[string]$Description = $itemValue[1].Text
			[string]$Pattern = $itemValue[2].Text
			[string]$Translation = $itemValue[3].Text
			[bool]$ExtensionValue = $itemValue[4].Text
							
			$orgIndex = $lvi.Index
			if($orgIndex -gt 0)
			{
				$index = $orgIndex - 1
					
				Write-Host "RUNNING: `$nr = New-CsVoiceNormalizationRule -Identity "${Scope}/${Name}" -Description $Description -Pattern $Pattern -Translation $Translation -Priority $index -IsInternalExtension $ExtensionValue -InMemory" -foreground "green"
				$nr = New-CsVoiceNormalizationRule -Identity "${Scope}/${Name}" -Description $Description -Pattern $Pattern -Translation $Translation -Priority $index -IsInternalExtension $ExtensionValue -InMemory			
				
				$NormArray = (Get-CsTenantDialPlan -identity $Scope | select NormalizationRules).NormalizationRules
				Write-Verbose "INITIAL ARRAY"
				foreach($item in $NormArray){Write-Verbose $item}
				Write-Host
				#Remove Item
				$NormArray.RemoveAt($orgIndex)
				Write-Verbose "AFTER DELETE"
				foreach($item in $NormArray){Write-Verbose $item}
				Write-Host 
				#Insert Item
				$NormArray.Insert($index,$nr)
				Write-Verbose "AFTER INSERT"
				foreach($item in $NormArray){Write-Verbose $item}
				
				Set-CsTenantDialPlan -Identity $Scope -NormalizationRules @{Replace=$NormArray}	
				
				GetNormalisationPolicy
				
				$lv.Items[$index].Selected = $true
				$lv.Items[$index].EnsureVisible()
			}
			else
			{
				Write-Host "INFO: Cannot move item any higher..." -foreground "Yellow"
			}
		}
	}

}
#>
<#
function Move-Down
{
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		foreach ($lvi in $lv.SelectedItems)
		{
			#GET SETTINGS OF SELECTED ITEM
			$item = $lv.Items[$lvi.Index]
			$itemValue = $item.SubItems

			[string]$Scope = $UserDropDownBox.SelectedItem.ToString()
			[string]$Name = $item.Text
			#[string]$Priority = $itemValue[1].Text
			[string]$Description = $itemValue[1].Text
			[string]$Pattern = $itemValue[2].Text
			[string]$Translation = $itemValue[3].Text
			[bool]$ExtensionValue = $itemValue[4].Text
			#$ExtensionValue = $ExtensionValue.ToLower()
												
			
			$orgIndex = $lvi.Index
			if($orgIndex -lt ($lv.Items.Count - 1))
			{
				$index = $orgIndex + 1
				
				Write-Host "RUNNING: `$nr = New-CsVoiceNormalizationRule -Identity "${Scope}/${Name}" -Description $Description -Pattern $Pattern -Translation $Translation -Priority $index -IsInternalExtension $ExtensionValue -InMemory" -foreground "green"
				$nr = New-CsVoiceNormalizationRule -Identity "${Scope}/${Name}" -Description $Description -Pattern $Pattern -Translation $Translation -Priority $index -IsInternalExtension $ExtensionValue  -InMemory			
				
				$NormArray = (Get-CsTenantDialPlan -identity $Scope | select NormalizationRules).NormalizationRules
				Write-Verbose "INITIAL ARRAY"
				foreach($item in $NormArray){Write-Verbose $item}
				Write-Host
				#Remove Item
				$NormArray.RemoveAt($orgIndex)
				Write-Verbose "AFTER DELETE"
				foreach($item in $NormArray){Write-Verbose $item}
				Write-Host 
				#Insert Item
				$NormArray.Insert($index,$nr)
				Write-Verbose "AFTER INSERT"
				foreach($item in $NormArray){Write-Verbose $item}
				
				Set-CsTenantDialPlan -Identity $Scope -NormalizationRules @{Replace=$NormArray}
				
				GetNormalisationPolicy
				
				$lv.Items[$index].Selected = $true
				$lv.Items[$index].EnsureVisible()
			}
			else
			{
				Write-Host "INFO: Cannot move item any lower..." -foreground "Yellow"
			}
		}
	}
}
#>

function GetNormalisationPolicy
{
		
	$theIdentity = $UserDropDownBox.SelectedItem.ToString()
	Write-Host "INFO: Getting rules for $theIdentity" -foreground "yellow"
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		ResetFields
        $UserCallingSetting = Get-CsUserCallingSettings -identity $theIdentity
        if($UserCallingSetting){
		    $SipUriLabelField.Text=$UserCallingSetting.sipuri
            $IsForwardingEnabledDropDownBox.text = $UserCallingSetting.IsForwardingEnabled
            $ForwardingTypeDropDownBox.text = $UserCallingSetting.ForwardingType
            $ForwardingTargetTextBox.Text=$UserCallingSetting.ForwardingTarget
            $ForwardingTargetTypeDropDownBox.text = $UserCallingSetting.ForwardingTargetType
            $IsUnansweredEnabledDropDownBox.text = $UserCallingSetting.IsUnansweredEnabled
            $UnansweredTargetTextBox.Text=$UserCallingSetting.UnansweredTarget
            $UnansweredTargetTypeDropDownBox.text = $UserCallingSetting.UnansweredTargetType
            $UnansweredDelayTextBox.Text = $UserCallingSetting.UnansweredDelay
        }
	}
	
}

function ApplyUserChanges
{
    $user=$UserDropDownBox.Text

    

    $IsForwardingEnabled=$false
    if($IsForwardingEnabledDropDownBox.text -eq "true"){
        $IsForwardingEnabled=$true
        
    }
    $IsUnansweredEnabled=$false
    if($IsUnansweredEnabledDropDownBox.text -eq "true"){
        $IsUnansweredEnabled=$true
    }
    
    $ForwardingType = $ForwardingTypeDropDownBox.text 
    $ForwardingTargetType= $ForwardingTargetTypeDropDownBox.text
    $ForwardingTarget= $ForwardingTargetTextBox.text  
    $UnansweredTargetType= $UnansweredTargetTypeDropDownBox.Text 
    $UnansweredDelay= $UnansweredDelayTextBox.Text 
    $UnansweredTarget= $UnansweredTargetTextBox.Text
    
    Set-CsUserCallingSettings -Identity $User -IsForwardingEnabled $IsForwardingEnabled -ForwardingType $ForwardingType -ForwardingTargetType $ForwardingTargetType -ForwardingTarget $ForwardingTarget
    Set-CsUserCallingSettings -Identity $User -IsUnansweredEnabled $IsUnansweredEnabled -UnansweredTargetType $UnansweredTargetType -UnansweredDelay $UnansweredDelay -UnansweredTarget $UnansweredTarget
    #Set-CsUserCallingSettings -Identity $User -IsForwardingEnabled $IsForwardingEnabled -ForwardingType $ForwardingType -ForwardingTargetType $ForwardingTargetType -ForwardingTarget $ForwardingTarget -IsUnansweredEnabled $IsUnansweredEnabled -UnansweredTargetType $UnansweredTargetType -UnansweredDelay $UnansweredDelay -UnansweredTarget $UnansweredTarget
}

function ResetUserParameter
{
    $user=$UserDropDownBox.Text
    Set-CsUserCallingSettings -Identity $User -IsForwardingEnabled $false 
    Set-CsUserCallingSettings -Identity $User -IsUnansweredEnabled $True -UnansweredTargetType VoiceMail -UnansweredDelay "00:00:10"
}
function ResetFields 
{
    $SipUriLabelField.Text=""
    $IsForwardingEnabledDropDownBox.SelectedIndex = -1
    $ForwardingTypeDropDownBox.SelectedIndex = -1
    $ForwardingTargetTextBox.Text=""
    $ForwardingTargetTypeDropDownBox.SelectedIndex = -1
    $IsUnansweredEnabledDropDownBox.SelectedIndex = -1
    $UnansweredTargetTextBox.Text=""
    $UnansweredTargetTypeDropDownBox.SelectedIndex = -1
    $UnansweredDelayTextBox.Text=""
}

function CompileUserList($search)
{
    Get-CsOnlineUser | where TeamsUpgradeEffectiveMode -eq "TeamsOnly" | select-object userprincipalname | ForEach-Object {[void] $UserDropDownBox.Items.Add($_.userprincipalname)}
}

function CompileUserList
{
    Get-CsOnlineUser | where TeamsUpgradeEffectiveMode -eq "TeamsOnly" | select-object userprincipalname | Sort-Object -Property userprincipalname | ForEach-Object {[void] $UserDropDownBox.Items.Add($_.userprincipalname)}
}

<#
function UpdateListViewSettings
{
	if($lv.SelectedItems.count -eq 0)
	{
		$NameTextBox.Text = ""
		$DescriptionTextBox.Text = ""
		$PatternTextBox.Text = ""
		$TranslationTextBox.Text = ""
		$ExtensionCheckBox.Checked = $false
	}
	else
	{
		foreach ($item in $lv.SelectedItems)
		{
			[string]$itemName = $item.Text
			$itemValue = $item.SubItems
			
			$NameTextBox.Text = $itemName
			
			[string]$settingValue1 = $itemValue[1].Text
			$DescriptionTextBox.Text = $settingValue1
			[string]$settingValue3 = $itemValue[2].Text
			$PatternTextBox.Text = $settingValue3
			[string]$settingValue4 = $itemValue[3].Text
			$TranslationTextBox.Text = $settingValue4
			[string]$settingValue4 = $itemValue[4].Text
			$settingValue4 = $settingValue4.ToLower()
			if($settingValue4 -eq "true")
			{$ExtensionCheckBox.Checked = $true}
			else
			{$ExtensionCheckBox.Checked = $false}
			
		}
	}
}
#>

#Add / Edit an item
<#
function AddSetting
{
	[string]$Scope = $policyDropDownBox.SelectedItem.ToString()
	[string]$Name = $NameTextBox.Text
	[string]$Description = $DescriptionTextBox.Text
	[string]$Pattern = $PatternTextBox.Text
	[string]$Translation = $TranslationTextBox.Text
	$ExtensionBool = $ExtensionCheckBox.Checked
	
	if($Scope -ne "" -and $Scope -ne $null -and $Name -ne "" -and $Name -ne $null -and $Pattern -ne "" -and $Pattern -ne $null -and $Translation -ne "" -and $Translation -ne $null)
	{
		$checkResult = CheckTeamsOnline
		if($checkResult)
		{
			[string]$Name = $NameTextBox.Text
			$EditSetting = $false
			$LoopNo = 0
			foreach($item in $lv.Items)
			{
				[string]$listName = $item.Text
				if($listName.ToLower() -eq $Name.ToLower())
				{
					$EditSetting = $true
					$Priority = $LoopNo
					break
				}
				$LoopNo++
			}
			if($EditSetting)
			{
				Write-Host "INFO: Name is already in the list. Editing setting" -foreground "yellow"
				
				Write-Verbose "SELECTED INDEX: $($lv.SelectedIndices[0])" 
				$orgIndex = $lv.SelectedIndices[0] #$lvi.Index
				$NewIndex = $orgIndex - 1
							
				$NormArray = (Get-CsTenantDialPlan -identity $Scope | select NormalizationRules).NormalizationRules
				foreach($item in $NormArray){Write-Verbose $item}
				#Remove Item
				$NormArray.RemoveAt($orgIndex)
				foreach($item in $NormArray){Write-Verbose $item}
				Write-Host 
				#Insert Item
				Write-Host "RUNNING: `$nr = New-CsVoiceNormalizationRule -Identity "${Scope}/${Name}" -Description $Description -Pattern $Pattern -Translation $Translation -IsInternalExtension $ExtensionBool -InMemory" -foreground "green"
				$nr = New-CsVoiceNormalizationRule -Identity "${Scope}/${Name}" -Description $Description -Pattern $Pattern -Translation $Translation -IsInternalExtension $ExtensionBool -InMemory
				$NormArray.Insert($orgIndex,$nr)
				foreach($item in $NormArray){Write-Verbose $item}
				
				Set-CsTenantDialPlan -Identity $Scope -NormalizationRules @{Replace=$NormArray}
							
				GetNormalisationPolicy
				
				$lv.Items[$Priority].Selected = $true
				$lv.Items[$Priority].EnsureVisible()
			}
			else   # ADD
			{
				
				Write-Host "RUNNING: `$nr = New-CsVoiceNormalizationRule -Identity "${Scope}/${Name}" -Description $Description -Pattern $Pattern -Translation $Translation -IsInternalExtension $ExtensionBool -InMemory" -foreground "green"
				$nr = New-CsVoiceNormalizationRule -Identity "${Scope}/${Name}" -Description $Description -Pattern $Pattern -Translation $Translation -IsInternalExtension $ExtensionBool -InMemory
				Write-Host "RUNNING: New-CsTenantDialPlan -Identity ${Scope} -NormalizationRules @{Add=`$nr}" -foreground "green"
				Set-CsTenantDialPlan -Identity ${Scope} -NormalizationRules @{Add=$nr}
				
				GetNormalisationPolicy
				
				$count = $lv.Items.Count - 1
				$lv.Items[$count].Selected = $true
				$lv.Items[$count].EnsureVisible()
				
			}
		}
	}
	else
	{
		Write-Host "ERROR: Please enter values for Name, Pattern and Translation." -foreground "red"
	}
}
#>
<#
function DeleteSetting
{
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		if($lv.SelectedItems.Count -le 0)
		{Write-Host "ERROR: No items selected to delete. Please select items before selecting delete." -foreground "red"}
		foreach ($lvi in $lv.SelectedItems)
		{
			#GET SETTINGS OF SELECTED ITEM
			$item = $lv.Items[$lvi.Index]
			$itemValue = $item.SubItems

			[string]$Scope = $policyDropDownBox.SelectedItem.ToString()
			[string]$Scope = $Scope.Replace("site:","")
			[string]$Name = $item.Text
			#[string]$Priority = $itemValue[1].Text
			[string]$Description = $itemValue[1].Text
			[string]$Pattern = $itemValue[2].Text
			[string]$Translation = $itemValue[3].Text
			
			$orgIndex = $lvi.Index
			
			Write-Host "INFO: Removing - ${Scope}/${Name}" -foreground "Yellow"
			
			Write-Host "RUNNING: (Get-CsTenantDialPlan -identity ${Scope}).NormalizationRules | Where-Object {$_.Name -eq ${Name}}" -foreground "green"
			$policyItem = (Get-CsTenantDialPlan -identity ${Scope}).NormalizationRules | Where-Object {$_.Name -eq "${Name}"}
			Write-Host "RUNNING: Set-CsTenantDialPlan -Identity $Scope -NormalizationRules @{Remove=$policyItem}" -foreground "green"
			Set-CsTenantDialPlan -Identity $Scope -NormalizationRules @{Remove=$policyItem}
			
			GetNormalisationPolicy
			
			if($orgIndex -ge 1)
			{
				$index = $orgIndex - 1
				
				$lv.Items[$index].Selected = $true
				$lv.Items[$index].EnsureVisible()
			}
			else
			{
				$lv.Items[0].Selected = $true
				$lv.Items[0].EnsureVisible()
			}
			
			UpdateListViewSettings
		}
	}
}
#>
<#
function DeleteAllSettings
{
	$checkResult = CheckTeamsOnline
	if($checkResult)
	{
		[string]$Scope = $policyDropDownBox.SelectedItem.ToString()

		Write-Host "RUNNING: (Get-CsTenantDialPlan -identity ${Scope}).NormalizationRules" -foreground "green"
		$policyItem = (Get-CsTenantDialPlan -identity ${Scope}).NormalizationRules #| Where-Object {$_.Name -eq "${Name}"}
		
		Write-Host "RUNNING: Set-CsTenantDialPlan -Identity $Scope -NormalizationRules @{Remove=$policyItem}" -foreground "green"
		Set-CsTenantDialPlan -Identity $Scope -NormalizationRules @{Remove=$policyItem}
		
		GetNormalisationPolicy
	}
}
#>
<#
function TestPhoneNumberNew()
{

	$TestPhoneResultTextLabel.Text = "Test Result: No Match"
	$TestPhonePatternTextLabel.Text = "Matched Pattern: No Match"
	$TestPhoneTranslationTextLabel.Text = "Matched Translation: No Match"
	
	foreach($tempitem in $lv.Items)
	{
		$tempitem.ForeColor = "Black"
	}
	$PhoneNumber = $TestPhoneTextBox.Text
	#$Rules = Get-CsAddressBookNormalizationRule
	
	Write-Host ""
	Write-Host "-------------------------------------------------------------" -foreground "Green"
	Write-Host "TESTING: $PhoneNumber" -foreground "Green"
	Write-Host ""

	$TopLoopNo = 0
	$firstFound = $true
	foreach($item in $lv.Items)
	{
		$itemValue = $item.SubItems

		[string]$Pattern = $itemValue[2].Text
		[string]$Translation = $itemValue[3].Text
		
		#Clean up the Phone Number
		$PhoneNumberStripped = $PhoneNumber.Replace(" ","").Replace("(","").Replace(")","").Replace("[","").Replace("]","").Replace("{","").Replace("}","").Replace(".","").Replace("-","").Replace(":","")
		
		Write-Verbose "TESTING PATTERN: $Pattern" #DEBUG
		
		$PatternStartEnd = "^$Pattern$"
		Try
		{
			$StartPatternResult = $PhoneNumberStripped -cmatch $PatternStartEnd
		}
		Catch
		{
			#This error was already reported. So don't bother reporting it again.
		}
		
		if($StartPatternResult)
		{
			if($firstFound)
			{
				Write-Host "First Matched Pattern: $Pattern" -foreground "Green"
				Write-Host "First Matched Translation: $Translation" -foreground "Green"
				$TestPhonePatternTextLabel.Text = "Matched Pattern: $Pattern"
				$TestPhoneTranslationTextLabel.Text = "Matched Translation: $Translation"
				
				$Group1 = [regex]::match($PhoneNumberStripped,$Pattern).Groups[1].Value
				$Group2 = [regex]::match($PhoneNumberStripped,$Pattern).Groups[2].Value
				$Group3 = [regex]::match($PhoneNumberStripped,$Pattern).Groups[3].Value
				$Group4 = [regex]::match($PhoneNumberStripped,$Pattern).Groups[4].Value
				$Group5 = [regex]::match($PhoneNumberStripped,$Pattern).Groups[5].Value
				$Group6 = [regex]::match($PhoneNumberStripped,$Pattern).Groups[6].Value
				$Group7 = [regex]::match($PhoneNumberStripped,$Pattern).Groups[7].Value
				$Group8 = [regex]::match($PhoneNumberStripped,$Pattern).Groups[8].Value
				$Group9 = [regex]::match($PhoneNumberStripped,$Pattern).Groups[9].Value
				
				Write-Host
				if($Group1 -ne ""){Write-Host "Group 1: " $Group1 -foreground "Yellow"}
				if($Group2 -ne ""){Write-Host "Group 2: " $Group2 -foreground "Yellow"}
				if($Group3 -ne ""){Write-Host "Group 3: " $Group3 -foreground "Yellow"}
				if($Group4 -ne ""){Write-Host "Group 4: " $Group4 -foreground "Yellow"}
				if($Group5 -ne ""){Write-Host "Group 5: " $Group5 -foreground "Yellow"}
				if($Group6 -ne ""){Write-Host "Group 6: " $Group6 -foreground "Yellow"}
				if($Group7 -ne ""){Write-Host "Group 7: " $Group7 -foreground "Yellow"}
				if($Group8 -ne ""){Write-Host "Group 8: " $Group8 -foreground "Yellow"}
				if($Group9 -ne ""){Write-Host "Group 9: " $Group9 -foreground "Yellow"}
				
				Write-Host				
				$Result = $Translation.Replace('$1',"$Group1")
				$Result = $Result.Replace('$2',"$Group2")
				$Result = $Result.Replace('$3',"$Group3")
				$Result = $Result.Replace('$4',"$Group4")
				$Result = $Result.Replace('$5',"$Group5")
				$Result = $Result.Replace('$6',"$Group6")
				$Result = $Result.Replace('$7',"$Group7")
				$Result = $Result.Replace('$8',"$Group8")
				$Result = $Result.Replace('$9',"$Group9")
				Write-Host "Result: " $Result -foreground "Green"
				$TestPhoneResultTextLabel.Text = "Test Result: ${Result}"
				
				$firstFound = $false
				$item.ForeColor = "Green"
			}
			else
			{
				$item.ForeColor = "Blue"
			}
		}
	}
	$lv.SelectedItems.Clear()
	$lv.Focus()	
	Write-Host "-------------------------------------------------------------" -foreground "Green"
}
#>
<#
function Import-Config
{
	$Filter = "All Files (*.*)|*.*"
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$fileForm = New-Object System.Windows.Forms.OpenFileDialog
	$fileForm.InitialDirectory = $pathbox.text
	$fileForm.Filter = $Filter
	$fileForm.Title = "Open File"
	$Show = $fileForm.ShowDialog()
	if ($Show -eq "OK")
	{
		#IMPORT CODE
	}
	else
	{
		Write-Host "INFO: Operation cancelled by user." -foreground "yellow"
	}
	
}
#>
<#
function Export-Config
{
	#File Dialog
	[string] $pathVar = "C:\"
	$Filter="All Files (*.*)|*.*"
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$objDialog = New-Object System.Windows.Forms.SaveFileDialog
	#$objDialog.InitialDirectory = 
	$objDialog.FileName = "Company_Phone_Number_Normalization_Rules.txt"
	$objDialog.Filter = $Filter
	$objDialog.Title = "Export File Name"
	$objDialog.CheckFileExists = $false
	$Show = $objDialog.ShowDialog()
	if ($Show -eq "OK")
	{
		[string]$outputFile = $objDialog.FileName
		$outputFile = "${outputFile}"
		$output = ""
		foreach($item in $lv.Items)
		{
			$itemValue = $item.SubItems

			[string]$Pattern = $itemValue[2].Text
			[string]$Translation = $itemValue[3].Text
			[string]$Name = $item.Text
			[string]$Description = $itemValue[1].Text
			
			$output += "# $Name $Description`r`n$Pattern`r`n$Translation`r`n`r`n"
		}
		
		$output | out-file -Encoding UTF8 -FilePath $outputFile -Force					
		Write-Host "Written File to $outputFile...." -foreground "yellow"
	}
	else
	{
		return
	}
	
}
#>

$result = CheckTeamsOnlineInitial


# Activate the form ============================================================
$mainForm.Add_Shown({$mainForm.Activate()})
[void] $mainForm.ShowDialog()	

