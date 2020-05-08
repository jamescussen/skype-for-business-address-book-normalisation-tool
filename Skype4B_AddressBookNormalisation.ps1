########################################################################
# Name: Skype for Business Address Book Normalisation Tool 
# Version: v1.03 (28/11/2019)
# Date: 1/7/2015
# Created By: James Cussen
# Web Site: http://www.myteamslab.com (formally http://www.myteamslab.com)
# 
# Notes: This is a Powershell tool. To run the tool, open it from the Powershell command line on a Lync server.
#		 For more information on the requirements for setting up and using this tool please visit http://www.myteamslab.com.
#
# Copyright: Copyright (c) 2019, James Cussen (www.myteamslab.com) All rights reserved.
# Licence: 	Redistribution and use of script, source and binary forms, with or without modification, are permitted provided that the following conditions are met:
#				1) Redistributions of script code must retain the above copyright notice, this list of conditions and the following disclaimer.
#				2) Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
#				3) Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
#				4) This license does not include any resale or commercial use of this software.
#				5) Any portion of this software may not be reproduced, duplicated, copied, sold, resold, or otherwise exploited for any commercial purpose without express written consent of James Cussen.
#			THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; LOSS OF GOODWILL OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#
# Release Notes:
# 1.00 Initial Release.
#		
# 1.01 Update:
#
#	- Added warning message on the Remove policy button to save you from yourself :)
#	- Removed second .txt from the export name.
#
# 1.02 Update:
#
#	- Script now doesn't strip ";" char before applying regex. (Thanks Daniel Appleby for reporting)
#   - Updated Signature (25-5-2016)
#
# 1.03 Test Number Update
#
#	- Updated the Test Number capability to be much more accurate.
#	- Moved around some UI elements to make things a bit clearer.
#	- Added some more error checking.
#	- Updated Icon.
########################################################################

$theVersion = $PSVersionTable.PSVersion
$MajorVersion = $theVersion.Major

Write-Host ""
Write-Host "--------------------------------------------------------------"
Write-Host "Powershell Version Check..." -foreground "yellow"
if($MajorVersion -eq  "1")
{
	Write-Host "This machine only has Version 1 Powershell installed.  This version of Powershell is not supported." -foreground "red"
}
elseif($MajorVersion -eq  "2")
{
	Write-Host "This machine has Version 2 Powershell installed. This version of Powershell is not supported." -foreground "red"
}
elseif($MajorVersion -eq  "3")
{
	Write-Host "This machine has version 3 Powershell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "4")
{
	Write-Host "This machine has version 4 Powershell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "5")
{
	Write-Host "This machine has version 4 Powershell installed. CHECK PASSED!" -foreground "green"
}
else
{
	Write-Host "This machine has version $MajorVersion Powershell installed. Unknown level of support for this version." -foreground "yellow"
}
Write-Host "--------------------------------------------------------------"
Write-Host ""

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


$Script:LyncModuleAvailable = $false
$Script:SkypeModuleAvailable = $false

$command = "Get-CsAddressBookNormalizationRule"
if(Get-Command $command -errorAction SilentlyContinue)
{
	Write-Host "Get-CsAddressBookNormalizationRule is available." -foreground "green"
}
else
{
	Write-Host "ERROR: Get-CsAddressBookNormalizationRule is not available. Make sure that you are running this tool on a Skype for Business system." -foreground "red"
	Write-Host "Exiting." -foreground "red"
	Write-Host ""
	Exit
}

Write-Host "--------------------------------------------------------------"
#Import Lync Module
if(Get-MyModule "Lync")
{
	Invoke-Expression "Import-Module Lync"
	Write-Host "Imported Lync Module..." -foreground "green"
	$Script:LyncModuleAvailable = $true
}
else
{
	Write-Host "Unable to import Lync Module... The Lync module is required to run this tool." -foreground "yellow"
}
#Import SkypeforBusiness Module
if(Get-MyModule "SkypeforBusiness")
{
	Invoke-Expression "Import-Module SkypeforBusiness"
	Write-Host "Imported SkypeforBusiness Module..." -foreground "green"
	$Script:SkypeModuleAvailable = $true
}
else
{
	Write-Host "Unable to import SkypeforBusiness Module... (Expected on a Lync 2013 system)" -foreground "yellow"
}


# Set up the form  ============================================================

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "SfB Address Book Normalisation Tool 1.03"
$objForm.Size = New-Object System.Drawing.Size(525,645) 
$objForm.MinimumSize = New-Object System.Drawing.Size(520,410) 
$objForm.StartPosition = "CenterScreen"
[byte[]]$WindowIcon = @(0, 0, 1, 0, 1, 0, 16, 16, 0, 0, 1, 0, 32, 0, 104, 4, 0, 0, 22, 0, 0, 0, 40, 0, 0, 0, 16, 0, 0, 0, 32, 0, 0, 0, 1, 0, 32, 0, 0, 0, 0, 0, 64, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 87, 40, 0, 194, 111, 82, 63, 255, 101, 68, 40, 255, 93, 51, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 93, 51, 0, 255, 97, 60, 24, 255, 99, 64, 33, 255, 92, 50, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 93, 52, 0, 255, 98, 62, 28, 255, 113, 86, 68, 255, 86, 37, 0, 190, 72, 0, 0, 255, 241, 240, 239, 255, 184, 177, 173, 255, 83, 24, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 67, 0, 0, 255, 231, 228, 227, 255, 243, 242, 241, 255, 61, 0, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 87, 36, 0, 255, 158, 145, 139, 255, 255, 255, 255, 255, 69, 0, 0, 255, 71, 0, 0, 255, 236, 234, 233, 255, 180, 172, 168, 255, 83, 24, 0, 255, 94, 54, 0, 255, 90, 45, 0, 255, 118, 94, 79, 255, 255, 254, 254, 255, 245, 244, 244, 255, 140, 124, 115, 255, 88, 40, 0, 255, 94, 54, 0, 255, 87, 36, 0, 255, 154, 141, 135, 255, 249, 249, 248, 255, 67, 0, 0, 255, 71, 0, 0, 255, 236, 234, 233, 255, 180, 172, 168, 255, 83, 24, 0, 255, 94, 54, 0, 255, 74, 0, 0, 255, 214, 210, 208, 255, 194, 188, 184, 255, 173, 163, 158, 255, 228, 225, 224, 255, 69, 0, 0, 255, 94, 54, 0, 255, 87, 36, 0, 255, 154, 141, 135, 255, 249, 249, 248, 255, 67, 0, 0, 255, 71, 0, 0, 255, 236, 234, 233, 255, 180, 172, 168, 255, 83, 24, 0, 255, 92, 50, 0, 255, 87, 38, 0, 255, 255, 255, 255, 255, 92, 50, 0,255, 67, 0, 0, 255, 255, 255, 255, 255, 110, 81, 62, 255, 91, 47, 0, 255, 87, 36, 0, 255, 154, 141, 135, 255, 249, 249, 248, 255, 67, 0, 0, 255, 71, 0, 0, 255,236, 234, 233, 255, 180, 172, 168, 255, 83, 24, 0, 255, 80, 7, 0, 255, 194, 188, 185, 255, 218, 214, 213, 255, 73, 0, 0, 255, 78, 0, 0, 255, 200, 195, 192, 255, 210, 205, 203, 255, 76, 0, 0, 255, 87, 36, 0, 255, 154, 141, 135, 255, 249, 249, 248, 255, 67, 0, 0, 255, 71, 0, 0, 255, 236, 234, 233, 255, 180, 172, 168, 255, 82, 21, 0, 255, 65, 0, 0, 255, 255, 255, 255, 255, 123, 101, 88, 255, 90, 44, 0, 255, 92, 49, 0, 255, 95, 55, 5, 255, 255, 255, 255, 255, 82, 21, 0, 255, 85,31, 0, 255, 154, 141, 135, 255, 249, 249, 248, 255, 67, 0, 0, 255, 71, 0, 0, 255, 236, 234, 233, 255, 180, 172, 168, 255, 71, 0, 0, 255, 171, 162, 156, 255, 236, 234, 233, 255, 66, 0, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 73, 0, 0, 255, 219, 216, 214, 255, 190, 183, 179, 255, 72, 0, 0, 255, 154, 141, 135, 255, 249, 249, 248, 255, 67, 0, 0, 255, 71, 0, 0, 255, 236, 234, 233, 255, 180, 172, 168, 255, 37, 0, 0, 255, 246, 246, 245, 255, 154, 141, 134, 255, 86, 36, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 89, 44, 0, 255, 125, 104, 91, 255, 255, 255, 255, 255, 48, 0, 0, 255, 154, 141, 134, 255, 249, 249, 248, 255, 67, 0, 0, 255, 71, 0, 0, 255, 236, 234, 233, 255, 178, 169, 164, 255, 137, 120, 110, 255, 252, 251, 251, 255, 60, 0, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 65, 0, 0, 255, 236, 235, 234, 255, 164, 153, 147, 255, 150, 136, 128, 255, 249, 249, 248, 255, 67, 0, 0, 255, 71, 0, 0, 255, 236, 234, 233, 255, 171, 161, 156, 255, 226, 223, 222, 255, 180, 172, 167, 255, 83, 23, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 86, 36, 0, 255, 155, 142, 135, 255, 242, 241, 240, 255, 140, 123, 114, 255, 249, 249, 248, 255, 67, 0, 0, 255, 71, 0, 0, 255, 235, 233, 232, 255, 190, 183, 179, 255, 255, 255, 255, 255, 71, 0, 0, 255, 93, 52, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 58, 0, 0, 255, 250, 249, 249, 255, 184, 176, 172, 255, 247, 247, 246, 255, 67, 0, 0, 255, 71, 0, 0, 255, 231, 229, 228, 255, 253, 253, 253, 255, 198, 192, 189, 255, 78, 0, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255,94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 83, 24, 0, 255, 177, 168, 163, 255, 252, 252, 252, 255, 244, 243, 243, 255, 67, 0, 0, 255, 71, 0, 0, 255, 231, 229, 228, 255, 255, 255, 255, 255, 97, 61, 25, 255, 91, 49, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 93, 52, 0, 255, 71, 0, 0, 255, 255, 255, 255, 255, 245, 244, 243, 255, 67, 0, 0, 255, 74, 0, 0, 255, 232, 229, 228, 255, 215, 211, 209, 255,73, 0, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 79, 2, 0, 255, 195, 189, 186, 255, 245, 244, 244, 255, 72, 0, 0, 255, 92, 51, 0, 194, 74, 0, 0, 255, 76, 0, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 94, 54, 0, 255, 80, 9, 0, 255, 72, 0, 0, 255, 91, 51, 0, 195, 0, 0, 0, 0, 0,0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
$objForm.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
$objForm.KeyPreview = $True
$objForm.TabStop = $false


$MyLinkLabel = New-Object System.Windows.Forms.LinkLabel
$MyLinkLabel.Location = New-Object System.Drawing.Size(330,7)
$MyLinkLabel.Size = New-Object System.Drawing.Size(120,15)
$MyLinkLabel.DisabledLinkColor = [System.Drawing.Color]::Red
$MyLinkLabel.VisitedLinkColor = [System.Drawing.Color]::Blue
$MyLinkLabel.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline
$MyLinkLabel.LinkColor = [System.Drawing.Color]::Navy
$MyLinkLabel.TabStop = $False
$MyLinkLabel.Text = "www.myteamslab.com"
$MyLinkLabel.add_click(
{
	 [system.Diagnostics.Process]::start("http://www.myteamslab.com")
})
$objForm.Controls.Add($MyLinkLabel)

<#
#ClientPolicyLabel Label ============================================================
$ClientPolicyLabel = New-Object System.Windows.Forms.Label
$ClientPolicyLabel.Location = New-Object System.Drawing.Size(25,10) 
$ClientPolicyLabel.Size = New-Object System.Drawing.Size(150,15) 
$ClientPolicyLabel.Text = "Normalisation Policies:"
$ClientPolicyLabel.TabStop = $false
$objForm.Controls.Add($ClientPolicyLabel)
#>

#Policy Label ============================================================
$policyLabel = New-Object System.Windows.Forms.Label
$policyLabel.Location = New-Object System.Drawing.Size(20,33) 
$policyLabel.Size = New-Object System.Drawing.Size(50,15) 
$policyLabel.Text = "Policies: "
$policyLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$policyLabel.TabStop = $false
$objForm.Controls.Add($policyLabel)


# Add Client Policy Dropdown box ============================================================
$policyDropDownBox = New-Object System.Windows.Forms.ComboBox 
$policyDropDownBox.Location = New-Object System.Drawing.Size(75,30) 
$policyDropDownBox.Size = New-Object System.Drawing.Size(250,20) 
$policyDropDownBox.DropDownHeight = 200 
$policyDropDownBox.tabIndex = 1
$policyDropDownBox.DropDownStyle = "DropDownList"
$policyDropDownBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$objForm.Controls.Add($policyDropDownBox) 

Get-CsAddressBookNormalizationConfiguration | select-object identity | ForEach-Object {[void] $policyDropDownBox.Items.Add($_.identity)}

$numberOfItems = $policyDropDownBox.count
if($numberOfItems -gt 0)
{
	$policyDropDownBox.SelectedIndex = 0
}

$policyDropDownBox.add_SelectedValueChanged(
{
	GetNormalisationPolicy
})


#NewPolicy button
$NewPolicyButton = New-Object System.Windows.Forms.Button
$NewPolicyButton.Location = New-Object System.Drawing.Size(330,30)
$NewPolicyButton.Size = New-Object System.Drawing.Size(50,20)
$NewPolicyButton.Text = "New.."
$NewPolicyButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$NewPolicyButton.Add_Click(
{
	
	$NewPolicyName = New-Policy -Message "Enter the site name of new policy:" -WindowTitle "New Address Book Policy" -DefaultText "Policy Name"
	
	if($NewPolicyName -ne "" -and $NewPolicyName -ne $null)
	{
		Write-host "New Policy: $NewPolicyName"
		Invoke-Expression "New-CsAddressBookNormalizationConfiguration -Identity `"$NewPolicyName`""

		#site:Redmond
		
		$policyDropDownBox.Items.Clear()
		Get-CsAddressBookNormalizationConfiguration | select-object identity | ForEach-Object {[void] $policyDropDownBox.Items.Add($_.identity)}
		$policyDropDownBox.SelectedIndex = $policyDropDownBox.Items.IndexOf("$NewPolicyName")

	}
	
})
$objForm.Controls.Add($NewPolicyButton)


#RemovePolicy button
$RemovePolicyButton = New-Object System.Windows.Forms.Button
$RemovePolicyButton.Location = New-Object System.Drawing.Size(385,30)
$RemovePolicyButton.Size = New-Object System.Drawing.Size(55,20)
$RemovePolicyButton.Text = "Remove"
$RemovePolicyButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$RemovePolicyButton.Add_Click(
{
	$thePolicyName = $policyDropDownBox.SelectedItem.ToString()
	$a = new-object -comobject wscript.shell 
	$intAnswer = $a.popup("Are you sure you want to remove the entire $thePolicyName policy?",0,"Remove Policy",4) 
	if ($intAnswer -eq 6) { 
					
		Write-Host "Removing Policy: $thePolicyName" -foreground "green"
		Invoke-Expression "Remove-CsAddressBookNormalizationConfiguration -Identity `"$thePolicyName`""
		
		$policyDropDownBox.Items.Clear()
		Get-CsAddressBookNormalizationConfiguration | select-object identity | ForEach-Object {[void] $policyDropDownBox.Items.Add($_.identity)}
		
		$numberOfItems = $policyDropDownBox.count
		if($numberOfItems -gt 0)
		{
			$policyDropDownBox.SelectedIndex = 0
		}
	}
})
$objForm.Controls.Add($RemovePolicyButton)


$lv = New-Object windows.forms.ListView
$lv.View = [System.Windows.Forms.View]"Details"
$lv.Size = New-Object System.Drawing.Size(422,290)
$lv.Location = New-Object System.Drawing.Size(20,57)
$lv.FullRowSelect = $true
$lv.GridLines = $true
$lv.HideSelection = $false
#$lv.MultiSelect = $false
#$lv.Sorting = [System.Windows.Forms.SortOrder]"Ascending"
$lv.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top
[void]$lv.Columns.Add("Name", 100)
[void]$lv.Columns.Add("Priority", 50)
[void]$lv.Columns.Add("Description", 100)
[void]$lv.Columns.Add("Pattern", 75)
[void]$lv.Columns.Add("Translation", 75)
$objForm.Controls.Add($lv)

$lv.add_MouseUp(
{
	UpdateListViewSettings
})

# Groups Key Event ============================================================
$lv.add_KeyUp(
{
	if ($_.KeyCode -eq "Up" -or $_.KeyCode -eq "Down") 
	{	
		UpdateListViewSettings
	}
})


#Up button
$UpButton = New-Object System.Windows.Forms.Button
$UpButton.Location = New-Object System.Drawing.Size(450,160)
$UpButton.Size = New-Object System.Drawing.Size(50,20)
$UpButton.Text = "UP"
$UpButton.TabStop = $false
$UpButton.Anchor = [System.Windows.Forms.AnchorStyles]::Right
$UpButton.Add_Click(
{
	DisableAllButtons
	Move-Up
	EnableAllButtons
})
$objForm.Controls.Add($UpButton)

# Priority Label ============================================================
$PriorityLabel = New-Object System.Windows.Forms.Label
$PriorityLabel.Location = New-Object System.Drawing.Size(455,140) 
$PriorityLabel.Size = New-Object System.Drawing.Size(60,15) 
$PriorityLabel.Text = "Priority"
$PriorityLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
$PriorityLabel.TabStop = $false
$objForm.Controls.Add($PriorityLabel)

#Down button
$DownButton = New-Object System.Windows.Forms.Button
$DownButton.Location = New-Object System.Drawing.Size(450,190)
$DownButton.Size = New-Object System.Drawing.Size(50,20)
$DownButton.Text = "DOWN"
$DownButton.TabStop = $false
$DownButton.Anchor = [System.Windows.Forms.AnchorStyles]::Right
$DownButton.Add_Click(
{
	DisableAllButtons
	Move-Down
	EnableAllButtons
})
$objForm.Controls.Add($DownButton)




#NameTextLabel Label ============================================================
$NameTextLabel = New-Object System.Windows.Forms.Label
$NameTextLabel.Location = New-Object System.Drawing.Size(40,367) 
$NameTextLabel.Size = New-Object System.Drawing.Size(45,15) 
$NameTextLabel.Text = "Name:"
$NameTextLabel.TabStop = $false
$NameTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$objForm.Controls.Add($NameTextLabel)


#Name Text box ============================================================
$NameTextBox = New-Object System.Windows.Forms.TextBox
$NameTextBox.location = new-object system.drawing.size(85,365)
$NameTextBox.size = new-object system.drawing.size(250,23)
$NameTextBox.tabIndex = 1
$NameTextBox.text = "Name"   
$NameTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$objform.controls.add($NameTextBox)
$NameTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{	
		#Do Nothing
	}
})


#Description Label ============================================================
$DescriptionTextLabel = New-Object System.Windows.Forms.Label
$DescriptionTextLabel.Location = New-Object System.Drawing.Size(13,393) 
$DescriptionTextLabel.Size = New-Object System.Drawing.Size(70,15) 
$DescriptionTextLabel.Text = "Description:"
$DescriptionTextLabel.TabStop = $false
$DescriptionTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$objForm.Controls.Add($DescriptionTextLabel)

#$DescriptionTextBox Text box ============================================================
$DescriptionTextBox = New-Object System.Windows.Forms.TextBox
$DescriptionTextBox.location = new-object system.drawing.size(85,390)
$DescriptionTextBox.size = new-object system.drawing.size(250,23)
$DescriptionTextBox.tabIndex = 1
$DescriptionTextBox.text = "Description"   
$DescriptionTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$objform.controls.add($DescriptionTextBox)
$DescriptionTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{	
		#AddSetting
	}
})


#Pattern Label ============================================================
$PatternTextLabel = New-Object System.Windows.Forms.Label
$PatternTextLabel.Location = New-Object System.Drawing.Size(35,418) 
$PatternTextLabel.Size = New-Object System.Drawing.Size(50,15) 
$PatternTextLabel.Text = "Pattern:"
$PatternTextLabel.TabStop = $false
$PatternTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$objForm.Controls.Add($PatternTextLabel)

#Pattern Text box ============================================================
$PatternTextBox = New-Object System.Windows.Forms.TextBox
$PatternTextBox.location = new-object system.drawing.size(85,415)
$PatternTextBox.size = new-object system.drawing.size(250,23)
$PatternTextBox.tabIndex = 1
$PatternTextBox.text = "Pattern"   
$PatternTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$objform.controls.add($PatternTextBox)
$PatternTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{	
		#AddSetting
	}
})


#Translation Label ============================================================
$TranslationTextLabel = New-Object System.Windows.Forms.Label
$TranslationTextLabel.Location = New-Object System.Drawing.Size(15,443) 
$TranslationTextLabel.Size = New-Object System.Drawing.Size(65,15) 
$TranslationTextLabel.Text = "Translation:"
$TranslationTextLabel.TabStop = $false
$TranslationTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$objForm.Controls.Add($TranslationTextLabel)

#Setting Text box ============================================================
$TranslationTextBox = New-Object System.Windows.Forms.TextBox
$TranslationTextBox.location = new-object system.drawing.size(85,440)
$TranslationTextBox.size = new-object system.drawing.size(250,23)
$TranslationTextBox.tabIndex = 1
$TranslationTextBox.text = "Translation"   
$TranslationTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$objform.controls.add($TranslationTextBox )
$TranslationTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{	
		AddSetting
	}
})

#Add button
$AddButton = New-Object System.Windows.Forms.Button
$AddButton.Location = New-Object System.Drawing.Size(353,375)
$AddButton.Size = New-Object System.Drawing.Size(87,18)
$AddButton.Text = "Add / Edit"
$AddButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$AddButton.Add_Click(
{
	DisableAllButtons
	AddSetting
	EnableAllButtons
})
$objForm.Controls.Add($AddButton)


#Delete button
$DeleteButton = New-Object System.Windows.Forms.Button
$DeleteButton.Location = New-Object System.Drawing.Size(353,395)
$DeleteButton.Size = New-Object System.Drawing.Size(87,18)
$DeleteButton.Text = "Delete"
$DeleteButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$DeleteButton.Add_Click(
{
	DisableAllButtons
	DeleteSetting
	EnableAllButtons
})
$objForm.Controls.Add($DeleteButton)


#Add button
$DeleteAllButton = New-Object System.Windows.Forms.Button
$DeleteAllButton.Location = New-Object System.Drawing.Size(353,415)
$DeleteAllButton.Size = New-Object System.Drawing.Size(87,18)
$DeleteAllButton.Text = "Delete All"
$DeleteAllButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$DeleteAllButton.Add_Click(
{
	DisableAllButtons
	DeleteAllSettings
	EnableAllButtons
})
$objForm.Controls.Add($DeleteAllButton)


#Test Label ============================================================
$TestPhoneTextLabel = New-Object System.Windows.Forms.Label
$TestPhoneTextLabel.Location = New-Object System.Drawing.Size(50,483) 
$TestPhoneTextLabel.Size = New-Object System.Drawing.Size(30,15) 
$TestPhoneTextLabel.Text = "Test:"
$TestPhoneTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$TestPhoneTextLabel.TabStop = $false
$objForm.Controls.Add($TestPhoneTextLabel)

#Test Text box ============================================================
$TestPhoneTextBox = New-Object System.Windows.Forms.TextBox
$TestPhoneTextBox.location = new-object system.drawing.size(85,480)
$TestPhoneTextBox.size = new-object system.drawing.size(250,23)
$TestPhoneTextBox.tabIndex = 1
$TestPhoneTextBox.text = "0407532999"   
$TestPhoneTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$objform.controls.add($TestPhoneTextBox)
$TestPhoneTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{	
		DisableAllButtons
		TestPhoneNumberNew
		EnableAllButtons
	}
})

#Add button
$TestPhoneButton = New-Object System.Windows.Forms.Button
$TestPhoneButton.Location = New-Object System.Drawing.Size(353,480)
$TestPhoneButton.Size = New-Object System.Drawing.Size(87,18)
$TestPhoneButton.Text = "Test Number"
$TestPhoneButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$TestPhoneButton.Add_Click(
{
	DisableAllButtons
	TestPhoneNumberNew
	EnableAllButtons
})
$objForm.Controls.Add($TestPhoneButton)


#Pattern Label ============================================================
$TestPhonePatternTextLabel = New-Object System.Windows.Forms.Label
$TestPhonePatternTextLabel.Location = New-Object System.Drawing.Size(75,503) 
$TestPhonePatternTextLabel.Size = New-Object System.Drawing.Size(400,15) 
$TestPhonePatternTextLabel.Text = "Matched Pattern:"
$TestPhonePatternTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$TestPhonePatternTextLabel.TabStop = $false
$objForm.Controls.Add($TestPhonePatternTextLabel)

#Translation Label ============================================================
$TestPhoneTranslationTextLabel = New-Object System.Windows.Forms.Label
$TestPhoneTranslationTextLabel.Location = New-Object System.Drawing.Size(75,523) 
$TestPhoneTranslationTextLabel.Size = New-Object System.Drawing.Size(400,15) 
$TestPhoneTranslationTextLabel.Text = "Matched Translation:"
$TestPhoneTranslationTextLabel.TabStop = $false
$TestPhoneTranslationTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$objForm.Controls.Add($TestPhoneTranslationTextLabel)


#Result Label ============================================================
$TestPhoneResultTextLabel = New-Object System.Windows.Forms.Label
$TestPhoneResultTextLabel.Location = New-Object System.Drawing.Size(75,543) 
$TestPhoneResultTextLabel.Size = New-Object System.Drawing.Size(400,15) 
$TestPhoneResultTextLabel.Text = "Test Result:"
$TestPhoneResultTextLabel.TabStop = $false
$Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
$TestPhoneResultTextLabel.Font = $Font 
$TestPhoneResultTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$objForm.Controls.Add($TestPhoneResultTextLabel)


#Import Label ============================================================
$ImportTextLabel = New-Object System.Windows.Forms.Label
$ImportTextLabel.Location = New-Object System.Drawing.Size(50,578) 
$ImportTextLabel.Size = New-Object System.Drawing.Size(80,15) 
$ImportTextLabel.Text = "Import/Export:"
$ImportTextLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$ImportTextLabel.TabStop = $false
$objForm.Controls.Add($ImportTextLabel)


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
$objForm.Controls.Add($ImportButton)

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
$objForm.Controls.Add($ExportButton)


$ToolTip = New-Object System.Windows.Forms.ToolTip 
$ToolTip.BackColor = [System.Drawing.Color]::LightGoldenrodYellow 
$ToolTip.IsBalloon = $true 
$ToolTip.InitialDelay = 500 
$ToolTip.ReshowDelay = 500 
$ToolTip.AutoPopDelay = 10000
#$ToolTip.ToolTipTitle = "Help:"
$ToolTip.SetToolTip($ImportButton, "This button will allow you to import a Company_Phone_Number_Normalization_Rules.txt file.") 
$ToolTip.SetToolTip($ExportButton, "This button will export a file in Company_Phone_Number_Normalization_Rules.txt format.") 



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
     	
	$SiteDropDownBox = New-Object System.Windows.Forms.ComboBox 
	$SiteDropDownBox.Location = New-Object System.Drawing.Size(10,40) 
	$SiteDropDownBox.Size = New-Object System.Drawing.Size(300,20) 
	$SiteDropDownBox.DropDownHeight = 200 
	$SiteDropDownBox.tabIndex = 1
	$SiteDropDownBox.DropDownStyle = "DropDownList"

	Get-CsSite | select-object identity | ForEach-Object {[void] $SiteDropDownBox.Items.Add($_.identity)}
	
	$numberOfItems = $SiteDropDownBox.count
	if($numberOfItems -gt 0)
	{
		$SiteDropDownBox.SelectedIndex = 0
	}
	
	# Create the OK button.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(150,75)
    $okButton.Size = New-Object System.Drawing.Size(75,25)
    $okButton.Text = "OK"
    $okButton.Add_Click({ $form.Tag = $SiteDropDownBox.SelectedItem.ToString(); $form.Close() })
     
    # Create the Cancel button.
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Size(240,75)
    $cancelButton.Size = New-Object System.Drawing.Size(75,25)
    $cancelButton.Text = "Cancel"
    $cancelButton.Add_Click({ $form.Tag = $null; $form.Close() })
     
    # Create the form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = $WindowTitle
    $form.Size = New-Object System.Drawing.Size(350,150)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Topmost = $True
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    $form.ShowInTaskbar = $true
	[byte[]]$WindowIcon = @(66, 77, 56, 3, 0, 0, 0, 0, 0, 0, 54, 0, 0, 0, 40, 0, 0, 0, 16, 0, 0, 0, 16, 0, 0, 0, 1, 0, 24, 0, 0, 0, 0, 0, 2, 3, 0, 0, 18, 11, 0, 0, 18, 11, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114,0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0,198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 205, 132, 32, 234, 202, 160,255, 255, 255, 244, 229, 208, 205, 132, 32, 202, 123, 16, 248, 238, 224, 198, 114, 0, 205, 132, 32, 234, 202, 160, 255,255, 255, 255, 255, 255, 244, 229, 208, 219, 167, 96, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 248, 238, 224, 198, 114, 0, 198, 114, 0, 223, 176, 112, 255, 255, 255, 219, 167, 96, 198, 114, 0, 198, 114, 0, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 198, 114, 0, 198,114, 0, 248, 238, 224, 255, 255, 255, 244, 229, 208, 198, 114, 0, 198, 114, 0, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 198, 114, 0, 216, 158, 80, 255, 255, 255, 255, 255, 255, 252, 247, 240, 209, 141, 48, 198, 114, 0, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 198, 114, 0, 241, 220, 192, 255, 255, 255, 252, 247, 240, 212, 149, 64, 234, 202, 160, 198, 114, 0, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 205, 132, 32, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 248, 238, 224, 202, 123, 16, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 234, 202, 160, 255, 255, 255, 255, 255, 255, 205, 132, 32, 198, 114, 0, 223, 176, 112, 223, 176, 112, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 244, 229, 208, 252, 247, 240, 255, 255, 255, 237, 211, 176, 198, 114, 0, 198, 114, 0, 202, 123, 16, 248, 238, 224, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 212, 149, 64, 255, 255, 255, 255, 255, 255, 255, 255, 255, 212, 149, 64, 198, 114, 0, 198, 114, 0, 198, 114, 0, 234, 202, 160, 255, 255,255, 255, 255, 255, 241, 220, 192, 205, 132, 32, 198, 114, 0, 198, 114, 0, 205, 132, 32, 227, 185, 128, 227, 185, 128, 227, 185, 128, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 205, 132, 32, 227, 185, 128, 227, 185,128, 227, 185, 128, 219, 167, 96, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 0, 0)
	$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
	$form.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
     
    # Add all of the controls to the form.
    $form.Controls.Add($label)
    $form.Controls.Add($textBox)
    $form.Controls.Add($okButton)
    $form.Controls.Add($cancelButton)
	$form.Controls.Add($SiteDropDownBox)
     
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
     
    # Return the text that the user entered.
    return $form.Tag
}

function DisableAllButtons()
{
	$policyDropDownBox.Enabled = $false
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
}


function EnableAllButtons()
{
	$policyDropDownBox.Enabled = $true
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
}


function Move-Up
{
	foreach ($lvi in $lv.SelectedItems)
	{
		#GET SETTINGS OF SELECTED ITEM
		$item = $lv.Items[$lvi.Index]
		$itemValue = $item.SubItems

		[string]$Scope = $policyDropDownBox.SelectedItem.ToString()
		[string]$Name = $item.Text
		[string]$Priority = $itemValue[1].Text
		[string]$Description = $itemValue[2].Text
		[string]$Pattern = $itemValue[3].Text
		[string]$Translation = $itemValue[4].Text
						
		$orgIndex = $lvi.Index
		if($orgIndex -gt 0)
		{
			$index = $orgIndex - 1
			
			Remove-CsAddressBookNormalizationRule -Identity "${Scope}/${Name}"
			New-CsAddressBookNormalizationRule -Parent $Scope -Name $Name -Description $Description -Pattern $Pattern -Translation $Translation -Priority $index
			
			GetNormalisationPolicy
			
			$lv.Items[$index].Selected = $true
			$lv.Items[$index].EnsureVisible()
		}
		else
		{
			Write-Host "INFO: Cannot move item any lower..."
		}
	}

}


function Move-Down
{
	foreach ($lvi in $lv.SelectedItems)
	{
		#GET SETTINGS OF SELECTED ITEM
		$item = $lv.Items[$lvi.Index]
		$itemValue = $item.SubItems

		[string]$Scope = $policyDropDownBox.SelectedItem.ToString()
		[string]$Name = $item.Text
		[string]$Priority = $itemValue[1].Text
		[string]$Description = $itemValue[2].Text
		[string]$Pattern = $itemValue[3].Text
		[string]$Translation = $itemValue[4].Text
		
		
		$orgIndex = $lvi.Index
		if($orgIndex -lt ($lv.Items.Count - 1))
		{
			$index = $orgIndex + 1
			
			Remove-CsAddressBookNormalizationRule -Identity "${Scope}/${Name}"
			New-CsAddressBookNormalizationRule -Parent $Scope -Name $Name -Description $Description -Pattern $Pattern -Translation $Translation -Priority $index
			
			GetNormalisationPolicy
			
			$lv.Items[$index].Selected = $true
			$lv.Items[$index].EnsureVisible()
		}
		else
		{
			Write-Host "INFO: Cannot move item any higher..."
		}
	}
	
}

function GetNormalisationPolicy
{
	$lv.Items.Clear()
	
	$theIdentity = $policyDropDownBox.SelectedItem.ToString()
	Write-Host "Getting rules for $theIdentity"
	$NormRules = Get-CsAddressBookNormalizationRule -identity $theIdentity
	
	foreach($NormRule in $NormRules)
	{
		$id = $NormRule.Identity
		$idsplit = $id.Split("/")
		
		$Name = $idsplit[1]
		$Priority = $NormRule.Priority
		$Description = $NormRule.Description
		if($Description -eq $null)
		{
			$Description = "<Not Set>"
		}
		$Pattern = $NormRule.Pattern
		$Tranlation = $NormRule.Translation
		
		$lvItem = new-object System.Windows.Forms.ListViewItem($Name)
		$lvItem.ForeColor = "Black"
		
		[void]$lvItem.SubItems.Add($Priority)
		[void]$lvItem.SubItems.Add($Description)
		[void]$lvItem.SubItems.Add($Pattern)
		[void]$lvItem.SubItems.Add($Tranlation)
		
		[void]$lv.Items.Add($lvItem)
	}
}

function UpdateListViewSettings
{
	if($lv.SelectedItems.count -eq 0)
	{
		$NameTextBox.Text = ""
		$DescriptionTextBox.Text = ""
		$PatternTextBox.Text = ""
		$TranslationTextBox.Text = ""
	}
	else
	{
		foreach ($item in $lv.SelectedItems)
		{
			[string]$itemName = $item.Text
			$itemValue = $item.SubItems
			
			$NameTextBox.Text = $itemName
			
			[string]$settingValue1 = $itemValue[2].Text
			$DescriptionTextBox.Text = $settingValue1
			[string]$settingValue3 = $itemValue[3].Text
			$PatternTextBox.Text = $settingValue3
			[string]$settingValue4 = $itemValue[4].Text
			$TranslationTextBox.Text = $settingValue4
		}
	}
}

#Add / Edit an item
function AddSetting
{
	[string]$Scope = $policyDropDownBox.SelectedItem.ToString()
	[string]$Name = $NameTextBox.Text
	[string]$Description = $DescriptionTextBox.Text
	[string]$Pattern = $PatternTextBox.Text
	[string]$Translation = $TranslationTextBox.Text
	
	if($Name -ne "" -and $Name -ne $null)
	{
		if($Pattern -ne "" -and $Pattern -ne $null)
		{
			if($Translation -ne "" -and $Translation -ne $null)
			{
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
					
					$orgIndex = $lvi.Index
					$index = $orgIndex
					
					Write-Host "RUNNING: Remove-CsAddressBookNormalizationRule -Identity "${Scope}/${Name}"" -foreground "green"
					Remove-CsAddressBookNormalizationRule -Identity "${Scope}/${Name}"
					
					Write-Host "RUNNING: New-CsAddressBookNormalizationRule -Parent $Scope -Name $Name -Description $Description -Pattern $Pattern -Translation $Translation -Priority $index" -foreground "green"
					New-CsAddressBookNormalizationRule -Parent $Scope -Name $Name -Description $Description -Pattern $Pattern -Translation $Translation -Priority $Priority
					
					GetNormalisationPolicy
					
					$lv.Items[$Priority].Selected = $true
					$lv.Items[$Priority].EnsureVisible()
				}
				else   # ADD
				{
					if($lv.SelectedItems -ne $null)
					{
						foreach ($lvi in $lv.SelectedItems)
						{
	
							$index = $lvi.Index
								
							Write-Host "RUNNING: New-CsAddressBookNormalizationRule -Parent $Scope -Name $Name -Description $Description -Pattern $Pattern -Translation $Translation -Priority $index" -foreground "green"
							New-CsAddressBookNormalizationRule -Parent $Scope -Name $Name -Description $Description -Pattern $Pattern -Translation $Translation -Priority $index
							
							GetNormalisationPolicy
							
							$lv.Items[$index].Selected = $true
							$lv.Items[$index].EnsureVisible()
						}
					}
					else
					{
						$count = $lv.Items.Count
						Write-Host "RUNNING: New-CsAddressBookNormalizationRule -Parent $Scope -Name $Name -Description $Description -Pattern $Pattern -Translation $Translation -Priority $count" -foreground "green"
						New-CsAddressBookNormalizationRule -Parent $Scope -Name $Name -Description $Description -Pattern $Pattern -Translation $Translation -Priority $count
						
						GetNormalisationPolicy
						
						$lv.Items[$count].Selected = $true
						$lv.Items[$count].EnsureVisible()
					}
				}
			}
			else
			{
				Write-Host "ERROR: You need to up specify a translation. Please update the translation textbox and try again." -foreground "red"
			}
		}
		else
		{
			Write-Host "ERROR: You need to up specify a pattern. Please update the pattern textbox and try again." -foreground "red"
		}
	}
	else
	{
		Write-Host "ERROR: You need to up specify a name for the rule." -foreground "red"
	}
}


function DeleteSetting
{
	foreach ($lvi in $lv.SelectedItems)
	{
		#GET SETTINGS OF SELECTED ITEM
		$item = $lv.Items[$lvi.Index]
		$itemValue = $item.SubItems

		[string]$Scope = $policyDropDownBox.SelectedItem.ToString()
		[string]$Scope = $Scope.Replace("site:","")
		[string]$Name = $item.Text
		[string]$Priority = $itemValue[1].Text
		[string]$Description = $itemValue[2].Text
		[string]$Pattern = $itemValue[3].Text
		[string]$Translation = $itemValue[4].Text
		
		$orgIndex = $lvi.Index
		#Write-Host "ORG INDEX: $orgIndex"
		
		Write-Host "Removing: ${Scope}/${Name}"
		Remove-CsAddressBookNormalizationRule -Identity "${Scope}/${Name}"
		GetNormalisationPolicy
		
		#Write-Host "List View Count " $lv.Items.Count
		
		if($orgIndex -ge 1)
		{
			$index = $orgIndex - 1
			
			#Write-Host "Set INDEX: $index"
			
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

function DeleteAllSettings
{
	foreach ($lvi in $lv.Items)
	{
		[string]$Scope = $policyDropDownBox.SelectedItem.ToString()
		[string]$Name = $lvi.Text
		
		Remove-CsAddressBookNormalizationRule -Identity "${Scope}/${Name}"
	
	}
	GetNormalisationPolicy
}

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
	$Rules = Get-CsAddressBookNormalizationRule
	
	Write-Host ""
	Write-Host "-------------------------------------------------------------" -foreground "Green"
	Write-Host "TESTING: $PhoneNumber" -foreground "Green"
	Write-Host ""

	$TopLoopNo = 0
	$firstFound = $true
	foreach($item in $lv.Items)
	{
		$itemValue = $item.SubItems

		[string]$Pattern = $itemValue[3].Text
		[string]$Translation = $itemValue[4].Text
		
		#In version 1.02 don't remove the ";" char from the number being tested...
		$PhoneNumberStripped = $PhoneNumber.Replace(" ","").Replace("(","").Replace(")","").Replace("[","").Replace("]","").Replace("{","").Replace("}","").Replace(".","").Replace("-","").Replace(":","")
		
		#Write-Host "TESTING PATTERN: $Pattern" #DEBUG
		
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
				$item.ForeColor = "Red"
			}
			else
			{
				$item.ForeColor = "Blue"
			}
		}
	}
	$lv.Focus()	
	Write-Host "-------------------------------------------------------------" -foreground "Green"
}


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
		[string] $filename = $fileForm.FileName
		[string] $identity = $policyDropDownBox.SelectedItem.ToString()
		Import-CsCompanyPhoneNormalizationRules -Identity $identity -FileName $filename -Verbose -Confirm:$false
		
		GetNormalisationPolicy
	}
	else
	{
		Write-Host "Operation cancelled by user."
	}
	
}

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

			[string]$Pattern = $itemValue[3].Text
			[string]$Translation = $itemValue[4].Text
			[string]$Name = $item.Text
			[string]$Description = $itemValue[2].Text
			
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

GetNormalisationPolicy


# Activate the form ============================================================
$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()	


# SIG # Begin signature block
# MIIcZgYJKoZIhvcNAQcCoIIcVzCCHFMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUjqb5Oa0B+ZNFdrlp4y9MoG43
# CZ+ggheVMIIFHjCCBAagAwIBAgIQDGWW2SJRLPvqOO0rxctZHTANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE5MDIwNjAwMDAwMFoXDTIwMDIw
# NjEyMDAwMFowWzELMAkGA1UEBhMCQVUxDDAKBgNVBAgTA1ZJQzEQMA4GA1UEBxMH
# TWl0Y2hhbTEVMBMGA1UEChMMSmFtZXMgQ3Vzc2VuMRUwEwYDVQQDEwxKYW1lcyBD
# dXNzZW4wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDHPwqNOkuXxh8T
# 7y2cCWgLtpW30x/3rEUFnrlCv2DFgULLfZHFTd+HWhCiTUMHVESj+X8s+cmgKVWN
# bmEWPri590V6kfUmjtC+4/iKdVpvjgwrwAm6O6lHZ91y4Sn90A7eUV/EvUmGREVx
# uFk2s7jD/cYjTzm0fACQBuPz5sVjTzgFzbZMndPcptB8uEjtIF/k6BGCy7XyAMn6
# 0IncNguxGZBsS/CQQlsXlVhTnBn0QQxa7nRcpJQs/84OXjDypgjW6gVOf3hOzfXY
# rXNR54nqIh/VKFKz+PiEIW11yLW0608cI0xEE03yBOg14NGIapNBwOwSpeLMlQbH
# c9twu9BhAgMBAAGjggHFMIIBwTAfBgNVHSMEGDAWgBRaxLl7KgqjpepxA8Bg+S32
# ZXUOWDAdBgNVHQ4EFgQU2P05tP7466o6clrA//AUqWO4b2swDgYDVR0PAQH/BAQD
# AgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1UdHwRwMG4wNaAzoDGGL2h0dHA6
# Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEuY3JsMDWgM6Ax
# hi9odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNy
# bDBMBgNVHSAERTBDMDcGCWCGSAGG/WwDATAqMCgGCCsGAQUFBwIBFhxodHRwczov
# L3d3dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEEATCBhAYIKwYBBQUHAQEEeDB2
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTgYIKwYBBQUH
# MAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFNIQTJBc3N1
# cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMBAf8EAjAAMA0GCSqGSIb3DQEB
# CwUAA4IBAQCdaeq4xJ8ISuvZmb+ojTtfPN8PDWxDIsWos6e0KJ4sX7jYR/xXiG1k
# LgI5bVrb95YQNDIfB9ZeaVDrtrhEBu8Z3z3ZQFcwAudIvDyRw8HQCe7F3vKelMem
# TccwqWw/UuWWicqYzlK4Gz8abnSYSlCT52F8RpBO+T7j0ZSMycFDvFbfgBQk51uF
# mOFZk3RZE/ixSYEXlC1mS9/h3U9o30KuvVs3IfyITok4fSC7Wl9+24qmYDYYKh8H
# 2/jRG2oneR7yNCwUAMxnZBFjFI8/fNWALqXyMkyWZOIgzewSiELGXrQwauiOUXf4
# W7AIAXkINv7dFj2bS/QR/bROZ0zA5bJVMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1
# U0O1b5VQCDANBgkqhkiG9w0BAQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMM
# RGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQD
# ExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAwWhcN
# MjgxMDIyMTIwMDAwWjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQg
# SW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2Vy
# dCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG9w0B
# AQEFAAOCAQ8AMIIBCgKCAQEA+NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/5aid
# 2zLXcep2nQUut4/6kkPApfmJ1DcZ17aq8JyGpdglrA55KDp+6dFn08b7KSfH03sj
# lOSRI5aQd4L5oYQjZhJUM1B0sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxKhwjf
# DPXiTWAYvqrEsq5wMWYzcT6scKKrzn/pfMuSoeU7MRzP6vIK5Fe7SrXpdOYr/mzL
# fnQ5Ng2Q7+S1TqSp6moKq4TzrGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi6CxR
# 93O8vYWxYoNzQYIH5DiLanMg0A9kczyen6Yzqf0Z3yWT0QIDAQABo4IBzTCCAckw
# EgYDVR0TAQH/BAgwBgEB/wIBADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYI
# KwYBBQUHAwMweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2Nz
# cC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgw
# OqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJ
# RFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwTwYDVR0gBEgwRjA4BgpghkgBhv1sAAIE
# MCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCgYI
# YIZIAYb9bAMwHQYDVR0OBBYEFFrEuXsqCqOl6nEDwGD5LfZldQ5YMB8GA1UdIwQY
# MBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+7A1a
# JLPzItEVyCx8JSl2qB1dHC06GsTvMGHXfgtg/cM9D8Svi/3vKt8gVTew4fbRknUP
# UbRupY5a4l4kgU4QpO4/cY5jDhNLrddfRHnzNhQGivecRk5c/5CxGwcOkRX7uq+1
# UcKNJK4kxscnKqEpKBo6cSgCPC6Ro8AlEeKcFEehemhor5unXCBc2XGxDI+7qPjF
# Emifz0DLQESlE/DmZAwlCEIysjaKJAL+L3J+HNdJRZboWR3p+nRka7LrZkPas7CM
# 1ekN3fYBIM6ZMWM9CBoYs4GbT8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR6mhs
# RDKyZqHnGKSaZFHvMIIGajCCBVKgAwIBAgIQAwGaAjr/WLFr1tXq5hfwZjANBgkq
# hkiG9w0BAQUFADBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5j
# MRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBB
# c3N1cmVkIElEIENBLTEwHhcNMTQxMDIyMDAwMDAwWhcNMjQxMDIyMDAwMDAwWjBH
# MQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxJTAjBgNVBAMTHERpZ2lD
# ZXJ0IFRpbWVzdGFtcCBSZXNwb25kZXIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAw
# ggEKAoIBAQCjZF38fLPggjXg4PbGKuZJdTvMbuBTqZ8fZFnmfGt/a4ydVfiS457V
# WmNbAklQ2YPOb2bu3cuF6V+l+dSHdIhEOxnJ5fWRn8YUOawk6qhLLJGJzF4o9GS2
# ULf1ErNzlgpno75hn67z/RJ4dQ6mWxT9RSOOhkRVfRiGBYxVh3lIRvfKDo2n3k5f
# 4qi2LVkCYYhhchhoubh87ubnNC8xd4EwH7s2AY3vJ+P3mvBMMWSN4+v6GYeofs/s
# jAw2W3rBerh4x8kGLkYQyI3oBGDbvHN0+k7Y/qpA8bLOcEaD6dpAoVk62RUJV5lW
# MJPzyWHM0AjMa+xiQpGsAsDvpPCJEY93AgMBAAGjggM1MIIDMTAOBgNVHQ8BAf8E
# BAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCCAb8G
# A1UdIASCAbYwggGyMIIBoQYJYIZIAYb9bAcBMIIBkjAoBggrBgEFBQcCARYcaHR0
# cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzCCAWQGCCsGAQUFBwICMIIBVh6CAVIA
# QQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMA
# YQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4A
# YwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAA
# UwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAA
# QQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkA
# YQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIA
# YQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUA
# LjALBglghkgBhv1sAxUwHwYDVR0jBBgwFoAUFQASKxOYspkH7R7for5XDStnAs0w
# HQYDVR0OBBYEFGFaTSS2STKdSip5GoNL9B6Jwcp9MH0GA1UdHwR2MHQwOKA2oDSG
# Mmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENBLTEu
# Y3JsMDigNqA0hjJodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1
# cmVkSURDQS0xLmNybDB3BggrBgEFBQcBAQRrMGkwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1aHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5jcnQwDQYJKoZIhvcN
# AQEFBQADggEBAJ0lfhszTbImgVybhs4jIA+Ah+WI//+x1GosMe06FxlxF82pG7xa
# FjkAneNshORaQPveBgGMN/qbsZ0kfv4gpFetW7easGAm6mlXIV00Lx9xsIOUGQVr
# NZAQoHuXx/Y/5+IRQaa9YtnwJz04HShvOlIJ8OxwYtNiS7Dgc6aSwNOOMdgv420X
# Ewbu5AO2FKvzj0OncZ0h3RTKFV2SQdr5D4HRmXQNJsQOfxu19aDxxncGKBXp2JPl
# VRbwuwqrHNtcSCdmyKOLChzlldquxC5ZoGHd2vNtomHpigtt7BIYvfdVVEADkitr
# wlHCCkivsNRu4PQUCjob4489yq9qjXvc2EQwggbNMIIFtaADAgECAhAG/fkDlgOt
# 6gAK6z8nu7obMA0GCSqGSIb3DQEBBQUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0wNjExMTAwMDAwMDBa
# Fw0yMTExMTAwMDAwMDBaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IEFzc3VyZWQgSUQgQ0EtMTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
# ggEBAOiCLZn5ysJClaWAc0Bw0p5WVFypxNJBBo/JM/xNRZFcgZ/tLJz4FlnfnrUk
# FcKYubR3SdyJxArar8tea+2tsHEx6886QAxGTZPsi3o2CAOrDDT+GEmC/sfHMUiA
# fB6iD5IOUMnGh+s2P9gww/+m9/uizW9zI/6sVgWQ8DIhFonGcIj5BZd9o8dD3QLo
# Oz3tsUGj7T++25VIxO4es/K8DCuZ0MZdEkKB4YNugnM/JksUkK5ZZgrEjb7Szgau
# rYRvSISbT0C58Uzyr5j79s5AXVz2qPEvr+yJIvJrGGWxwXOt1/HYzx4KdFxCuGh+
# t9V3CidWfA9ipD8yFGCV/QcEogkCAwEAAaOCA3owggN2MA4GA1UdDwEB/wQEAwIB
# hjA7BgNVHSUENDAyBggrBgEFBQcDAQYIKwYBBQUHAwIGCCsGAQUFBwMDBggrBgEF
# BQcDBAYIKwYBBQUHAwgwggHSBgNVHSAEggHJMIIBxTCCAbQGCmCGSAGG/WwAAQQw
# ggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3LmRpZ2ljZXJ0LmNvbS9zc2wtY3Bz
# LXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUHAgIwggFWHoIBUgBBAG4AeQAgAHUA
# cwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQAaQBmAGkAYwBhAHQAZQAgAGMA
# bwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUAcAB0AGEAbgBjAGUAIABvAGYA
# IAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMAUAAvAEMAUABTACAAYQBuAGQA
# IAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEAcgB0AHkAIABBAGcAcgBlAGUA
# bQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkAdAAgAGwAaQBhAGIAaQBsAGkA
# dAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8AcgBwAG8AcgBhAHQAZQBkACAA
# aABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIAZQBuAGMAZQAuMAsGCWCGSAGG
# /WwDFTASBgNVHRMBAf8ECDAGAQH/AgEAMHkGCCsGAQUFBwEBBG0wazAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRw
# Oi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0Eu
# Y3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20v
# RGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsNC5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMB0GA1UdDgQW
# BBQVABIrE5iymQftHt+ivlcNK2cCzTAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYun
# pyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEARlA+ybcoJKc4HbZbKa9Sz1LpMUer
# Vlx71Q0LQbPv7HUfdDjyslxhopyVw1Dkgrkj0bo6hnKtOHisdV0XFzRyR4WUVtHr
# uzaEd8wkpfMEGVWp5+Pnq2LN+4stkMLA0rWUvV5PsQXSDj0aqRRbpoYxYqioM+Sb
# OafE9c4deHaUJXPkKqvPnHZL7V/CSxbkS3BMAIke/MV5vEwSV/5f4R68Al2o/vsH
# OE8Nxl2RuQ9nRc3Wg+3nkg2NsWmMT/tZ4CMP0qquAHzunEIOz5HXJ7cW7g/DvXwK
# oO4sCFWFIrjrGBpN/CohrUkxg0eVd3HcsRtLSxwQnHcUwZ1PL1qVCCkQJjGCBDsw
# ggQ3AgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNI
# QTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0ECEAxlltkiUSz76jjtK8XLWR0w
# CQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcN
# AQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUw
# IwYJKoZIhvcNAQkEMRYEFGAK+QWdiORIo0DJmYSq8sfvveYbMA0GCSqGSIb3DQEB
# AQUABIIBALvpQzsZY9YX+N2oaD8ueGSB5wEAMUhPhARgpRECtudnTgwfBVib6M7R
# wY+bcnyl8h+J/DXw4IAUlDj5vJMid4UhS41F6NhKVjM68R98kzmSs/iMadAhlbLJ
# 9/iBnyd9Qx40BNctT92+71fBmGyaP5febTpHkI3cz+NxSHuUAhEeTd03JXNaO7TC
# CgZkfPRSTL+TON1UuW6r8GT8qulKLZ5+kXQRvW+UysFHu72L6eBh8tC05CO+DFr/
# BKOZgrMGcaKTD7sYTigzeXfZMC04NnOFWE6oUmCCjkFKjk9h2Q9vj0icc1jHKNBY
# 1woYFPl+i4rSDEHoBmvxN76MzIH7v5ChggIPMIICCwYJKoZIhvcNAQkGMYIB/DCC
# AfgCAQEwdjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkw
# FwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1
# cmVkIElEIENBLTECEAMBmgI6/1ixa9bV6uYX8GYwCQYFKw4DAhoFAKBdMBgGCSqG
# SIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE5MTEyODEwNDQw
# MlowIwYJKoZIhvcNAQkEMRYEFGm4ic7HLsi/8INgWg6hSweH0kUxMA0GCSqGSIb3
# DQEBAQUABIIBAHfQbsI/GjtSS8DcF5BuOQ0IyU9HyOFI/HIF9i+9ANCrxvvWnxh1
# 4nZbMTrhAOvNUCNYxlm5xiETEHnoLHKwT6C+QI5iGPAGJBweqM/Kh2xFB6LksD81
# UsCFAc4SXp1A2hxD3vkJ6MJU5B5bK4m4LraRyGsAAd6DWvuYc61BELJEAbx29kGa
# sx/q8mN/mgPdp8xFYjLo8RnDGewQ3zytMneygFE8d2nTuO6c7+dS3jqvygwQIxCw
# OAVyIoon8p+IOF0NegCuSg4ISWaGnB4PJCn4L+HU3pgABba7Fxw3ipruSbblU80X
# 4ig1cpJLUFUtHrt9luGzGZyfHXSbRZ7dNmA=
# SIG # End signature block
