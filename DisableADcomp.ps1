<#Copyright 2021 Arthur Kahrman

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
#>
# Created by Arthur Kahrman

### Delete computers from AD from list for Kiosk returns ###


#--- Hide PowerShell Window ---#
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0) #0 hide

#---Check if Admin---#
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}
if ((Test-Admin) -eq $false)  {
   exit
}

#--- SET GUI PARAMETERS ---#
Add-Type -AssemblyName System.Windows.Forms

[System.Windows.Forms.Application]::EnableVisualStyles()

$LocalForm = New-Object System.Windows.Forms.Form
$LocalForm.ClientSize = '200,300'
$LocalForm.Text = "Disabler"
$LocalForm.BackColor = "#ffffff"
$localForm.FormBorderStyle = "FixedDialog"

$dateselect = New-Object System.Windows.Forms.DateTimePicker
$dateselect.Font = "Arial, 10"
$dateselect.CustomFormat = "yyyy-MM-dd"

$namelabel = New-Object System.Windows.Forms.label
$namelabel.Text = "Initials:"
$namelabel.Width = 60
$namelabel.Location = New-Object System.Drawing.Point(15,32)
$namelabel.Font = "Arial, 10"

$compbox = New-Object System.Windows.Forms.RichTextBox
$compbox.Width = 150; $compbox.Height = 200
$compbox.Font = "Arial, 10"
$compbox.Location = New-Object System.Drawing.Point(25,55)

$inibox = New-Object System.Windows.Forms.TextBox
$inibox.Width = 40
$inibox.Font = "Arial, 9"
$inibox.Location = New-Object System.Drawing.Point(90,30)
$inibox.ReadOnly = $True

$disablebutton = New-Object System.Windows.Forms.Button
$disablebutton.Text = "Disable Computers"
$disablebutton.Width = 150; $disablebutton.Height = 30
$disablebutton.Font = "Arial, 10"
$disablebutton.Location = New-Object System.Drawing.Point(25,260)


$SearchBaseForm = New-Object System.Windows.Forms.Form
$SearchBaseForm.MaximizeBox = $false
$SearchBaseForm.MinimizeBox = $false
$SearchBaseForm.ShowIcon = $false
$SearchBaseForm.Size = "800, 150"
$SearchBaseForm.TopMost = $true
$SearchBaseForm.Text = "Search Base"
$SearchBaseForm.FormBorderStyle = "FixedDialog"

$SearchBaseLabel = New-Object System.Windows.Forms.Label
$SearchBaseLabel.Font = "Arial, 13pt"
$SearchBaseLabel.Location = "103, 13"
$SearchBaseLabel.Size = "584, 30"
$SearchBaseLabel.Text = "Please Enter the Search Base for Active Directory or edit the one below:"

$SearchBaseTextBox = New-Object System.Windows.Forms.TextBox
$SearchBaseTextBox.Font = "Arial, 15pt"
$SearchBaseTextBox.Location = "27, 56"
$SearchBaseTextBox.Size = "636, 30"

$SearchBaseButton = New-Object System.Windows.Forms.Button
$SearchBaseButton.Font = "Arial, 10pt"
$SearchBaseButton.Size = "97, 29"
$SearchBaseButton.Location = "671, 58"
$SearchBaseButton.Text = "Confirm"


$ComputerPrefixForm = New-Object System.Windows.Forms.Form
$ComputerPrefixForm.MaximizeBox = $false
$ComputerPrefixForm.MinimizeBox = $false
$ComputerPrefixForm.ShowIcon = $false
$ComputerPrefixForm.Size = "800, 150"
$ComputerPrefixForm.TopMost = $true
$ComputerPrefixForm.Text = "Computer Prefix"
$ComputerPrefixForm.FormBorderStyle = "FixedDialog"

$ComputerPrefixLabel = New-Object System.Windows.Forms.Label
$ComputerPrefixLabel.Font = "Arial, 13pt"
$ComputerPrefixLabel.Location = "25, 13"
$ComputerPrefixLabel.Size = "800, 30"
$ComputerPrefixLabel.Text = "Please Enter the Computer Prefix EX: (ABC-*-*) would just need the last digits of the computer"

$ComputerPrefixTextbox = New-Object System.Windows.Forms.TextBox
$ComputerPrefixTextbox.Font = "Arial, 15pt"
$ComputerPrefixTextbox.Location = "27, 56"
$ComputerPrefixTextbox.Size = "636, 30"

$ComputerPrefixButton = New-Object System.Windows.Forms.Button
$ComputerPrefixButton.Font = "Arial, 10pt"
$ComputerPrefixButton.Size = "97, 29"
$ComputerPrefixButton.Location = "671, 58"
$ComputerPrefixButton.Text = "Confirm"

#--- Gather Information ---#
$currentcomp = $env:COMPUTERNAME
$Username = $env:USERNAME
$ADUsername = Get-ADUser $Username
[string]$FirstName = $ADUsername.GivenName
[String]$LastName = $ADUsername.Surname
[String]$Initial = $ADUsername.GivenName.Substring(0,1) + $ADUsername.Surname.Substring(0,1)
$inibox.Text = $Initial

#--- Functions ---#
function remove {
    $runtime = Get-Date -Format "YYYY-MM-dd - hh:mm"
    #If the user hasn't run the script before make them a new folder with a master log file for later use
    if (!(Test-Path -Path ".\Listed Devices\$username")) {
        New-Item -Path ".\Listed Devices" -ItemType Directory
        New-Item -Path ".\Listed Devices\$username" -ItemType Directory
        New-Item -Path ".\Listed Devices\$username\$initial.log" -ItemType File
        "#START OF LOG#" | Out-File ".\Listed Devices\$Username\$initial.log"
    }
    if ($compbox.Text.Contains("`*")){
        Write-Host "Wildcard Detected!"
        "$Username attempted to use a wildcard at $runtime" | Out-File ".\Listed Devices\$Username\$initial.log" -Append
        Return
    } else {
        Write-Host "No wildcard detected"
    }
    #If the user hasn't run the script before make them a new folder with a master log file for later use
    $comptime= $dateselect.Value.ToString("yyyy-MM-dd")
    
    $stringbox = $compbox.Text.Split("`n") | ?{$_.Trim() -ne ""}
    foreach ($compbox in $stringbox) {
        write-host "$compbox is being edited"
        $compname1 = $(Get-ADComputer -Filter "Name -Like '$prefix$compbox'" -SearchBase "$searchbase" -Properties Name).Name
        # Add line to split entries in the event that there are two names linked to the same asset
        $compname2 = $compname1 -split "\s" | ?{$_.Trim() -ne ""}
        Write-Host $compname1
        Write-Host $compname2
        foreach ($compname1 in $compname2) {
            if ($compname1 -ne $null -or '' -or "\s"){
            $msgbox = [System.Windows.Forms.MessageBox]::Show("Would you like to disable $compname1",'---WARNING---','YesNo')
            switch ($msgbox){ 'Yes' {SelectYes} 'No' {SelectNo}}
            } else {
                SelectElse #CURRENTLY NOT WORKING. COMPUTERS NOT IN AD ARE IGNORED AT THE MOMENT AND ARE NOT LOGGED
                Write-Host "Yes"
            } #If there is no computer linked to the asset name it will use selectelse
        }
    }
    #CURRENTLY NOT WORKING"---------------------NOT FOUND---------------------" | Out-File -Force "./Listed Devices/$username/ADList_$comptime.txt" -Append
    Get-Content  ADListnoaction_$comptime.tmp -Erroraction SilentlyContinue | Out-File -Force "./Listed Devices/$username/ADList_$comptime.txt" -Append -Erroraction SilentlyContinue
    "---------------------NOT ACTIONED---------------------" | Out-File -Force "./Listed Devices/$username/ADList_$comptime.txt" -Append
    Get-Content ADListnotremoved_$comptime.tmp -Erroraction SilentlyContinue | Out-File -Force "./Listed Devices/$username/ADList_$comptime.txt" -Append -Erroraction SilentlyContinue
    removetmp
    
}
function SelectYes {
    "$compname1 Has been actioned on $comptime by $Initial" | Out-File -Force "./Listed Devices/$username/ADList_$comptime.txt" -Append
    Get-ADComputer -Filter "Name -Like '$prefix$compbox'" -SearchBase "$searchbase" | Disable-ADAccount
    Set-ADComputer -Identity "$compname1" -Description "$comptime DISABLED (return) - $Initial"
}
function SelectNo {
    "$compname1 was not removed due to user selection" | Out-File ADListnotremoved_$comptime.tmp -Append
}
function SelectElse {
    "$compbox was not changed as it could not be found in AD or contains a wildcard" | Out-File ADListnoaction_$comptime.tmp -Append #CURRENTLY NOT FUNCTIONAL
}
function WriteSearchBase {
    $SearchBaseTextBox.Text | Out-File -FilePath ".\search_base.ini"
    $SearchBaseForm.Hide()
}
function WriteComputerPrefix {
    $ComputerPrefixTextBox.Text | Out-File -FilePath ".\computer_search_prefix.ini"
    $ComputerPrefixForm.Hide()
}
function removetmp {
    Remove-Item *$comptime.tmp -ErrorAction SilentlyContinue
}

#---Buttons---#
$disablebutton.Add_Click({ remove })
$SearchBaseButton.Add_Click({ WriteSearchBase })
$ComputerPrefixButton.Add_Click({ WriteComputerPrefix })

if (!(Test-Path -Path ".\search_base.ini")) {
        New-Item -Path ".\search_base.ini" -ItemType File
        $possiblesearchbase = $(Get-ADcomputer -Identity $currentcomp).DistinguishedName
        $SearchBaseTextBox.Text = [regex]::Match($possiblesearchbase,'(?<=,).*').Value
        $Add_To_SearchBase_Form = @($SearchBaseLabel , $SearchBaseTextBox, $SearchBaseButton)
        ForEach ($item in $Add_To_SearchBase_Form) {
            $SearchBaseForm.Controls.Add($item)
            }
        $SearchBaseButton.Add_Click({ WriteSearchBase })
        [void]$SearchBaseForm.ShowDialog()
        }
if (!(Test-Path -Path ".\computer_search_prefix.ini")) {
        New-Item -Path ".\computer_search_prefix.ini" -ItemType File
        
        $Add_To_Prefix_Form = @($ComputerPrefixLabel , $ComputerPrefixTextbox, $ComputerPrefixButton)
        ForEach ($item in $Add_To_Prefix_Form) {
            $ComputerPrefixForm.Controls.Add($item)
            }
        $ComputerPrefixButton.Add_Click({ WriteComputerPrefix })
        [void]$ComputerPrefixForm.ShowDialog()
        }

$prefix = Get-Content .\computer_search_prefix.ini
$searchbase = Get-Content .\search_base.ini

$Add_To_Form = @($disablebutton , $compbox , $dateselect , $namelabel , $inibox)
ForEach ($item in $Add_To_Form) {
    $LocalForm.Controls.Add($item)
}

[void]$LocalForm.ShowDialog()
