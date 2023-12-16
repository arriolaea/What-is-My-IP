# This PowerShell 5.1 script pulls local workstation/laptop network information and displays it in a GUI window.
# Not compatible with PowerShell 7.0 or above

# Copyright (C) 2021 Ernesto Arriola

# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

Add-Type -AssemblyName System.Windows.Forms

# Explicitly declaring variables.
$ErrorActionPreference = "SilentlyContinue"

# Pulls system public IP.
# Utilizes ipify.org's website which provides the public IP.

try {
    $MyPubIP = (Invoke-RestMethod -Uri "https://api64.ipify.org?format=json" -ErrorAction Stop).ip
} catch {
    $MyPubIP = "Failed to get IP"
}


# Network connector to pull system network info.
$WshNetwork = New-Object -ComObject WScript.Network

# Variables
$strComputer = "."
$strMsg = ""

# WMI connector to pull system service information.
$objWMIService = Get-WmiObject -Query "Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE"

# Sets object reference variables.
foreach ($IPConfig in $objWMIService) {
    if ($IPConfig.IPAddress -ne $null) {
        foreach ($IPAddress in $IPConfig.IPAddress) {
            if (-not ($IPAddress -like "*:*")) {
                $strMsg += "$IPAddress`n`t`t"
            }
        }
    }
}

# Create a GUI form
$form = New-Object System.Windows.Forms.Form

# Define the size, title and background color
$form.AutoSize   = $true
$form.Text       = "Network Information"
$form.Font     = 'Microsoft Sans Serif,12,style=Bold'
$form.BackColor  = "#ffffff"

# Other elemtents
$Description                     = New-Object system.Windows.Forms.Label
$Description.Text                = "Add a description to the user."
$Description.AutoSize            = $true
$Description.Font                = 'Microsoft Sans Serif,13'
$Description.Location = New-Object System.Drawing.Point(10, 0)

# Create labels to display information
$label1 = New-Object System.Windows.Forms.Label
$label1.Text = "System Domain:"
$label1.AutoSize = $true
$label1.Font     = 'Microsoft Sans Serif,12,style=Bold'
$label1.Location = New-Object System.Drawing.Point(10, 50)
$label1a = New-Object System.Windows.Forms.Label
$label1a.Text = "$($WshNetwork.UserDomain)"
$label1a.AutoSize = $true
$label1a.Font     = 'Microsoft Sans Serif,12'
$label1a.Location = New-Object System.Drawing.Point(180, 50)

$label2 = New-Object System.Windows.Forms.Label
$label2.Text = "User Name:"
$label2.AutoSize = $true
$label2.Font     = 'Microsoft Sans Serif,12,style=Bold'
$label2.Location = New-Object System.Drawing.Point(10, 75)
$label2a = New-Object System.Windows.Forms.Label
$label2a.Text = "$($WshNetwork.UserName)"
$label2a.AutoSize = $true
$label2a.Font     = 'Microsoft Sans Serif,12'
$label2a.Location = New-Object System.Drawing.Point(180, 75)

$label3 = New-Object System.Windows.Forms.Label
$label3.Text = "Computer Name:"
$label3.AutoSize = $true
$label3.Font     = 'Microsoft Sans Serif,12,style=Bold'
$label3.Location = New-Object System.Drawing.Point(10, 100)
$label3a = New-Object System.Windows.Forms.Label
$label3a.Text = "$($WshNetwork.ComputerName)"
$label3a.AutoSize = $true
$label3a.Font     = 'Microsoft Sans Serif,12'
$label3a.Location = New-Object System.Drawing.Point(180, 100)

$label4 = New-Object System.Windows.Forms.Label
$label4.Text = "Public IP Address:"
$label4.AutoSize = $true
$label4.Font     = 'Microsoft Sans Serif,12,style=Bold'
$label4.Location = New-Object System.Drawing.Point(10, 125)
$label4a = New-Object System.Windows.Forms.Label
$label4a.Text = "$($MyPubIP)"
$label4a.AutoSize = $true
$label4a.Font     = 'Microsoft Sans Serif,12'
$label4a.Location = New-Object System.Drawing.Point(180, 125)

$label5 = New-Object System.Windows.Forms.Label
$label5.Text = "Network IP(s):"
$label5.AutoSize = $true
$label5.Font     = 'Microsoft Sans Serif,12,style=Bold'
$label5.Location = New-Object System.Drawing.Point(10, 150)
$label5a = New-Object System.Windows.Forms.Label
$label5a.Text = "`t`t" + $strMsg + "`n`n`n"
$label5a.width    = 170
$label5a.Font     = 'Microsoft Sans Serif,12'
$label5a.Location = New-Object System.Drawing.Point(180, 150)

$label6 = New-Object System.Windows.Forms.Label
$label6.Text = "Service Desk:"
$label6.AutoSize = $true
$label6.Font = 'Microsoft Sans Serif,12,style=Bold'
$label6.Location = New-Object System.Drawing.Point(10, 180)

$label6a = New-Object System.Windows.Forms.Label
$label6a.Text = "(123) 456-7890"
$label6a.AutoSize = $true
$label6a.Font = 'Microsoft Sans Serif,12'
$label6a.Location = New-Object System.Drawing.Point(180, 180)

$form.controls.AddRange(@($Description, $label1, $label1a, $label2, $label2a, $label3, $label3a, $label4, $label4a, $label5, $label5a, $label6, $label6a))

# Display the form
[void]$form.ShowDialog()
