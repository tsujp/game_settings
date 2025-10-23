# sp = Set-ItemProperty
# gpv = Get-ItemPropertyValue
# gp = Get-ItemProperty

Import-Module BitsTransfer

# Output directory for all the stuff we (might) download.
$tjp_dir = Join-Path -Path ([Environment]::GetFolderPath("Desktop")) -ChildPath "tjp_scripty"

# Create output directory silently if it does not already exist.
[System.IO.Directory]::CreateDirectory($tjp_dir) | Out-Null

Write-Output "Temp downloads at: $tjp_dir"

function tjp-temp {
    param (
        [Parameter(Mandatory)]
        [string]$sub_path
    )

    return Join-Path -Path $tjp_dir -ChildPath $sub_path
}


# Check Steam settings.
if ((Test-path HKCU:\Software\Valve\Steam) -eq $true) {
    # https://stackoverflow.com/questions/15511809/how-do-i-get-the-value-of-a-registry-key-and-only-the-value-using-powershell/79128911#79128911
    $steam_lang_item = Get-Item -Path HKCU:\Software\Valve\Steam

    # Change Steam language to English if it isn't already.
    if ((gpv $steam_lang_item.PSPath -Name Language) -ne "english") {
        sp $steam_lang_item.PSPath -Name Language -Value "english" -Type String
        Write-Output "Set Steam language to English"
    }
}
# TODO: Handle Steam not installed?
Write-Output "Steam configured"


# Check Sublime Text is installed.
$sublime_text_bin = "sublime_text_build_4200_x64_setup.exe"
if (-not (get-package "Sublime Text" -ErrorAction SilentlyContinue)) {
    Write-Output "Sublime Text not found, attempting to install"

    Start-BitsTransfer "https://download.sublimetext.com/$sublime_text_bin" $tjp_dir

    Write-Progress -CurrentOperation "InstallSublime" ("Run Sublime Text installer ... ")
    Start-Process (tjp-temp $sublime_text_bin) -Wait -ArgumentList "/SILENT"

    Write-Progress -CurrentOperation "InstallSublime" ("Check Sublime Text is available ... ")
    if (-not (get-package "Sublime Text" -ErrorAction SilentlyContinue)) {
        Write-Output "Sublime Text still not available after install, aborting"
        exit 1
    }

    Write-Progress -CurrentOperation "InstallSublime" ("Check Sublime Text is available ... Done")
    Start-Sleep -Milliseconds 350
}
Write-Output "Sublime Text present"


# Check 7-Zip is installed.
$7zip_bin = "7z2501-x64.exe"
if (-not (get-package "7-Zip*" -ErrorAction SilentlyContinue)) {
    Write-Output "7-Zip not found, attempting to install"

    Start-BitsTransfer "https://www.7-zip.org/a/$7zip_bin" $tjp_dir

    Write-Progress -CurrentOperation "Install7zip" ("Run 7-Zip installer ... ")
    Start-Process (tjp-temp $7zip_bin) -Wait -ArgumentList "/S"

    Write-Progress -CurrentOperation "Install7zip" ("Check 7-Zip is available ... ")
    if (-not (get-package "7-Zip*" -ErrorAction SilentlyContinue)) {
        Write-Output "7-Zip still not available after install, aborting"
        exit 1
    }

    Write-Progress -CurrentOperation "Install7zip" ("Check 7-Zip is available ... Done")
    Start-Sleep -Milliseconds 350
}
Write-Output "7-Zip present"


# Check Nvidia driver version. 
$nvidia_driver_wmi = Get-WmiObject -ClassName Win32_VideoController | Where-Object { $_.Name -match "NVIDIA" }
$nvidia_driver_version = ($nvidia_driver_wmi.DriverVersion.Replace('.', '')[-5..-1] -join '').insert(3, '.')
Write-Output "Current Nvidia driver: $nvidia_driver_version"

# TODO: Mapping of PSID and PFID to cards

if (($nvidia_driver_version -lt 576.88)) {
    $uri = 'https://gfwsl.geforce.com/services_toolkit/services/com/nvidia/services/AjaxDriverService.php' +
    '?func=DriverManualLookup' +
    '&psid=131' +
    '&pfid=1070' +
    '&osID=135' +
    '&languageCode=1033' +
    '&isWHQL=1' +
    '&dch=1' +
    '&sort1=0' +
    '&numberOfResults=1'

    $res = Invoke-WebRequest -Uri $uri -Method GET -UseBasicParsing
    $payload = $res.Content | ConvertFrom-Json
    $latest_driver_version = $payload.IDS[0].downloadInfo.Version

    Write-Output "Latest Nvidia driver: $latest_driver_version"

    # Download that shit.
    $latest_driver_uri = $payload.IDS[0].downloadInfo.DownloadURL
    Start-BitsTransfer $latest_driver_uri $tjp_dir

    # Install that shit.
    $nvidia_driver_bin = Split-Path $latest_driver_uri -Leaf
    Write-Output $nvidia_driver_bin

    Write-Progress -CurrentOperation "InstallNvidiaDrivers" ("Run Nvidia installer ... ")
    # TODO: Arguments like not installing the nvidia app? Idk, seems to work fine as-is.
    Start-Process (tjp-temp $nvidia_driver_bin) -Wait -ArgumentList "-s"
    Write-Progress -CurrentOperation "InstallNvidiaDrivers" ("Run Nvidia installer ... ")
    Start-Sleep -Milliseconds 350
}


# Go to https://www.nvidia.com/en-gb/geforce/drivers/
# Use the form on that page, and there's lots of info e.g. it makes this request when I select the product series
# https://gfwsl.geforce.com/nvidia_web_services/controller.php?com.nvidia.services.Drivers.getMenuArrays/{%22pt%22:%221%22,%22pst%22:%22131%22,%22driverType%22:%22all%22}
# has a lot of json data in response
# that big json array allows us to get the psid and pfid programmatically
# Battlefield 6 requires at least driver version 576.88

# TODO: The first form field in the first json array is for Product Type, and is 1 for GeForce
        # {
        #     "id": 1,
        #     "menutext": "GeForce",
        #     "_explicitType": "MenuItemVO"
        # },

# TODO: So, look at the second array (which corresponds to "Product Series", as that's the second form item)
#       and search for GeForce RTX 50 Series (we can take our GPU name, cut the last two digits, and add "Series")
#       and we see the PSID is now 131
        # {
        #     "id": 131,
        #     "menutext": "GeForce RTX 50 Series",
        #     "_explicitType": "MenuItemVO"
        # },

# Note that if we need to re-query (as by default it only shows Notebook GPUs), we can send this GET request
# https://gfwsl.geforce.com/nvidia_web_services/controller.php?com.nvidia.services.Drivers.getMenuArrays/{%22pt%22:%221%22,%22pst%22:%22131%22,%22driverType%22:%22all%22}
# The url is the following query string:
# com.nvidia.services.Drivers.getMenuArrays/{"pt":"1","pst":"131","driverType":"all"}


# TODO: The same but for the third array, corresponds to "Product" (which is Product Family ID, pfid)
#  now the third array has non-notebook info, we can see the entry for NVIDIA GeForce RTX 5070 as
#  so our pfid is 1070
        # {
        #     "id": 1070,
        #     "menutext": "NVIDIA GeForce RTX 5070",
        #     "_explicitType": "MenuItemVO"
        # },

# TODO: The 4th array is null for some reason

# TODO: Windows 11 is 135
        # {
        #     "id": 135,
        #     "menutext": "Windows 11",
        #     "_explicitType": "MenuItemVO"
        # },

# TODO: English US is 1033
        # {
        #     "id": "1033",
        #     "menutext": "English (US)",
        #     "_explicitType": "MenuItemVO"
        # },

# Those get pumped to (as a filled form)
# https://gfwsl.geforce.com/services_toolkit/services/com/nvidia/services/AjaxDriverService.php?func=DriverManualLookup&psid=131&pfid=1070&osID=135&languageCode=1033&beta=0&isWHQL=0&dltype=-1&dch=1&upCRD=0&qnf=0&sort1=1&numberOfResults=10
# GET
# func
# DriverManualLookup
# psid
# 131
# pfid
# 1070
# osID
# 135
# languageCode
# 1033
# beta
# 0
# isWHQL
# 0
# dltype
# -1
# dch
# 1
# upCRD
# 0
# qnf
# 0
# sort1
# 1
# numberOfResults
# 10
# That responds with some weird data format which has the URL we want:
#    "https://us.download.nvidia.com/Windows/581.57/581.57-desktop-win10-win11-64bit-international-dch-whql.exe"


######################## Old stuff


# https://github.com/lord-carlos/nvidia-update/blob/master/nvidia.ps1

# Write-Host "Installing Sublime Text 4"
# Write-Progress -CurrentOperation "EnablingFeatureXYZ" ( "Enabling feature XYZ ... " )
# Start-Process "sublime_text_build_4200_x64_setup.exe" -Wait -ArgumentList "/SILENT"
# Write-Progress -CurrentOperation "EnablingFeatureXYZ" ( "Enabling feature XYZ ... Done" )

# Total time to sleep
# $start_sleep = 10

# Time to sleep between each notification
# $sleep_iteration = 2

# Write-Output ( "Sleeping {0} seconds ... " -f ($start_sleep) )
# for ($i=1 ; $i -le ([int]$start_sleep/$sleep_iteration) ; $i++) {
#     Start-Sleep -Seconds $sleep_iteration
#     Write-Progress -CurrentOperation ("Sleep {0}s" -f ($start_sleep)) ( " {0}s ..." -f ($i*$sleep_iteration) )
# }
# Write-Progress -CurrentOperation ("Sleep {0}s" -f ($start_sleep)) -Completed "Done waiting for X to finish"


# $sublime_text = "https://download.sublimetext.com/sublime_text_build_4200_x64.zip"
# $output = "$PSScriptRoot\Help\SublimeText.zip"
# (New-Object System.Net.WebClient).DownloadFile($sublime_text, $output)


# https://stackoverflow.com/questions/3896258/how-do-i-output-text-without-a-newline-in-powershell

# Pop-up Steam dialog to install Battlefield 6.
# Start-Process "steam://install/2807960"

# Save current location
# Push-Location

# Restore prior saved location
# Pop-Location