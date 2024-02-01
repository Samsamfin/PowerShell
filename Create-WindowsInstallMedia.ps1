#Requires -RunAsAdministrator
<#
    .SYNOPSIS
    Create Windows installation media with manufacturer and model specific drivers. WinPE drivers are also added to WinRE.

    Sami Törönen
    01.02.2024

    Version 1.5.1
    .DESCRIPTION
    For this to work, create the following folder structure. The script will use this structure by default, but if you choose to use
    different folder names, make sure to use the additional parameters described below to set the desired paths.
        C:\Temp\
            Model-Drivers\
            WindowsSource\
            WinPE-Drivers\

    -Copy model specific drivers into "Model-Drivers" folder. This folder will be read recursively.
    -Copy manufacturer specific Win PE drivers into "WinPE-Drivers" folder. This folder will be read recursively.
    -Copy Windows installation files into "WindowsSource" folder.
    .PARAMETER SKU
    Set the needed Windows SKU. If no parameter is given, Pro Education is used.
    Available SKUs can be checked with the following command:
    Get-WindowsImage -ImagePath 'Path\To\install.wim'
    .PARAMETER SplitImage
    Split the final install.wim to 3800 MB .SWM parts so that the installation media can be used with FAT32 formatted USB-media.
    .PARAMETER WindowsSourceFolder
    Path to Windows installation source files. Default path is C:\Temp\WindowsSource.
    .PARAMETER WinPEDriverFolder
    Path to WinPE drivers. Default path is C:\Temp\WinPE-Drivers.
    .PARAMETER ModelDriversFolder
    Path to device model drivers. Default path is C:\Temp\Model-Drivers.
    .EXAMPLE
    PS> Create-WindowsInstallMedia.ps1
    .EXAMPLE
    PS> Create-WindowsInstallMedia.ps1 -SKU "Windows 10 Pro"
    .EXAMPLE
    PS> Create-WindowsInstallMedia.ps1 -SKU "Windows 10 Pro" -SplitImage
    .EXAMPLE
    PS> Create-WindowsInstallMedia.ps1 -SKU "Windows 10 Pro" -WindowsSourceFolder "C:\Temp\Source"
#>

[CmdLetBinding()]
Param (
    [string]$SKU,
    [string]$WindowsSourceFolder = "C:\Temp\WindowsSource",
    [string]$WinPEDriverFolder = "C:\Temp\WinPE-Drivers",
    [string]$ModelDriversFolder = "C:\Temp\Model-Drivers",
    [switch]$SplitImage = $false
)

#Set split size
$SplitSize = "3800"

### DO NOT MODIFY BELOW THIS LINE ###

#Create mount folders if they do not exist
If (!(Test-Path "$($env:ProgramData)\Create-WindowsInstallMedia\Mount")){
    Mkdir "$($env:ProgramData)\Create-WindowsInstallMedia\Mount" > $null
}

If (!(Test-Path "$($env:ProgramData)\Create-WindowsInstallMedia\WinRE")){
    Mkdir "$($env:ProgramData)\Create-WindowsInstallMedia\WinRE" > $null
}

#Folder variables
$WindowsMountFolder = "$($env:ProgramData)\Create-WindowsInstallMedia\Mount"
$WinREMountFolder = "$($env:ProgramData)\Create-WindowsInstallMedia\WinRE"

#Begin if required folders are found
If((Test-Path $WinPEDriverFolder) -and (Test-Path $ModelDriversFolder) -and (Test-Path $WindowsMountFolder) -and (Test-Path $WinREMountFolder) -and (Test-Path $WindowsSourceFolder)){
    Write-Host "Required folders found" -ForegroundColor Green
    
    #Folder testing variables
    $WindowsSourceTest = Test-Path "$WindowsSourceFolder\sources\install.wim" -PathType leaf
    $PEDriverFolderCheck = (Get-ChildItem -Path $WinPEDriverFolder\*.inf -Force -Recurse | Where-Object {!$_.PSIsContainer} | Measure-Object).Count
    $ModelDriversFolderCheck = (Get-ChildItem -Path $ModelDriversFolder\*.inf -Force -Recurse | Where-Object {!$_.PSIsContainer} | Measure-Object).Count
    $MountFolderCheck = (Get-ChildItem -Path $WindowsMountFolder\* -Force -Recurse | Measure-Object).Count
    $REMountFolderCheck = (Get-ChildItem -Path $WinREMountFolder\* -Force -Recurse | Measure-Object).Count

    #Begin if both mount folders are empty
    If($MountFolderCheck -eq 0 -and $REMountFolderCheck -eq 0){

        #Begin if Windows source files are found
        If($WindowsSourceTest){
            Write-Host "Windows source files found in $WindowsSourceFolder" -ForegroundColor Green

            #Set SKU to Pro Education if variable was not given
            If(!($SKU)){
                Write-Host "SKU was not defined, will try Pro Education" -ForegroundColor Yellow
                $SKUDefault = Get-WindowsImage -ImagePath $WindowsSourceFolder\sources\install.wim | Where-Object {$_.ImageName -like "Windows * Pro Education"}
                $SKU = $SKUDefault.ImageName
            }

            #Begin if SKU matches
            If(Get-WindowsImage -ImagePath $WindowsSourceFolder\sources\install.wim | Where-Object {$_.ImageName -like "$SKU"}){
                Write-Host "$SKU found" -ForegroundColor Green

                #Begin if at least one driver folder has content
                If($PEDriverFolderCheck -ne 0 -or $ModelDriversFolderCheck -ne 0){
                    Write-Host "Drivers found" -ForegroundColor Green

                    #Inject WinPE drivers to boot.wim if driver folder has content
                    If($PEDriverFolderCheck -ne 0){

                        #Get image name for mounting the first boot.wim image
                        $BootWimIndex = Get-WindowsImage -ImagePath $WindowsSourceFolder\sources\boot.wim | Where-Object {$_.ImageIndex -eq 1}
                        $BootWimName = $BootWimIndex.ImageName

                        Write-Host "Mounting boot.wim, Index:1, $BootWimName" -ForegroundColor Green
                        try {
                            Dism /Mount-image /imagefile:$WindowsSourceFolder\sources\boot.wim /Index:1 /MountDir:$WindowsMountFolder  > $null
                        }
                        catch {
                            Write-host "Error encountered while mounting"$_.Exception.Message -ForegroundColor Red
                        }

                        Write-Host "Injecting $PEDriverFolderCheck WinPE driver(s) to boot.wim, Index:1, $BootWimName" -ForegroundColor Green
                        try {
                            Dism /Image:$WindowsMountFolder /Add-Driver /Driver:$WinPEDriverFolder /Recurse /ForceUnsigned  > $null
                        }
                        catch {
                            Write-host "Error encountered while injecting drivers"$_.Exception.Message -ForegroundColor Red
                        }

                        Write-Host "Committing changes to boot.wim, Index:1, $BootWimName" -ForegroundColor Green
                        try {
                            Dism /Unmount-Image /MountDir:$WindowsMountFolder /Commit > $null
                        }
                        catch {
                            Write-host "Error encountered while committing changes"$_.Exception.Message -ForegroundColor Red
                        }
                        
                        #Get image name for mounting the second boot.wim image
                        $BootWimIndex = Get-WindowsImage -ImagePath $WindowsSourceFolder\sources\boot.wim | Where-Object {$_.ImageIndex -eq 2}
                        $BootWimName = $BootWimIndex.ImageName
                        
                        Write-Host "Mounting boot.wim, Index:2, $BootWimName" -ForegroundColor Green
                        try {
                            Dism /Mount-image /imagefile:$WindowsSourceFolder\sources\boot.wim /Index:2 /MountDir:$WindowsMountFolder > $null
                        }
                        catch {
                            Write-host "Error encountered while mounting"$_.Exception.Message -ForegroundColor Red
                        }

                        Write-Host "Injecting $PEDriverFolderCheck WinPE driver(s) to boot.wim, Index:2, $BootWimName" -ForegroundColor Green
                        try {
                            Dism /Image:$WindowsMountFolder /Add-Driver /Driver:$WinPEDriverFolder /Recurse /ForceUnsigned > $null
                        }
                        catch {
                            Write-host "Error encountered while injecting drivers"$_.Exception.Message -ForegroundColor Red
                        }
                        
                        Write-Host "Committing changes to boot.wim, Index:2, $BootWimName" -ForegroundColor Green
                        try {
                            Dism /Unmount-Image /MountDir:$WindowsMountFolder /Commit > $null
                        }
                        catch {
                            Write-host "Error encountered while committing changes"$_.Exception.Message -ForegroundColor Red
                        }
                        
                    }
                    Else{
                        Write-Host "$WinPEDriverFolder folder does not contain any drivers, skipping boot.wim modification" -ForegroundColor Yellow
                    }

                    #Mount selected SKU from install.wim
                    If($ModelDriversFolderCheck -ne 0 -or $PEDriverFolderCheck -ne 0){
                        $Image = Get-WindowsImage -ImagePath $WindowsSourceFolder\sources\install.wim | Where-Object {$_.ImageName -like "$SKU"}
                        $Index = $Image.ImageIndex

                        Write-Host "Mounting install.wim, Index:$Index - $SKU" -ForegroundColor Green
                        try {
                            Dism /Mount-Image /imagefile:$WindowsSourceFolder\sources\install.wim /Index:$Index /MountDir:$WindowsMountFolder > $null
                        }
                        catch {
                            Write-host "Error encountered while mounting"$_.Exception.Message -ForegroundColor Red
                        }

                        #Mount Winre.wim and inject WinPE drivers if driver folder has content
                        If($PEDriverFolderCheck -ne 0){
                            $WinREWimIndex = Get-WindowsImage -ImagePath $WindowsMountFolder\Windows\System32\Recovery\winre.wim | Where-Object {$_.ImageIndex -eq 1}
                            $WinREWimName = $WinREWimIndex.ImageName

                            Write-Host "Mounting Winre.wim, Index:1, $WinREWimName" -ForegroundColor Green
                            try {
                                Dism /Mount-Wim /WimFile:$WindowsMountFolder\Windows\System32\Recovery\winre.wim /index:1 /MountDir:$WinREMountFolder > $null
                            }
                            catch {
                                Write-host "Error encountered while mounting"$_.Exception.Message -ForegroundColor Red
                            }

                            Write-Host "Injecting $PEDriverFolderCheck WinPE driver(s) to Winre.wim, Index:1, $WinREWimName" -ForegroundColor Green
                            try {
                                Dism /Image:$WinREMountFolder /Add-Driver /Driver:$WinPEDriverFolder /Recurse /ForceUnsigned > $null
                            }
                            catch {
                                Write-host "Error encountered while injecting drivers"$_.Exception.Message -ForegroundColor Red
                            }
                            
                            Write-Host "Cleanup Winre.wim, Index:1, $WinREWimName" -ForegroundColor Green
                            try {
                                Dism /Image:$WinREMountFolder /Cleanup-Image /StartComponentCleanup > $null
                            }
                            catch {
                                Write-host "Error encountered while cleaning up"$_.Exception.Message -ForegroundColor Red
                            }
                            
                            Write-Host "Committing changes to Winre.wim, Index:1, $WinREWimName" -ForegroundColor Green
                            try {
                                Dism /Unmount-Image /MountDir:$WinREMountFolder /Commit > $null
                            }
                            catch {
                                Write-host "Error encountered while committing changes"$_.Exception.Message -ForegroundColor Red
                            }
                            
                        }
                        Else{
                            Write-Host "$WinPEDriverFolder folder does not contain any drivers, skipping Winre.wim modification" -ForegroundColor Yellow
                        }

                        #Inject device model specific drivers to install.wim
                        If($ModelDriversFolderCheck -ne 0){
                            Write-Host "Injecting $ModelDriversFolderCheck driver(s) to install.wim, Index:$Index, $SKU" -ForegroundColor Green
                            try {
                                Dism /Image:$WindowsMountFolder /Add-Driver /Driver:$ModelDriversFolder /Recurse /ForceUnsigned > $null
                            }
                            catch {
                                Write-host "Error encountered while injecting drivers"$_.Exception.Message -ForegroundColor Red
                            }
                            
                        }
                        Else{
                            Write-Host "$ModelDriversFolder folder does not contain any drivers, skipping install.wim driver injection" -ForegroundColor Yellow
                        }
                        
                        #Save changes to install.wim
                        Write-Host "Committing changes to install.wim, Index:$Index, $SKU - this step will take 5-30 minutes" -ForegroundColor Green
                        try {
                            Dism /Unmount-Image /MountDir:$WindowsMountFolder /Commit > $null
                        }
                        catch {
                            Write-host "Error encountered while committing changes"$_.Exception.Message -ForegroundColor Red
                        }
                        
                        #Keep only the selected SKU
                        Write-Host "Exporting Index:$Index, $SKU" -ForegroundColor Green
                        try {
                            Dism /Export-Image /SourceImageFile:$WindowsSourceFolder\sources\install.wim /SourceIndex:$Index /DestinationImageFile:$WindowsSourceFolder\sources\temp.wim > $null
                        }
                        catch {
                            Write-host "Error encountered while exporting Index:$Index, $SKU"$_.Exception.Message -ForegroundColor Red
                        }
                        
                        #Remove install.wim and rename temp.wim to install.wim
                        Remove-Item -Path "$WindowsSourceFolder\sources\install.wim"
                        Rename-Item -Path "$WindowsSourceFolder\sources\temp.wim" -NewName "install.wim"

                        #Split install.wim
                        If($SplitImage -ne $false){
                            Write-Host "Splitting install.wim to $SplitSize MB parts" -ForegroundColor Green
                            try {
                                Dism /Split-Image /ImageFile:$WindowsSourceFolder\sources\install.wim /SWMFile:$WindowsSourceFolder\sources\install.swm /FileSize:$SplitSize > $null
                            }
                            catch {
                                Write-host "Error encountered while splitting install.wim".Exception.Message -ForegroundColor Red
                            }
                            
                            Remove-Item -Path "$WindowsSourceFolder\sources\install.wim"
                        }

                        #Cleanup
                        Remove-Item -Path "$($env:ProgramData)\Create-WindowsInstallMedia" -Recurse -Force

                        #Success
                        Write-Host "Media created successfully" -ForegroundColor Green
                    }
                }
                Else{
                    Write-Host "No drivers found. Please check folders $WinPEDriverFolder and $ModelDriversFolder." -ForegroundColor Red
                }
            }
            Else{
                Write-Host "Unable to find $SKU SKU." -ForegroundColor Red
            }
        }
        Else{
            Write-Host "Windows source files not found. Please check folder $WindowsSourceFolder." -ForegroundColor Red
        }
    }
    Else{
        Write-Host "DISM mounting folders are not empty. Check $WindowsMountFolder and $WinREMountFolder." -ForegroundColor Red    
    }
}
Else{
    Write-Host "Required folder(s) are not found. Please check the folder structure." -ForegroundColor Red
}