<#
    .SYNOPSIS
    Create Windows installation media with manufacturer and model specific drivers. WinPE drivers are also added to WinRE.

    Sami Törönen
    23.01.2024

    Version 1.2
    .DESCRIPTION
    For this to work, create the following folder structure. You can use different folder names, just don't
    forget to change the folder variables accordingly.
        C:\Temp\
            Model-Drivers\
            Mount\
            WindowsSource\
            WinPE-Drivers\
            WinRE\

    -Copy model specific drivers into "Model-Drivers" folder. This folder will be read recursively.
    -Copy manufacturer specific Win PE drivers into "WinPE-Drivers" folder. This folder will be read recursively.
    -Copy Windows installation files into "WindowsSource" folder.

    When running the script, please do not open "$WindowsMountFolder" or "$WinREMountFolder" folder while running
    the script, because this might cause errors while DISM tries to unmount images.
    .PARAMETER SKU
    Set the needed Windows SKU. If no parameter is given, Pro Education is used.
    Available SKUs can be checked with the following command:
    Get-WindowsImage -ImagePath 'Path\To\install.wim'
    .PARAMETER SplitImage
    Split the final install.wim to 3800MB .SWM parts so that the installation media can be used with FAT32 formatted USB-media.
    .EXAMPLE
    PS> Create-WindowsInstallMedia.ps1
    .EXAMPLE
    PS> Create-WindowsInstallMedia.ps1 -SKU "Windows 10 Pro"
    .EXAMPLE
    PS> Create-WindowsInstallMedia.ps1 -SKU "Windows 10 Pro" -SplitImage
#>

[CmdLetBinding()]
Param (
    [string]$SKU,
    [switch]$SplitImage = $false
)

#Folder variables
$WindowsSourceFolder = "C:\Temp\WindowsSource"
$WinPEDriverFolder = "C:\Temp\WinPE-Drivers"
$ModelDriversFolder = "C:\Temp\Model-Drivers"
$WindowsMountFolder = "C:\Temp\Mount"
$WinREMountFolder = "C:\Temp\WinRE"

#Begin if required folders are found
If((Test-Path $WinPEDriverFolder) -and (Test-Path $ModelDriversFolder) -and (Test-Path $WinPEDriverFolder) -and (Test-Path $ModelDriversFolder) -and (Test-Path $WindowsSourceFolder)){
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
                Write-Host "SKU was not defined, will try to use Pro Education" -ForegroundColor Yellow
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
                        $BootWimIndex = Get-WindowsImage -ImagePath $WindowsSourceFolder\sources\boot.wim | Where-Object {$_.ImageIndex -eq 1}
                        $BootWimName = $BootWimIndex.ImageName
                        Write-Host "Mounting boot.wim, Index:1, $BootWimName" -ForegroundColor Green
                        Dism /Mount-image /imagefile:$WindowsSourceFolder\sources\boot.wim /Index:1 /MountDir:$WindowsMountFolder
                        Write-Host "Injecting WinPE drivers to boot.wim, Index:1, $BootWimName" -ForegroundColor Green
                        Dism /Image:$WindowsMountFolder /Add-Driver /Driver:$WinPEDriverFolder /Recurse /ForceUnsigned
                        Write-Host "Committing changes to boot.wim, Index:1, $BootWimName" -ForegroundColor Green
                        Dism /Unmount-Image /MountDir:$WindowsMountFolder /Commit
                        $BootWimIndex = Get-WindowsImage -ImagePath $WindowsSourceFolder\sources\boot.wim | Where-Object {$_.ImageIndex -eq 2}
                        $BootWimName = $BootWimIndex.ImageName
                        Write-Host "Mounting boot.wim, Index:2, $BootWimName" -ForegroundColor Green
                        Dism /Mount-image /imagefile:$WindowsSourceFolder\sources\boot.wim /Index:2 /MountDir:$WindowsMountFolder
                        Write-Host "Injecting WinPE drivers to boot.wim, Index:2, $BootWimName" -ForegroundColor Green
                        Dism /Image:$WindowsMountFolder /Add-Driver /Driver:$WinPEDriverFolder /Recurse /ForceUnsigned
                        Write-Host "Committing changes to boot.wim, Index:2, $BootWimName" -ForegroundColor Green
                        Dism /Unmount-Image /MountDir:$WindowsMountFolder /Commit
                    }
                    Else{
                        Write-Host "$WinPEDriverFolder folder does not contain any drivers, skipping boot.wim modification" -ForegroundColor Yellow
                    }

                    #Mount selected SKU from install.wim
                    If($ModelDriversFolderCheck -ne 0 -or $PEDriverFolderCheck -ne 0){
                        $Image = Get-WindowsImage -ImagePath $WindowsSourceFolder\sources\install.wim | Where-Object {$_.ImageName -like "$SKU"}
                        $Index = $Image.ImageIndex
                        Write-Host "Mounting install.wim, Index:$Index - $SKU" -ForegroundColor Green
                        Dism /Mount-Image /imagefile:$WindowsSourceFolder\sources\install.wim /Index:$Index /MountDir:$WindowsMountFolder

                        #Mount Winre.wim and inject WinPE drivers if driver folder has content
                        If($PEDriverFolderCheck -ne 0){
                            $WinREWimIndex = Get-WindowsImage -ImagePath $WindowsMountFolder\Windows\System32\Recovery\winre.wim | Where-Object {$_.ImageIndex -eq 1}
                            $WinREWimName = $WinREWimIndex.ImageName
                            Write-Host "Mounting Winre.wim, Index:1, $WinREWimName" -ForegroundColor Green
                            Dism /Mount-Wim /WimFile:$WindowsMountFolder\Windows\System32\Recovery\winre.wim /index:1 /MountDir:$WinREMountFolder
                            Write-Host "Injecting WinPE drivers to Winre.wim, Index:1, $WinREWimName" -ForegroundColor Green
                            Dism /Image:$WinREMountFolder /Add-Driver /Driver:$WinPEDriverFolder /Recurse /ForceUnsigned
                            Write-Host "Cleanup Winre.wim, Index:1, $WinREWimName" -ForegroundColor Green
                            Dism /Image:$WinREMountFolder /Cleanup-Image /StartComponentCleanup
                            Write-Host "Committing changes to Winre.wim, Index:1, $WinREWimName" -ForegroundColor Green
                            Dism /Unmount-Image /MountDir:$WinREMountFolder /Commit
                        }
                        Else{
                            Write-Host "$WinPEDriverFolder folder does not contain any drivers, skipping Winre.wim modification" -ForegroundColor Yellow
                        }

                        #Inject device model specific drivers to install.wim
                        If($ModelDriversFolderCheck -ne 0){
                            Write-Host "Injecting drivers to install.wim, Index:$Index, $SKU" -ForegroundColor Green
                            Dism /Image:$WindowsMountFolder /Add-Driver /Driver:$ModelDriversFolder /Recurse /ForceUnsigned
                        }
                        Else{
                            Write-Host "$ModelDriversFolder folder does not contain any drivers, skipping install.wim driver injection" -ForegroundColor Yellow
                        }
                        
                        #Save changes to install.wim
                        Write-Host "Committing changes to install.wim, Index:$Index, $SKU - this step will take 5-30 minutes" -ForegroundColor Green
                        Dism /Unmount-Image /MountDir:$WindowsMountFolder /Commit

                        #Keep only the selected SKU
                        Write-Host "Exporting Index:$Index, $SKU" -ForegroundColor Green
                        Dism /Export-Image /SourceImageFile:$WindowsSourceFolder\sources\install.wim /SourceIndex:$Index /DestinationImageFile:$WindowsSourceFolder\sources\temp.wim
                        Remove-Item -Path "$WindowsSourceFolder\sources\install.wim"
                        Rename-Item -Path "$WindowsSourceFolder\sources\temp.wim" -NewName "install.wim"

                        #Split install.wim at 3800MB
                        If($SplitImage -ne $false){
                            Write-Host "Splitting install.wim to 3800MB parts" -ForegroundColor Green
                            Dism /Split-Image /ImageFile:$WindowsSourceFolder\sources\install.wim /SWMFile:$WindowsSourceFolder\sources\install.swm /FileSize:3800
                            Remove-Item -Path "$WindowsSourceFolder\sources\install.wim"
                        }

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