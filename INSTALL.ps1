
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] “Administrator”))
{
    Write-Warning “You do not have required permisions to run this script!`nPlease run Installation script as an Administrator”
    Break
}

$Version = "1.0.0"
$ModuleName = "PS.Wiki"
$Description = "Powershell Confluence Automation Module"
$Author = "Håkon Jensen"
$Company = "Sicra AS"
$InstallPath = Join-Path -Path $env:ProgramFiles -ChildPath "WindowsPowerShell\Modules\$($ModuleName)"
$NestedModules = @(
					"Confluence.psm1",
                     "vSphere.psm1",
                     "Helpers.psm1"
				)

#Remove Old FolderStructure
if (Test-Path $InstallPath){
	Write-Host "Remove Old Installation Folder"
	Remove-Item -Path $InstallPath -Recurse -Force
}

#Create Folder Structure
Write-Host "Create Directory"
mkdir $InstallPath | Out-Null

#Copy Modules
Write-Host "Copy Modules"
Copy-Item $PSScriptRoot\Modules\*.psm1 $InstallPath

#Copy Configuration
Write-Host "Copy Configuration"
Copy-Item "$($PSScriptRoot)\config.xml" $InstallPath\Config.xml

#Create Manifest
Write-Host "Create Manifest"
New-ModuleManifest -Path $InstallPath\$ModuleName.psd1 -ModuleVersion $Version  -Description $Description -PowerShellVersion 4.0 -Author $Author -RequiredModules ConfluencePS,VMware.VimAutomation.Core -NestedModules $NestedModules -CompanyName $Company