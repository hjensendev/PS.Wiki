#
#
# MIT License
#
# Copyright (c) 2018 Håkon Jensen
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

function Load-StoredVSphereCredential{
<#

.SYNOPSIS

Function for reading a encrypted System.Security.SecureString to a PSCredential object



.DESCRIPTION

This function creates a PSCredential object to be used when connecting to vSphere



.PARAMETER Account

The username associated with the password stored in the vsphere.cred file



.NOTES
    Version : 1.0  
    Author  : Håkon Jensen, Sicra AS  
       

#>
	Param(
		[Parameter(Mandatory=$true)][String]$Account
	)

	$pwFileName = "$($global:Config.Settings.LogFolder)\vsphere.cred"

	if (Test-Path $pwFileName){
		$PWsecure = Get-Content $pwFileName | ConvertTo-SecureString
	} else {
		if (!(Test-Path $global:Config.Settings.LogFolder)){
			Write-Host "Creating folder $($global:Config.Settings.LogFolder)"
			New-Item -ItemType Directory -Path $global:Config.Settings.LogFolder -Force
		}
		$PWSecure  = Read-Host "Enter password for your vSphere account $($Account)" -AsSecureString
		$PWEncyptedText = $PWSecure | ConvertFrom-SecureString		
		$result = Set-Content $PWFileName $PWEncyptedText 
	}

	[PSCredential]$cred = New-Object -TypeName System.Management.Automation.PSCredential ($Account, $PWsecure)
	$global:credVsphere = $cred
}









function Get-VsphereServerData{
<#

.SYNOPSIS

Function for reading virtual server configuration details from vSphere



.DESCRIPTION

This function returns a PSObject containing structured data that can be passed to Confluence.psm1/New-ConfluenceServerBody



.PARAMETER VmName

The virtual server name in vSphere from which to read configuration details



.NOTES
    Version : 1.0  
    Author  : Håkon Jensen, Sicra AS  
       

#>
	Param(
		[Parameter(Mandatory=$true)][String]$VmName
	)	
	
	# Create Result Obect
	$result = New-Object -TypeName PSObject

	$Error.Clear()
	# Get VM object
	$vm = Get-VM -Name $VmName -ErrorAction Stop

	# Get VM Config
	$vmConfig = Get-View -Id $vm.Id

	# Get Wiki Page Id
	$wikiPageId = Get-Annotation $vm -CustomAttribute WikiPageId
	
	# Get Funksjon
	$vmFunksjon = Get-Annotation $vm -CustomAttribute Funksjon

	# Get Forvaltningsansvarlig
	$vmForvaltningsansvarlig = Get-Annotation $vm -CustomAttribute Forvaltningsansvarlig

	# Get Tags
	$vmTags =  Get-TagAssignment $vm
	$collectionTags = New-Object System.Collections.ArrayList
	foreach ($vmTag in $vmTags){
		$collectionTags.Add($vmTag.Tag) | Out-Null
	}


	# Get Nics
	$vmNics = Get-NetworkAdapter -VM $VM
	$collectionNics = New-Object System.Collections.ArrayList
	foreach ($vmNic in $vmNics){
		$nic = New-Object -TypeName PSObject
		$nic | Add-Member -NotePropertyName MacAddress -NotePropertyValue $vmNic.MacAddress
		$nic | Add-Member -NotePropertyName NetworkName -NotePropertyValue $vmNic.NetworkName
		$nic | Add-Member -NotePropertyName Type -NotePropertyValue $vmNic.Type
		foreach ($guestNic in $vm.Guest.Nics){
			if ($vmNic.MacAddress -eq $guestNic.MacAddress){
				if ($guestNic.IPAddress) {$ip = [system.String]::Join(", ", $guestNic.IPAddress)}
				$nic | Add-Member -NotePropertyName IPAddress -NotePropertyValue $ip
			}
		}
		$collectionNics.Add($nic) | Out-Null
	}
	$collectionNics = $collectionNics | Sort-Object -Property NetworkName 

	# Get Disks
	$vmDisks = $vm.Guest.Disks
	$collectionDisks = New-Object System.Collections.ArrayList
	foreach ($vmDisk in $vmDisks){
		[int]$diskGB = $vmDisk.CapacityGB
		$disk = New-Object -TypeName PSObject
		$disk | Add-Member -NotePropertyName Path -NotePropertyValue $vmDisk.Path
		$disk | Add-Member -NotePropertyName CapacityGB -NotePropertyValue  $diskGB
		$collectionDisks.Add($disk) | Out-Null
	}
	$collectionDisks = $collectionDisks | Sort-Object -Property Path 

	# Add collected data to result object
	$result | Add-Member -NotePropertyName ServerName -NotePropertyValue $VmName
	$result | Add-Member -NotePropertyName Folder -NotePropertyValue $vm.Folder
	$result | Add-Member -NotePropertyName OperatingSystem -NotePropertyValue $vmConfig.Guest.GuestFullName
	$result | Add-Member -NotePropertyName Cores -NotePropertyValue $vm.NumCpu
	$result | Add-Member -NotePropertyName CoresPerSocket -NotePropertyValue $vm.CoresPerSocket
	$result | Add-Member -NotePropertyName MemoryGB -NotePropertyValue $vm.MemoryGB
	$result | Add-Member -NotePropertyName Notes -NotePropertyValue $vm.Notes
	$result | Add-Member -NotePropertyName Funksjon -NotePropertyValue $vmFunksjon.Value
	$result | Add-Member -NotePropertyName Forvaltningsansvarlig -NotePropertyValue $vmForvaltningsansvarlig.Value
	$result | Add-Member -NotePropertyName WikiPageId -NotePropertyValue $wikiPageId.Value
	$result | Add-Member -NotePropertyName Nics -NotePropertyValue $collectionNics
	$result | Add-Member -NotePropertyName Disks -NotePropertyValue $collectionDisks
	$result | Add-Member -NotePropertyName Tags -NotePropertyValue $collectionTags

	# Return result
	Write-Verbose $result
	return $result
}

Export-ModuleMember -Function Get-VsphereServerData, Load-StoredVSphereCredential