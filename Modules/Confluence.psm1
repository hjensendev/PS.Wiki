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

function Load-StoredWikiCredential{
<#

.SYNOPSIS

Function for reading a encrypted System.Security.SecureString to a PSCredential object



.DESCRIPTION

This function creates a PSCredential object to be used when connecting to Confluence



.PARAMETER Account

The username associated with the password stored in the wiki.cred file



.NOTES
    Version : 1.0  
    Author  : Håkon Jensen, Sicra AS  
       

#>
	Param(
		[Parameter(Mandatory=$true)][String]$Account
	)

	$pwFileName = "$($global:Config.Settings.LogFolder)\wiki.cred"

	if (Test-Path $pwFileName){
		$PWsecure = Get-Content $pwFileName | ConvertTo-SecureString
	} else {
		if (!(Test-Path $global:Config.Settings.LogFolder)){
			Write-Host "Creating folder $($global:Config.Settings.LogFolder)"
			New-Item -ItemType Directory -Path $global:Config.Settings.LogFolder -Force
		}
		$PWSecure  = Read-Host "Enter password for your Wiki account $($Account)" -AsSecureString
		$PWEncyptedText = $PWSecure | ConvertFrom-SecureString		
		$result = Set-Content $PWFileName $PWEncyptedText 
	}

	[PSCredential]$cred = New-Object -TypeName System.Management.Automation.PSCredential ($Account, $PWsecure)
	$global:credWiki = $cred
}









function Convert-CollectionToConfluenceTable{
<#

.SYNOPSIS

Function for converting a collection (hashtable) to a confluence table



.DESCRIPTION

This function returns html table of css class confluenceTable. 



.PARAMETER Collection

The collection that is to be convrertred to a table



.PARAMETER SortDescending

A switch that specifies that the table headers are sorted descending



.NOTES
    Version : 1.0  
    Author  : Håkon Jensen, Sicra AS  
       

#>

	Param(
		[Parameter(Mandatory=$true)][Object]$Collection,
		[Switch]$SortDescending
	)	

	# Check that the collection actually contains data
	if ($Collection -eq $null){
		return [string]::Empty
	}

	# Create Table
	$table = "<table class='confluenceTable'><colgroup><col/><col/></colgroup><tbody><tr>"

	# Create Headers
	if ($SortDescending) {
		$headers = Get-Member -InputObject $Collection[0] -MemberType NoteProperty |Sort -Property Name -Descending
	}
	else {
		$headers = Get-Member -InputObject $Collection[0] -MemberType NoteProperty |Sort -Property Name
	}
	
	foreach ($header in $headers){
		$table = $table + "<th class='confluenceTh'>$($header.Name)</th>"
	}
	$table = $table + "</tr>"

	# Create Rows
	foreach ($item in $Collection){
		if ($SortDescending) {
			$row = Get-Member -InputObject $item -MemberType NoteProperty |Sort -Property Name -Descending
		}
		else {
			$row = Get-Member -InputObject $item -MemberType NoteProperty
		}
		$table = $table + "<tr>"
		foreach($td in $row){
			[string]$value = $td.Definition.ToString()
			$value = $value.SubString($value.IndexOf("=")+1)
			$table = $table + "<td class='confluenceTd'>$($value)</td>"
		}
		$table = $table + "</tr>"
	}

	# Close Table
	$table = $table + "</tbody></table>"
	return $table
}









function Get-VmRunningEnvironment{
<#

.SYNOPSIS

Function for creating a html span of confluence css class status-macro



.DESCRIPTION

This function returns html span with different background color based on the input 



.PARAMETER Folder

String that will be used for determening span background color and text



.NOTES
    Version : 1.0  
    Author  : Håkon Jensen, Sicra AS  
       

#>
	Param(
		[Parameter(Mandatory=$true)][String]$Folder
	)	

	$prod = "<span class='status-macro aui-lozenge aui-lozenge-success conf-macro output-inline' data-hasbody='false' data-macro-name='status'>Produksjon</span>"
	$stag = "<span class='status-macro aui-lozenge aui-lozenge-current conf-macro output-inline' data-hasbody='false' data-macro-name='status'>Staging</span>"
	$dev = "<span class='status-macro aui-lozenge aui-lozenge-complete conf-macro output-inline' data-hasbody='false' data-macro-name='status'>Development</span>"
	$unknown = "<span class='status-macro aui-lozenge aui-lozenge-error conf-macro output-inline' data-hasbody='false' data-macro-name='status'>$($Folder)</span>"

	switch ($Folder) {
		"PRODuction" {return $prod}
		"STAGing" {return $stag}
		"Development" {return $dev}
		default { return $unknown }
	}
}









function New-ConfluenceServerBody{
<#

.SYNOPSIS

Function for creating a html body containing server configuration 



.DESCRIPTION

This function returns html code based on a definde PSObject passed in the ServerData parameter



.PARAMETER ServerData

PSObject that contians the server configuration data collected by the data collector function



.NOTES
    Version : 1.0  
    Author  : Håkon Jensen, Sicra AS  
       

#>
	Param(
		[Parameter(Mandatory=$true)][Object]$ServerData
	)	

	Write-Verbose "Create new body"

	if ([string]::IsNullOrEmpty($global:template)){
	    # Read Server Template and store it in a global variable for faster processing of multiple servers
	    # The Template PageID is retreived by using Get-ConfluencePage on the template page.
		Write-Verbose "Loading template"
		$global:template =  Get-ConfluencePage -PageID $global:Config.Settings.ConfluenceServerTemplatePageId -ApiURi "$($global:Config.Settings.ConfluenceServerBaseURI)/rest/api" -Credential $global:credWiki
	}

	if ([string]::IsNullOrEmpty($global:template)){
		Write-Error "Unable to read template page from $($global:Config.Settings.ConfluenceServerBaseURI)" -ErrorAction Stop
	}
	$body = $global:template.Body

	# Get all single valued properties
	$body = $body.Replace("(ServerName)",$ServerData.ServerName)
	$body = $body.Replace("(OperatingSystem)",$ServerData.OperatingSystem)
	$body = $body.Replace("(Environment)", (Get-VmRunningEnvironment -Folder $ServerData.Folder))
	$body = $body.Replace("(Cores)",$ServerData.Cores)
	$body = $body.Replace("(CoresPerSocket)",$ServerData.CoresPerSocket)
	$body = $body.Replace("(MEMSize)",$ServerData.MemoryGB)
	$body = $body.Replace("(Notes)",$ServerData.Notes)
	$body = $body.Replace("(Forvaltningsansvarlig)",$ServerData.Forvaltningsansvarlig)
	$body = $body.Replace("(Funksjon)",$ServerData.Funksjon)


	# Get Tags
	foreach ($tag in $ServerData.Tags){
		Write-Verbose "Processing tag $($tag)"

		if ($tag.Category.Name -eq "Backup"){
			if (![string]::IsNullOrEmpty($backupJobs)) {
				$backupJobs = $backupJobs + ", "
			}
			$backupJobs = $backupJobs + $tag.Name.ToString()
		}

		if ($tag.Category.Name -eq "Tjeneste"){
			if (![string]::IsNullOrEmpty($tjenester)) {
				$tjenester = $tjenester + ", "
			}
			$tjenester = $tjenester + $tag.Name.ToString()
		}
	}
	$body = $body.Replace("(BackupPlan)",$backupJobs)
	$body = $body.Replace("(Tjenester)",$tjenester)

	# Get Storage
	if ($ServerData.Disks) {$tableDisks = Convert-CollectionToConfluenceTable -Collection $ServerData.Disks -SortDescending}
	$body = $body.Replace("(Disks)",$tableDisks)

	# Get Nics
	if ($ServerData.Disks) {$tableNics = Convert-CollectionToConfluenceTable -Collection $ServerData.Nics}
	$body = $body.Replace("(NICS)",$tableNics)

	# Return result
	Write-Verbose $body
	return $body
}









function Add-TjenesteLabel{
<#

.SYNOPSIS

Function for appending labels stored as tjeneste-tags to existing label collection 



.DESCRIPTION

This function is used to add/append tjeneste-labels to the servers wiki page



.PARAMETER ServerData

PSObject that contians the server configuration data collected by the data collector function



.PARAMETER ExistingLabels

PSObject that contians the existing labels on the wiki page



.NOTES
    Version : 1.0  
    Author  : Håkon Jensen, Sicra AS  
       

#>
	Param(
		[Parameter(Mandatory=$true)][Object]$ServerData,
		[Parameter(Mandatory=$true)][Object]$ExistingLabels
	)

	foreach ($tag in $ServerData.Tags){
		if ($tag.Category.Name -eq "Tjeneste"){
			$ExistingLabels = $ExistingLabels + "servertjeneste-$($tag.Name)"
		}
	}
	Write-Verbose "Attaching labels: $($ExistingLabels)"
	return $ExistingLabels
}









function Process-Server{
<#

.SYNOPSIS

Function for processing server configuration and updating Wiki page if configuration has changed



.DESCRIPTION

This function creates a new wiki page and a hash from the collected data. The hash is compared with existing page hash.
If the hash is different, the the old wiki page is overwritten with the new configuration.



.PARAMETER vm

Object that contain the virtual server object to be processed.



.NOTES
    Version : 1.0  
    Author  : Håkon Jensen, Sicra AS  
       

#>
	Param(
		[Parameter(Mandatory=$true)][Object]$Vm
	)	
	
	# Write information to log
	Write-Log -Text "Processing $($vm.name)"

	# Collect current server data from vSphere
	Write-Verbose "Get server data from vSphere"
	$serverData = Get-VsphereServerData -VmName $Vm


	# Create new page body with newly collected vSphere data
	$newBody = New-ConfluenceServerBody -ServerData $serverData
	$newBodyHash = Get-StringHash -String $newBody
	Write-Verbose "Configruation hash:$($newBodyHash)"

	# Write infomation to log
	Write-Log -Text "PageId: $($serverData.WikiPageId), Hash: $($newBodyHash)"


	# If this server has a page in the wiki, the vm object in vSphere should contain a property containing the Wiki Page ID
	if (! ([string]::IsNullOrEmpty($serverData.WikiPageId))){
		# Use the PageID if present
		Write-Verbose "Read PageId $($serverData.WikiPageId) from wiki"
		$existingPage = Get-ConfluencePage -ApiURi "$($global:Config.Settings.ConfluenceServerBaseURI)/rest/api" -Credential $global:credWiki -PageID $serverData.WikiPageId -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
	}
	


	# Check if the server has an existing page in Wiki

	if ($existingPage -eq $null)
	{
		# This server did not have an existing page, create a new page
		Write-Log -Text "Creating new page"
		$newPage = New-ConfluencePage -SpaceKey $global:Config.Settings.ConfluenceSpaceKey -Title $Vm -Credential $global:credWiki -ApiURi "$($global:Config.Settings.ConfluenceServerBaseURI)/rest/api" -Body $newBody -ParentID $global:Config.Settings.ConfluenceServerParentPageId 
		Write-Verbose "Created new page with PageId $($newPage.ID)"

		try{
			# Update vSphere object with the Wiki PageId
			$result = Set-Annotation $vm -CustomAttribute "WikiPageId" -Value $newPage.ID -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

			# Set standard Labels
			$standardLabels = @("md5/$($newBodyHash)","server","autocreated")
			$labels = Add-TjenesteLabel -ServerData $serverData -ExistingLabels $standardLabels
			$labels = Set-ConfluenceLabel -apiURi "$($global:Config.Settings.ConfluenceServerBaseURI)/rest/api" -Credential $global:credWiki -PageID $newPage.ID -Label $labels -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
		}
		catch {
			# An error occured when creating new page
			# This is most likely caused by an invalid WikiPageId on the vSphere server object
			Write-Log -Text "ERROR: An error occured when creating new page $($vm)."

			# Find page with conflicting name
			$existingPage = Get-ConfluencePage -ApiURi "$($global:Config.Settings.ConfluenceServerBaseURI)/rest/api" -Credential $global:credWiki -Title $vm -SpaceKey $global:Config.Settings.ConfluenceSpaceKey -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
			if ($existingPage -eq $null){
				Write-Log -Text "ERROR: Unknow error"
			} 
			else {
				# Update existing page with automation-error label
				Write-Log -Text "ERROR: Invalid WikiPageId $($serverData.WikiPageId) on vSphere server $($Vm)"
				$labels = Set-ConfluenceLabel -apiURi "$($global:Config.Settings.ConfluenceServerBaseURI)/rest/api" -Credential $global:credWiki -PageID $existingPage.ID -Label @("automation-error","server","autocreated")
			}
		}
	}

	# Update exsting page
	else
	{
		# Get labels for existing page
		Write-Verbose "Read labels on PageID $($existingPage.ID)"
		$labels = Get-ConfluenceLabel -apiURi "$($global:Config.Settings.ConfluenceServerBaseURI)/rest/api" -Credential $global:credWiki -PageID $existingPage.ID

		# Find label containing MD5 hash
		foreach ($label in $labels.Labels){
			if ($label.ToString().StartsWith("md5/")){
				$existingBodyHash = $label.ToString().SubString(4)
			}
		}

		# Compare new wiki page body with existing page body
		if ($existingBodyHash -eq $newBodyHash){
			# No change for this server
			Write-Log "No change"
		} 
		else 
		{
			# Update page for this server
			Write-Log -Text "Update page"
			try {
				# Update page with new data
				$updatedPage = Set-ConfluencePage -PageID $existingPage.ID -Body $newBody -Credential $global:credWiki -ApiURi "$($global:Config.Settings.ConfluenceServerBaseURI)/rest/api"  -Title $vm -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

				# Replace labels on page for this server
				$standardLabels = @("md5/$($newBodyHash)","server","autocreated")
				$labels = Add-TjenesteLabel -ServerData $serverData -ExistingLabels $standardLabels
				$labels = Set-ConfluenceLabel -apiURi "$($global:Config.Settings.ConfluenceServerBaseURI)/rest/api" -Credential $global:credWiki -PageID $updatedPage.ID -Label $labels
			}
			catch {
				Write-Log -Text "ERROR: An error occured when updating page $($vm) with PageID $($existingPage.ID)"
				$result = Set-ConfluenceLabel -apiURi "$($global:Config.Settings.ConfluenceServerBaseURI)/rest/api" -Credential $global:credWiki -PageID $existingPage.ID -Label @("automation-error","server","autocreated")
			}
		}
	}
}









function Start-WikiAutomation {
<#

.SYNOPSIS

Function for staring the automation process


.DESCRIPTION

This function starts the automation process that reads existing pages and compares server configuration.



.PARAMETER Filter

String that contians a filter to use when fetching server objects



.NOTES
    Version : 1.0  
    Author  : Håkon Jensen, Sicra AS  
       

#>
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory=$false)][Object]$Filter
	)

    [datetime]$scriptStart = Get-Date
	Write-Log -Text "Script started"


	# Ignore Self Signed Certificate
	$result = Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -ParticipateInCEIP $false -Confirm: $false | Out-Null
	
	# Connect to VSphere
	Write-Log -Text "Connecting to $($global:Config.Settings.vSphereServer)"
	$viConn = Connect-VIServer -Server $global:Config.Settings.vSphereServer  -Credential $global:credVsphere -Force | Out-Null

	# Get all servers from vSphere
	if ([string]::IsNullOrEmpty($Filter)){
		$vms = Get-VM
	} 
	else {
		$vms = Get-VM $Filter
	}
	
	Foreach ($vm in $vms){
		If ($vm.PowerState -eq "PoweredOn"){
			Process-Server -Vm $vm
		}	
	}

	[Timespan]$scriptRunningTime = New-TimeSpan -Start $scriptStart -End (Get-Date)
	Write-Log -Text "Script complete in $($scriptRunningTime.Hours) hours $($scriptRunningTime.Minutes) minutes $($scriptRunningTime.Seconds) seconds "
}









Export-ModuleMember -Function Start-WikiAutomation, New-ConfluenceServerBody, Load-StoredWikiCredential