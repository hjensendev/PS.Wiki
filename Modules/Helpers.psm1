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

Function Get-StringHash()
{
<#

.SYNOPSIS

FUnction for hashing a string.



.DESCRIPTION

This functions returns a hash of the input string.
The hash is created using System.Security.Cryptography and the hash type can be specified with the Hashname parameter.
The default hash is MD5.



.PARAMETER String

The string to be hashed 


.PARAMETER Hashname
 
Specify System.Security.Cryptography hash to be used. MD5 is default.


.NOTES
    Version : 1.0  
    Author  : Håkon Jensen, Sicra AS  
       

#>


	Param(
		[Parameter(Mandatory=$true)][String]$String,
		[Parameter(Mandatory=$false)][String]$Hashname = "MD5"
	)


    $StringBuilder = New-Object System.Text.StringBuilder
    [System.Security.Cryptography.CryptoConfig]::CreateFromName($Hashname).ComputeHash([System.Text.Encoding]::UTF8.GetBytes($String))|%{
    [Void]$StringBuilder.Append($_.ToString("x2"))
    }
    return $StringBuilder.ToString()
}









function Write-Log{

<#

.SYNOPSIS

FUnction for Writing to a log file



.DESCRIPTION

This functions appends to a log file.
Path to log file is stored in global variable Config.Settings.LogFolder and defined in the Config.xml



.PARAMETER Text

The string to be written to the log file



.NOTES
    Version : 1.0  
    Author  : Håkon Jensen, Sicra AS  
       

#>
	Param(
		[Parameter(Mandatory=$true)][String]$Text
	)

	# Create LogFolder if it does not exist
	if(!(Test-Path -Path $global:Config.Settings.LogFolder )){
		New-Item -ItemType directory -Path $global:Config.Settings.LogFolder
	}

	$logFile = (Get-Date).ToString("yyyy_MM_dd") + "_Log.txt"
	$timestamp = (Get-Date).ToString("dd-MM-yyyy HH:mm:ss")
	$Text =  "$($timestamp)`t$($Text)"

	if ([Environment]::UserInteractive) {
		Write-Host $Text
	}

	$Text | Out-File -Append -Encoding utf8 -FilePath "$($global:Config.Settings.LogFolder)\$($logFile)"
}









function Load-Module{
<#

.SYNOPSIS

Helper function for loading configuration data when the module is loaded using Import-Module



.DESCRIPTION

This functions loads the information from Config.xml into a global varble used by all other functions.



.NOTES
    Version : 1.0  
    Author  : Håkon Jensen, Sicra AS  
       

#>
	# Read XML Config File  to global variable
	[xml]$global:Config =  Get-Content -Path "$($PSScriptRoot)\Config.xml"
	
	# Global variable for Confluence Server Template Page
	[string]$global:template


	# Load credentias from file that is inserted to global variable to be used by other functions
	Load-StoredWikiCredential -Account $global:Config.Settings.WikiAccount -ErrorAction SilentlyContinue
    Load-StoredVSphereCredential -Account $global:Config.Settings.VsphereAccount -ErrorAction SilentlyContinue
}


#Load configuration
Load-Module

Export-ModuleMember -Function Get-StringHash, Write-Log