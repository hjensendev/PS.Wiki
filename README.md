# Wiki.PS
A project for creating Confluence pages with information from vSphere.

# Getting Started
1. Download this repository to any computer with Powershell 4.0 or higher. The computer must be able to reach both your Confluence and vSphere server.

2. Install requirements **VMware.PowerCLI** and **ConfluencePS**

3. Create a the following pages in Confluence
* A server template page
* A placeholder parent page

4. Create your custom attributes and tags in vSphere. You can validate your vSphere configuration by running *Get-VsphereServerData*.

5.  Configure *config.xml* for your environment

6.  Run .INSTALL.PS1

7.  Start the script

### Server template page
Create a template page with your desired layout. I recommend using a Page Property Report Macro for storing key information.

![Image of ServerTemplate](https://sicra.no/wp-content/uploads/2018/04/psWikiServerTemplate.png)

Update the *New-ConfluenceServerBody*-function in **Confluence.psm1** to match your template.
```
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
```

### Parent page
Create a parent page placeholder for where your server pages will be created.

![Image of ServerTemplate](https://sicra.no/wp-content/uploads/2018/04/psWikiParentPage.png)

### Configuration
Configure *config.xml* for your environment, before installing the module

Property | Description | Example Vaule
------------ | ------------- | ------------- 
LogFolder | Where log files and credentials are stored | C:\Operation\PS.Wiki
vSphereServer | FQDN of your vSphere Server | srv-vc01.sicralabs.local
vSphereAccount | An account with read and write permissions to VM servers in vSphere | SICRALABS\VMUSER or vmuser@sicralabs.local
WikiAccount | An account with read and write permissions to the Confluence Space where your server pages will be created | SICRALABS\WIKIUSER or wikiuser@sicralabs.local
ConfluenceServerBaseURI | URL to your Confluence server | https://wiki.sicralabs.local:8090
ConfluenceServerTemplatePageId | The Page ID of the Server Template Page | 8364821
ConfluenceServerParentPageId | The Page ID of the parent page where server wiki pages will be created | 8364806
ConfluenceSpaceKey | The shortname (key) where ConfluenceServerParentPageId is created | IT

### Installation
Open a powershell console with RunAs Administrator and run `.\INSTALL.PS1`

### Starting the script
Open a powershell console and type `Start-WikiAutomation`

#### Arguments
`Start-WikiAutomation [-Filter ServerName]`

# Requirements
## VMware.PowerCLI
Can be downloaded from Powershell Gallery using Install-Module

[Link](http://vmware.com/go/powercli)


## ConfluencePS 
Can be downloaded from Powershell Gallery using Install-Module

[Link](https://github.com/AtlassianPS/ConfluencePS)
