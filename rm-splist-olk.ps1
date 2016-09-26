<#
	Removes the specified Sharepoint Lists from Outlook.
	Uses Outlook MAPI and must run on each user's workstation. Deploy as a script with GPO or System Center
	Author: Quinten Steenhuis, 9/26/2016
#>

<#
# Uncomment this section and update with your own lists to remove. If left
# unspecified, this script will rely on the default list set via GPO for SP 2013
# Please note you'll need to update the GPO before next refresh cycle if relying on
# that default selection, so specifying the list below is recommended.
$defaultLists = @(
	"List 1",
	"List 2"
	)
#>

# Check the registry for the Adminstrative Template that set default SP lists
function get-defaultlists {
	$DefaultListPath = "hkcu:\Software\Policies\Microsoft\Office\15.0\outlook\options\"
	if (test-path $DefaultListPath) {
		$a = get-childitem "hkcu:\Software\Policies\Microsoft\Office\15.0\outlook\options\"
		return $a.Property
	} else {
		return @()
	}
}

$listsToRemove = if ($defaultLists) {$defaultLists} else {get-defaultlists}

$mapped = $listsToRemove | %{join-path "\\SharePoint Lists" $_}

$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$sp = $ns.Folders | Where-Object {$_.FolderPath -like "\\Sharepoint Lists"}

$spFolders = $sp.folders

foreach($folder in $spFolders) {
	if ($mapped -contains $folder.folderpath){
		write-host "Deleting SP list" $folder.folderpath
		$folder.Delete()
	}
}

