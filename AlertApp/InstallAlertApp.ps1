
$ModuleName = "SharePointPnPPowerShellOnline"
$sPSite = "https://racwanpe.sharepoint.com"
$SPAdminSite = "https://racwanpe-admin.sharepoint.com"
$LogFile = "C:\Temp\RAC_SPFxInstall_Log.txt"
$csvFile = "C:\temp\AssociatedSites.csv"


function LogMessage($message) 
{
    $DateNow = Get-Date -Format "dd/mm/yyyy HH:mm:ss" 
    $LogTxt = $DateNow.toString() + " - " + $message
    $LogTxt | Out-File $LogFile -append
}

function Uninstall-SPO()
{
   
  #Get-InstalledModule -Name $ModuleName -RequiredVersion 3.14.19101 | Uninstall-Module
  Write-Host "Uninstalling previous version of $ModuleName module was successful"
}

function Install-SPO()
{
 
  Install-Module -Name "SharePointPnPPowerShellOnline" -RequiredVersion 3.23.2007.1
  Write-Host "Installation the previous version of $ModuleName was successful"

}

function GetSPOPnPVersion($viewMode)
{
    #Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable | Select Name,Version
    
    if ($viewMode -eq 1)
    {
        Write-Host "Get list of all availiable $ModuleName version"
        Get-Module SharePointPnPPowerShell* -ListAvailable | Select-Object Name,Version | Sort-Object Version -Descending
        #Get-Module -Name $ModuleName -ListAvailable | Select Name,Version 
    }
    else {
        Write-Host "Dislay which $ModuleName version was installed"
        Get-InstalledModule -Name "SharePointPnPPowerShellOnline"
    }
}

function PreRequirement()
{
    Connect-PnPOnline -Url $sPSite -PnPManagementShell 
}

function ConnectSPO()
{
    #Register-PnPManagementShellAccess

    Write-Host "Connecting to $SPSite ..."
    Connect-PnPOnline -Url $sPSite
    
}

function InstallSPFX()
{
    ConnectSPO
    Apply-PnPTenantTemplate -Path starterkit.pnp #starterkit-spfx-only.pnp
}


#This function will export all associated hub site for intranet hub into CSV
function GetAssociatedSites($Hub)
{
    Connect-PnPOnline -Url $SPAdminSite -UseWebLogin
    $coll = [System.Collections.ArrayList]@()

    Write-Host "Hub site: $Hub"

     Get-PnPHubSiteChild -Identity $Hub | % {
       
        Write-Host -ForegroundColor Yellow "Assoicated hub site: $($_) "

        $obj = New-Object PSObject
        $obj | Add-Member -MemberType NoteProperty "Url"  -Value $_
        $obj
        $coll.Add($obj) | Out-Null

        $coll  | Export-CSV $csvFile –NoTypeInformation      

    }
    
    $logMsg = "Assoicated hub sites has been exported to $csvFile"
    Write-Host $logMsg
    LogMessage $logMsg
    Disconnect-PnPOnline
}

function ActivateAlertAppToSites()
{
    
    try
    {
        ActivateAlertApp $sPSite #Activate alert app on hub site

        #Activate alert app for all associated hub sites
        $table = Import-Csv $csvFile -Delimiter ","
        foreach($row in $table)
        {                 
            ActivateAlertApp $row.Url              
        }

        Disconnect-PnPOnline


    }
     catch {

		$errMsg = "Error at ActivateAlertAppToSites " + $Error[0].Exception.Message 
		Write-Host $errMsg -ForegroundColor Red
		LogMessage $errMsg      

	}#end try
}

function ActivateAlertApp($siteURL){
     
    try
    {
         Write-Host "Connecting to hub site: " $siteURL -foreground yellow
         Connect-PnPOnline -Url $siteURL #-UseWebLogin
         Add-PnPCustomAction -Name "HubOrSiteAlertsApplicationCustomizer" -Title "HubOrSiteAlertsApplicationCustomizer" -ClientSideComponentId 29df5d8b-1d9b-4d32-971c-d66162396ed3 -Location "ClientSideExtension.ApplicationCustomizer" -ClientSideComponentProperties "{}" -Scope Site

         $logMsg = "Alert app has been activated for assoicated hub site $siteURL"
         Write-Host $logMsg
         LogMessage $logMsg

    }
     catch {

		$errMsg = "Error at ActivateAlertApp " + $Error[0].Exception.Message 
		Write-Host $errMsg -ForegroundColor Red
		LogMessage $errMsg      

	}#end try

}

#------ MAIN -------
#GetSPOPnPVersion 1

GetAssociatedSites $sPSite
ActivateAlertAppToSites