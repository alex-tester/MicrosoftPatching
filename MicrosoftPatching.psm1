<#
.SYNOPSIS
Downloads and Installs all available Microsoft Patches
.DESCRIPTION
Uses Microsoft.Update COM Object methods to search for available updates, download, and install them.
.EXAMPLE
Install-MicrosoftPatches
#>
function Install-MicrosoftPatches
{

    $NoPatches = $false

    #Define update criteria
    $Criteria = "IsInstalled=0 and Type='Software'"

    #Search for relevant updates
    $Searcher = New-Object -ComObject Microsoft.Update.Searcher

    $SearchResult = $Searcher.Search($Criteria).Updates

    #If no patches are returned, initiate a check in, then try again after 5 minutes
    if($SearchResult.count -eq 0)
    {
        wuauclt /reportnow
        start-sleep -Seconds 300

        #Search

        $Searcher = New-Object -ComObject Microsoft.Update.Searcher

        $SearchResult = $Searcher.Search($Criteria).Updates 
    }


    if($SearchResult.count -gt 0)
    {
       
        foreach($item in $SearchResult)
        {
            #log the patches to be installed
            Write-Output "Preparing to install $item.Title"
        }

    }
    else
    {
        $NoPatches = $true
    }


    if($NoPatches -eq $false)
    {

        #Download updates.
        $Session = New-Object -ComObject Microsoft.Update.Session

        $Downloader = $Session.CreateUpdateDownloader()

        $Downloader.Updates = $SearchResult

        $DownloadResult = $Downloader.Download()


        #Install updates

        $Installer = New-Object -ComObject Microsoft.Update.Installer

        $Installer.Updates = $SearchResult

        $InstallResult = $Installer.Install()


        if($InstallResult.HResult -eq 0)
        {
            Write-Output "Patches installed successfully"
        }
        else
        {
            Write-Output "Patches failed to install with error code $InstallResult.HResult"
        }
        
    }
    else
    {
    Write-Output "No updates available for installation"
    }

return $installresult
}