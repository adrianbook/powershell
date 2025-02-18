
Import-Module PnP.PowerShell
Import-Module C:\UtilitarianScripturesAndFunctins\powershell\logging\LoggingUtility.ps1

function Create-SharePointFolderStructure {
    param(
        [string]$PathLocal,
        [string]$PathSP
    )
    
    $localDirectory = Get-Item $PathLocal
    if ($localDirectory.getType().Name -ne 'DirectoryInfo'){
        Write-Error "Found '$($localDirectory.GetType().Name)' expected 'DirectoryInfo'"
        return
    }

    $spFolderPath = Join-Path $PathSP $localDirectory.Name

    try {
        $spFolder = Resolve-PnPFolder -SiteRelativePath $spFolderPath
        Write-Host "Added folder $spFolderPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Error writing directory $spFolderPath $_"
        return
    }

    UploadFilesToSPFolder -Files (dir $localDirectory -File) -SPFolder $spFolder

    $subDirectories = dir $localDirectory -Directory

    if ($subDirectories -ne $null){
        foreach ($subDirectory in $subDirectories){
            Create-SharePointFolderStructure -PathLocal $subDirectory.FullName -PathSP $spFolderPath
        }
    }
}

function UploadFilesToSPFolder {
    param(
        [Microsoft.SharePoint.Client.Folder]$SPFolder,
        [System.IO.FileInfo[]]$Files
    )

    $sizeThreshold = 10 * 1024 * 1024

    foreach ($file in $Files){
        if ($file.length -le $sizeThreshold){
            UploadFile -File $file -SPFolder $SPFolder
        }
        else {
            UploadFileWithStream -File $file -SPFolder $SPFolder
        }
    }
}





function UploadFile {
    param(
        [System.IO.FileInfo]$File,
        [Microsoft.SharePoint.Client.Folder]$SPFolder    
    )

    $fileName = ResolveSPFileName $File.Name

    try {
        Add-PnPFile -Path $File.FullName -Folder $SPFolder -NewFileName $fileName > $null
        Write-Host "    Added file $($SPFolder.ServerRelativePath)/$fileName" -ForegroundColor Green
    }
    catch {
        Write-Host "    Failed upload file: $($File.FullName) to: $($SPFolder.ServerRelativePath)" -ForegroundColor Red
    }
}


function UploadFileWithStream {
    param(
        [System.IO.FileInfo]$File,
        [Microsoft.SharePoint.Client.Folder]$SPFolder    
    )

    $fileName = ResolveSPFileName $File.Name
    $stream = [System.IO.File]::OpenRead($File.FullName)

    try {
        Add-PnPFile -Folder $SPFolder -FileName $fileName -Stream $stream > $null
        Write-Host "    Added file $($SPFolder.ServerRelativePath)/$fileName"
    }
    catch {
        Write-Host "    Failed upload file with stream: $($File.FullName) to: $($SPFolder.ServerRelativePath)" -ForegroundColor Red
    }
    finally {
        $stream.Close()
    }
}



function ResolveSPFileName {
    param([string]$LocalFileName)

    $splitName = $LocalFileName.Split(".")
    if ($splitName.Count -gt 1) {
        $suffix = $splitName[$splitName.Count - 1]
        $name = (($splitName | Select-Object -SkipLast 1 | %{ "$_." }) -join "").TrimEnd(".").Trim()
    }
    else {
        $name = $splitName
    }

    $name = $name.Replace('"',"'")
    $name = $name.Replace('*',"#")
    $name = $name.Replace(':',";")
    $name = $name.Replace('<',"(")
    $name = $name.Replace('>',")")
    $name = $name.Replace('?',".")
    $name = $name.Replace('/',"%")
    $name = $name.Replace('\',"&")
    $name = $name.Replace('|',"I")


    if ($suffix -eq $null){
        return $name
    }
    else {
        return "$name.$suffix"
    }
}


function CreateSpFolderIfNotPresent {    
    [OutputType([Microsoft.SharePoint.Client.Folder])]
    param(
	    [string]$BasePath,
        [string]$FolderName
    )

    $spFolderPath = Join-Path $BasePath $FolderName

    return Resolve-PnPFolder -SiteRelativePath $spFolderPath
}

function Connect-To-SharePoint-And-Run {
        param(
        [Parameter(Mandatory)]
        [string]$SiteUrl,
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock
    )
    
    try {
        Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $env:PnPPowershellStenaClientId
        & $ScriptBlock
    }
    finally {
        Disconnect-PnPOnline
    }
}

function ResolveSpUrl {
    param([string]$Url)

    $matches = $Url.Trim() | select-string "(https://[^\.]+\.sharepoint.com/sites/[^/]+)/(.+)$"

    if ($matches -eq $null) {
        Write-Error "$Url is not an url to a sharepoint site"
        return
    }
    if ($matches.Matches.Count > 1) {
        Write-Error "Can only handle one url at a time"
        return
    }

    return @{
        SiteUrl = $matches.Matches[0].Groups[1].Value
        SpFolderSiteRelativePath = $matches.Matches[0].Groups[2].Value
    }

}

function RecursivelyCopyDirectoryToSP {
    param(
        [Parameter(Mandatory=$false)]
        [string]$LocalPath,
        [Parameter(Mandatory=$false)]
        [string]$SPFolderUrl,
        [Parameter(Mandatory=$false)]
        [string]$LogFilePath,
        [Parameter(Mandatory=$false)]
        [string]$LogLabel
    )

    if ($LocalPath -eq $null -or $LocalPath -eq "") {
        Write-Error "Parameter -LocalPath must be provided"
        return
    }

    if ((Test-Path $LocalPath -PathType Container) -eq $false) {
        Write-Error "$LocalPath is not resolvable to a directory"
        return
    }

    $spSiteUrlAndFolderPath = ResolveSpUrl $SPFolderUrl

    if ($spSiteUrlAndFolderPath -eq $null) {
        return
    }

    $spSiteUrl = $spSiteUrlAndFolderPath.SiteUrl
    $spFolderPath = $spSiteUrlAndFolderPath.SpFolderSiteRelativePath

    Connect-To-SharePoint-And-Run -SiteUrl $spSiteUrl -ScriptBlock {
        $spFolder = Resolve-PnPFolder -SiteRelativePath $spFolderPath
        if ($spFolder -eq $null) {
            Write-Error "$spFolderPath is not a folder"
            return
        }
        
        $label = if ($LogLabel -ne $null -and $LogLabel -ne "") {$LogLabel} else { "CopyDirectoryToSP $LocalPath $SPFolderUrl" }
                
        WriteLogsToFile -LogLabel $label -LogDirectoryPath $LogFilePath -Script {
            Create-SharePointFolderStructure -PathLocal $LocalPath -PathSP $spFolderPath
        }
    }

}