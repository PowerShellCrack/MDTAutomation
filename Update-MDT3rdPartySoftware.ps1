<#
.SYNOPSIS
    Update 3rd party update files in MDT Apllication section
.DESCRIPTION
    Parses third party updates  CliXML list generated from Get-3rdPartySoftware.ps1. Then check MDT for applications similiar to the list using filters. 
.PARAMETER 
    NONE
.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -file "Update-MDT3rdPartySoftware.ps1"
.NOTES
    Script name: Update-MDT3rdPartySoftware
    Version:     1.0
    Author:      Richard Tracy
    DateCreated: 2018-11-02
#>

#==================================================
# FUNCTIONS
#==================================================

function Test-IsISE {
# try...catch accounts for:
# Set-StrictMode -Version latest
    try {    
        return $psISE -ne $null;
    }
    catch {
        return $false;
    }
}        
        
Function Import-SMSTSENV{
    ## Get the name of this function
    [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

    try{
        # Create an object to access the task sequence environment
        $Script:tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment 
        #$tsenv.GetVariables() | % { Write-Output "$ScriptName - $_ = $($tsenv.Value($_))" }
    }
    catch{
        Write-Output "${CmdletName} - TS environment not detected. Running in stand-alone mode."
    }
    Finally{
        #set global Logpath
        if ($tsenv){
            #grab the progress UI
            $Script:TSProgressUi = New-Object -ComObject Microsoft.SMS.TSProgressUI

            # Query the environment to get an existing variable
            # Set a variable for the task sequence log path
            #$Global:Logpath = $tsenv.Value("LogPath")
            $Global:Logpath = $tsenv.Value("_SMSTSLogPath")

            # Or, convert all of the variables currently in the environment to PowerShell variables
            $tsenv.GetVariables() | % { Set-Variable -Name "$_" -Value "$($tsenv.Value($_))" }
        }
        Else{
            $Global:Logpath = $env:TEMP
        }
    }
}

#==================================================
# FUNCTIONS
#==================================================
function Test-IsISE {
# try...catch accounts for:
# Set-StrictMode -Version latest
    try {    
        return $psISE -ne $null;
    }
    catch {
        return $false;
    }
}

Function Format-ElapsedTime($ts) {
    $elapsedTime = ""
    if ( $ts.Minutes -gt 0 ){$elapsedTime = [string]::Format( "{0:00} min. {1:00}.{2:00} sec.", $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );}
    else{$elapsedTime = [string]::Format( "{0:00}.{1:00} sec.", $ts.Seconds, $ts.Milliseconds / 10 );}
    if ($ts.Hours -eq 0 -and $ts.Minutes -eq 0 -and $ts.Seconds -eq 0){$elapsedTime = [string]::Format("{0:00} ms.", $ts.Milliseconds);}
    if ($ts.Milliseconds -eq 0){$elapsedTime = [string]::Format("{0} ms", $ts.TotalMilliseconds);}
    return $elapsedTime
}

Function Format-DatePrefix{
    [string]$LogTime = (Get-Date -Format 'HH:mm:ss.fff').ToString()
	[string]$LogDate = (Get-Date -Format 'MM-dd-yyyy').ToString()
    $CombinedDateTime = "$LogDate $LogTime"
    return ($LogDate + " " + $LogTime)
}

Function Write-LogEntry{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [Parameter(Mandatory=$false,Position=2)]
		[string]$Source = '',

        [parameter(Mandatory=$false)]
        [ValidateSet(0,1,2,3,4)]
        [int16]$Severity,

        [parameter(Mandatory=$false, HelpMessage="Name of the log file that the entry will written to.")]
        [ValidateNotNullOrEmpty()]
        [string]$OutputLogFile = $Global:LogFilePath,

        [parameter(Mandatory=$false)]
        [switch]$Outhost = $Global:OutToHost
    )
    ## Get the name of this function
    [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

    [string]$LogTime = (Get-Date -Format 'HH:mm:ss.fff').ToString()
	[string]$LogDate = (Get-Date -Format 'MM-dd-yyyy').ToString()
	[int32]$script:LogTimeZoneBias = [timezone]::CurrentTimeZone.GetUtcOffset([datetime]::Now).TotalMinutes
	[string]$LogTimePlusBias = $LogTime + $script:LogTimeZoneBias
    #  Get the file name of the source script

    Try {
	    If ($script:MyInvocation.Value.ScriptName) {
		    [string]$ScriptSource = Split-Path -Path $script:MyInvocation.Value.ScriptName -Leaf -ErrorAction 'Stop'
	    }
	    Else {
		    [string]$ScriptSource = Split-Path -Path $script:MyInvocation.MyCommand.Definition -Leaf -ErrorAction 'Stop'
	    }
    }
    Catch {
	    $ScriptSource = ''
    }
    
    
    If(!$Severity){$Severity = 1}
    $LogFormat = "<![LOG[$Message]LOG]!>" + "<time=`"$LogTimePlusBias`" " + "date=`"$LogDate`" " + "component=`"$ScriptSource`" " + "context=`"$([Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " + "type=`"$Severity`" " + "thread=`"$PID`" " + "file=`"$ScriptSource`">"
    
    # Add value to log file
    try {
        Out-File -InputObject $LogFormat -Append -NoClobber -Encoding Default -FilePath $OutputLogFile -ErrorAction Stop
    }
    catch {
        Write-Host ("[{0}] [{1}] :: Unable to append log entry to [{1}], error: {2}" -f $LogTimePlusBias,$ScriptSource,$OutputLogFile,$_.Exception.ErrorMessage) -ForegroundColor Red
    }
    If($Outhost){
        If($Source){
            $OutputMsg = ("[{0}] [{1}] :: {2}" -f $LogTimePlusBias,$Source,$Message)
        }
        Else{
            $OutputMsg = ("[{0}] [{1}] :: {2}" -f $LogTimePlusBias,$ScriptSource,$Message)
        }

        Switch($Severity){
            0       {Write-Host $OutputMsg -ForegroundColor Green}
            1       {Write-Host $OutputMsg -ForegroundColor Gray}
            2       {Write-Warning $OutputMsg}
            3       {Write-Host $OutputMsg -ForegroundColor Red}
            4       {If($Global:Verbose){Write-Verbose $OutputMsg}}
            default {Write-Host $OutputMsg}
        }
    }
}


##* ==============================
##* VARIABLES
##* ==============================

## Variables: Script Name and Script Paths
## Instead fo using $PSScriptRoot variable, use the custom InvocationInfo for ISE runs
If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation } Else { $InvocationInfo = $MyInvocation }
#Since running script within Powershell ISE doesn't have a $scriptpath...hardcode it
If(Test-IsISE){$scriptPath = "\\filer.s3i.org\s3isoftware\software\Scripts\PowerShell\MDTAutomation\Update-MDT3rdPartySoftware.ps1"}Else{$scriptPath = $InvocationInfo.MyCommand.Path}
[string]$scriptDirectory = Split-Path $scriptPath -Parent
[string]$scriptName = Split-Path $scriptPath -Leaf
[string]$scriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($scriptName)

#Get required folder and File paths
[string]$ConfigPath = Join-Path -Path $scriptDirectory -ChildPath 'Configs'

#Try to Import SMSTSEnv from MDT server
Import-SMSTSENV

[string]$MDTXMLFile = (Get-Content "$ConfigPath\mdt_configs.xml" -ReadCount 0) -replace '&','&amp;'
[xml]$MDTConfigs = $MDTXMLFile

#get the list of aoftware
[string]$3rdSoftwareRootPath = $MDTConfigs.mdtConfigs.softwareListCliXml.software.rootPath
[string]$3rdSoftwareListPath = $MDTConfigs.mdtConfigs.softwareListCliXml.software.listPath

#get the list of MDT servers to update
$MDTConfigs.mdtConfigs.server | Foreach {
    write-host ("Updating Software on DeploymentShare: " + $_.Host + "\" + $_.share + "...")
    [string]$MDTHost = $_.Host
    [string]$MDTShare = $_.share
    [string]$MDTPhysicalPath = $_.PhysicalPath

    [boolean]$RemoteMDTProvider = [boolean]::Parse($_.remoteMDTProvider)


    ##* ==============================
    ##* IMPORT MODULE/EXTENSIONS
    ##* ==============================
    #build mdt path to pull powershell
    $MDTSharePath = "\\" + $MDTHost + "\" + $MDTShare
    [string]$RemoteModulesPath = Join-Path -Path "$MDTSharePath\Tools" -ChildPath 'Modules'

    #import TaskSequence Module
    Import-Module $RemoteModulesPath\ZTIUtility -ErrorAction SilentlyContinue

    #Try to use remote mdt provider is enabled
    If($RemoteMDTProvider){
        [System.Management.Automation.PSCredential]$MDTCreds = Import-Clixml ($scriptRoot + "\" + $MDTConfigs.mdtConfigs.server.remoteAuthFile)

        Try{
            $Session = New-PSSession -ComputerName $MDTHost -Credential $MDTCreds
            Invoke-Command -Session $Session -Script {
                Import-Module "C:\Program Files\Microsoft Deployment Toolkit\bin\MicrosoftDeploymentToolkit.psd1"; 
                $Drive = New-PSDrive -Name DS001 -PSProvider mdtprovider -Root $args[0]
                cd $Drive.Root 
            } -Args ($MDTConfigs.mdtConfigs.server.PhysicalPath)
            #Enter-PSSession -Session $Session
            $MDTModuleLoaded = $true 
        }
        Catch{
            Write-Host "Failed to remote into: $($_.Exception.Message)" -ForegroundColor Red
            $MDTModuleLoaded = $false
        }
    }
    Else{
        $MDTModuleLoaded = $false
    }

    #load local mdt module if found and remote path did not work
    $LocalModulePath = "C:\Program Files\Microsoft Deployment Toolkit\bin\MicrosoftDeploymentToolkit.psd1"
    If(!$MDTModuleLoaded -and (Test-Path $LocalModulePath)){
    
        Import-Module $LocalModulePath
        #map to mdt drive (must have MDT module loaded)
        $MDTPSProvider = Get-PSProvider -PSProvider MDTProvider -ErrorAction SilentlyContinue
        If(!$MDTPSProvider){
            Try{
                $MDTDrive = New-PSDrive -Name DS001 -PSProvider mdtprovider -Root $MDTSharePath
            }
            Catch{
                Write-Host "Failed to load MDT module; Please try enabled remoteconfig" -ForegroundColor Red
            }
        }
    }
    Else{
        Write-Host "No MDT module were imported. Either try enabling remote import or go to [https://www.microsoft.com/en-us/download/details.aspx?id=54259]"
        break
    }


    ##* ==============================
    ##* MAIN
    ##* ==============================

    #Grab Variables from MDT's Control folder
    If(Test-Path "$MDTSharePath\Control\Settings.xml"){
        $MDTSettings = [Xml] (Get-Content "$MDTSharePath\Control\Settings.xml")
        [string]$MDT_Physical_Path = $MDTSettings.Settings.PhysicalPath
        [string]$MDT_UNC_Path = $MDTSettings.Settings.UNCPath


        $MDTAppGroupsFile = [Xml] (Get-Content "$MDTSharePath\Control\ApplicationGroups.xml")
        [xml]$MDTApps = Get-Content "$MDTSharePath\Control\Applications.xml" -Credential $MDTCreds

        $NewSoftwareList = Import-Clixml $3rdSoftwareListPath

        <#Test only
        $NewSoftware = ($NewSoftwareList | Where{($_.Product -match 'Chrome')} | Select -First 2)[1]
        $NewSoftware = $NewSoftwareList | Where{($_.Product -match 'Notepad\+\+')} | Select -First 1
        $NewSoftware = $NewSoftwareList | Where{($_.Product -match 'Reader DC')} | Select -first 1
        $NewSoftware = $NewSoftwareList | Where{($_.Product -match 'Reader DC')} | Select -last 1
        $NewSoftware = $NewSoftwareList | Where{($_.Product -match 'Flash Plugin')} | Select -last 1
        $NewSoftware = $NewSoftwareList | Where{($_.Product -match 'Java')} | Select -first 1
        $NewSoftware = $NewSoftwareList | Where{($_.Product -match 'Java')} | Select -last 1
        #>

        $UpdateMDTAppXML = 0
        $UpdatedAppCount = 0
        $ExistingAppCount = 0
        $MissingAppCount = 0

        foreach($NewSoftware in $NewSoftwareList)
        {

            If($NewSoftware.Arch){
                Write-Host ("Found [{0} {1} ({2}) - {3} bit] in software list" -f $NewSoftware.Publisher,$NewSoftware.Product,$NewSoftware.Version,$NewSoftware.Arch) -ForegroundColor Cyan
            }
            Else{
                Write-Host ("Found [{0} {1} ({2})] in software list" -f $NewSoftware.Publisher,$NewSoftware.Product,$NewSoftware.Version) -ForegroundColor Cyan
            }
            # clear the working app variable
            $MDTAppProducts = $null
            $MDTApp = $null

            #remove parent path of where the software was downloaded to. Attach rootPath from config
            $SplitSoftwarePath = $NewSoftware.FilePath -split "Software" 
            $NewUNCPath = $3rdSoftwareRootPath + '\Software' + $SplitSoftwarePath[-1]
            # find an MDT app that matches the software list base on Publisher, Product Name and Product Type (not always specified)
            $MDTAppProducts = $MDTApps.applications.application | Where{($_.Publisher -eq $NewSoftware.Publisher) -and ($_.Name -match [regex]::Escape($NewSoftware.Product))}
    
            #if more than 1 are found, filter on product type to reduce it
            If($MDTAppProducts.Count -ge 2){
                $MDTAppFilter1 = $MDTAppProducts | Where {($_.Name -match [regex]::Escape($NewSoftware.ProductType))}
                If($MDTAppFilter1){$MDTAppProducts = $MDTAppFilter1}
            }
    
            #if more than 1 are found, filter on arch match, if specified (names labeled with x64 or x86) to reduce it. 
            If($MDTAppProducts.Count -ge 2){
                $MDTAppFilter2 = $MDTAppProducts | Where {($_.Name -match $NewSoftware.Arch) -or ($_.ShortName -match $NewSoftware.Arch)}
                If($MDTAppFilter2){$MDTAppProducts = $MDTAppFilter2}
            }

            #if more than 1 are found, filter on arch no match, if NOT specified (usually labeled with x86) to reduce it.
            If($MDTAppProducts.Count -ge 2){
                $MDTAppFilter3 = $MDTAppProducts | Where {($_.Name -notmatch 'x64') -and ($_.ShortName -notmatch 'x64')}
                If($MDTAppFilter3){$MDTAppProducts = $MDTAppFilter3}
            }    

            $MDTApp = $MDTAppProducts | Select -First 1

            #If and app is found
            If($MDTApp){
                Write-Host ("Filtered Application in MDT to [{0}]" -f $MDTApp.Name) -ForegroundColor DarkYellow

                #remove share from path to get relative path
                $mappedPath = ($MDTApp.WorkingDirectory).Replace('.',$MDTSharePath)

                #' Get current folders
                #' ===================================
                ## Parse working directory and drill only two folders deep.
                ## Anything else deeper doesn't matter because root folder will be deleted if needed
                $CurrentFolders = Get-ChildItem -Path $mappedPath -Recurse -Depth 1 -Force | ?{ $_.PSIsContainer }

            
                $VersionFolderFound = $false
                $SourceFolderFound = $false
                $ConfigFolderFound = $false
                $CurrentVersion = $null
                $ExtraFolders = @()
                $IgnoreFolders = @()
                $KeepFolders = @()

                ## if mutiple folders exist, loop through them to see if its a version folder name.
                ## anything other than the identified folders will be deleted later on
                $CurrentFolders | Foreach-Object {
            
                    switch($_.Name) { 
                        "Source"            {
                                                $SourceFolderFound = $true
                                                Write-Host "Found [Source] folder; " -NoNewline
                                                $KeepFolders += $_
                                            }
                    
                        "Configs"           {
                                                $ConfigFolderFound = $true
                                                Write-Host "Found [Config] folder; " -NoNewline
                                                $IgnoreFolders += $_
                                            }
                    
                        "Updates"           {
                                                $UpdatesFolderFound = $true
                                                Write-Host "Found [Updates] folder; " -NoNewline
                                                $KeepFolders += $_
                                            }

                        "$($NewSoftware.Version)" {
                                                $NewVersion = $NewSoftware.Version
                                                Write-Host "Found [$NewVersion] folder; " -NoNewline
                                                $VersionFolderFound = $true
                                                $KeepFolders += $_
                                            }

                        default             {   $CurrentVersion = $MDTApp.Version
                                                Write-Host "Found [$CurrentVersion] folder; " -NoNewline
                                                $VersionFolderFound = $true
                                                $ExtraFolders += $_
                                            }
                    }
                    write-host ("Categorizing folder [" + $_.Name + "]")
                }


                #ensure folders are different and deliminated for matching
                $ExtraFolders = ($ExtraFolders | Select -Unique) -join "|"

                #ensure folders are different and deliminated for matching
                $IgnoreFolders = ($IgnoreFolders | Select -Unique) -join "|"

                #Always add the new verison as a keeper
                $KeepFolders += $NewSoftware.Version
                $KeepFolders = ($KeepFolders | Select -Unique) -join "|"

                #' Build folder paths 
                #' ===================================
                ## Does current directory has version folder and source folder
                ## mimic same path structure with new version
                If($sourceFolderFound){$subpath = '\Source\'}Else{$subpath = '\'}
                If($versionFolderFound){$leafpath = $NewSoftware.Version}Else{$leafpath = ''}
                $DestinationPath = ($mappedPath + $subpath + $leafpath)

                #' Compare the versions
                #' ===================================
                # and check to see if file exists
        
                # assume if version match that previous similiar software updated the application. 
                # This issue exists when multiple architecture version exists
                If($MDTApp.Version -eq $NewSoftware.Version){
                    Write-Host ("Application [{0}] version [{1}] was already found in MDT, checking if file exists..." -f $MDTApp.Name,$MDTApp.Version) -ForegroundColor Gray
                    If(-not(Test-Path "$DestinationPath\$($NewSoftware.File)")){

                        ##' If the copy fails return to the stop processing the new software
                        Try{
                            Copy-Item $NewUNCPath -Destination $DestinationPath -Force | Out-null
                            Write-Host ("Copied File [{0}] to [{1}]" -f $NewSoftware.File,$DestinationPath) -ForegroundColor Green
                        }
                        Catch{
                            Write-Host ("Failed to copy File [{0}] to [{1}]" -f $NewSoftware.File,$DestinationPath) -ForegroundColor Red
                            Return
                        }
                    }
                    Else{
                        Write-Host ("Application [{0}] version [{1}] was already found in MDT" -f $MDTApp.Name,$MDTApp.Version) -ForegroundColor Green
                        $ExistingAppCount ++
                    }
                }

                # Update the version
                Else{
            
                    #' Copy new application files
                    #' ================================
                    ##' Do this before deleting the old files just in case. 
                    ##' If the copy fails return to the stop processing the new software
                    New-Item $DestinationPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
                    Try{
                        Copy-Item $NewUNCPath -Destination $DestinationPath -Force | Out-null
                        Write-Host ("Copied File [{0}] to [{1}]" -f $NewSoftware.File,$DestinationPath) -ForegroundColor Green
                    }
                    Catch{
                        Write-Host ("Failed to copy File [{0}] to [{1}]" -f $NewSoftware.File,$DestinationPath) -ForegroundColor Red
                        Return
                    }
            

                    #' Update Script Installer
                    #' =========================
                    $Command = ($MDTApp.CommandLine).split(" ")
                    $CommandUpdated = $false

                    switch($Command[0]){
                    #second update the installer scripts
                        'cscript' { 
                                    Write-Host ("Found a cscript [{0}] for the installer" -f $Command[1]) -ForegroundColor Gray
                                    #grab content from script that installs application
                                    $content = Get-Content "$($mappedPath + '\' + $Command[1])" | Out-String
                                    #find text line that has sVersion
                                    $pattern = 'sVersion\s*=\s*(\"[\w.]+\")'
                                    $content -match $pattern | Out-Null
                                    # if found in cscript installer, update it and save it
                                    If($matches){
                                        $NewContentVer = $content.Replace($matches[1],'"' + $NewSoftware.Version + '"')
        
                                        #add updated version to vbscript
                                        $NewContentVer | Set-Content -Path "$($mappedPath + '\' + $Command[1])" 
                                        Write-Host ("Updated [{0}] variable [sVersion] from [{1}] to [{2}]" -f $Command[1],$matches[1].replace('"',''),$NewSoftware.Version) -ForegroundColor DarkYellow
                                        $CommandUpdated = $true
                                    }
                                    Else{
                                        Write-Host ("Unable to find [sVersion] variable in [{0}], there may be an issue during deployment" -f $Command[1]) -ForegroundColor Red
                                    }

                                    #Clear matches
                                    $matches = $null       
                                }


                        '*.exe' {
                                    Write-Host ("Found a executable [{0}] for the installer" -f $Command[1]) -ForegroundColor Gray
                                }

                        'Powershell*' {
                                    Write-Host ("Found a powershell script [{0}] for the installer" -f $Command[1]) -ForegroundColor Gray
                                }

                        'msiexec*' {
                                    Write-Host ("Found a msi file [{0}] for the installer" -f $Command[1]) -ForegroundColor Gray
                                }
                    }


                    If($CommandUpdated){

                        #' Remove old application files
                        #' =================================
                
                        ## Delete any extra folders found (not keep folders and ignore anything in the with a fullpath of ignored folders)
                        Get-ChildItem -Path $mappedPath -Recurse -Depth 1 -Force -Directory | ?{ $_.Name -match $ExtraFolders -and $_.name -notmatch $KeepFolders -and $_.Fullname -notmatch $IgnoreFolders} | ForEach-Object{
                            #if mutiple files exist, loop through them to see if its a version name.
                            Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
                            Write-Host ("Deleted Folder: {0}" -f $_.FullName) -ForegroundColor Red
                        }

                        #' Update MDT Listing
                        #' =========================
                        $MDTApp.Version = $NewSoftware.Version
                        Write-Host ("Configured to change MDT's Application [{0}] version property to [{1}]" -f $MDTApp.Name,$NewSoftware.Version) -ForegroundColor DarkGreen
                        $UpdateMDTAppXML ++

                    }
                }

                #' Save MDT Listing
                #' =========================
                Try{
                    If($UpdateMDTAppXML -gt 0){$mdtapps.save("$MDTSharePath\Control\Applications.xml")}
                    Write-Host ("Saved changes to MDT's Application configuration file [{0}] for [{1}]" -f "$MDTSharePath\Control\Applications.xml",$MDTApp.Name) -ForegroundColor Green
                    #reset back to 0
                    $UpdateMDTAppXML = 0
                }
                Catch{
                    Write-Host ("Failed write changes to MDT's Application configuration file [{0}] for [{1}]" -f "$MDTSharePath\Control\Applications.xml",$MDTApp.Name) -ForegroundColor Red
                    Return
                }
            }
            Else{
                Write-Host ("Application [{0} {1} ({2})] was not found in MDT" -f $NewSoftware.Publisher,$NewSoftware.Product,$NewSoftware.Arch) -ForegroundColor Yellow
                $MissingAppCount ++
            }

    
        } 

        Write-host ("Updated " + $UpdatedAppCount + " Applications in MDT")
        Write-host ("Found " + $MissingAppCount + " missing Applications in MDT")
        Write-host ("Existing " + $ExistingAppCount + " Applications already up-to-date")
    }
    Else{
        Write-Host ("Failed write get to MDT's Settings from [{0}]" -f "$MDTSharePath\Control\Settings.xml") -ForegroundColor Red
    }
}