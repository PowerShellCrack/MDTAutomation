##*===========================================================================
##* FUNCTIONS
##*===========================================================================

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
        [switch]$Outhost
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

function Show-ProgressStatus
{
    <#
    .SYNOPSIS
        Shows task sequence secondary progress of a specific step
    
    .DESCRIPTION
        Adds a second progress bar to the existing Task Sequence Progress UI.
        This progress bar can be updated to allow for a real-time progress of
        a specific task sequence sub-step.
        The Step and Max Step parameters are calculated when passed. This allows
        you to have a "max steps" of 400, and update the step parameter. 100%
        would be achieved when step is 400 and max step is 400. The percentages
        are calculated behind the scenes by the Com Object.
    
    .PARAMETER Message
        The message to display the progress
    .PARAMETER Step
        Integer indicating current step
    .PARAMETER MaxStep
        Integer indicating 100%. A number other than 100 can be used.
    .INPUTS
         - Message: String
         - Step: Long
         - MaxStep: Long
    .OUTPUTS
        None
    .EXAMPLE
        Set's "Custom Step 1" at 30 percent complete
        Show-ProgressStatus -Message "Running Custom Step 1" -Step 100 -MaxStep 300
    
    .EXAMPLE
        Set's "Custom Step 1" at 50 percent complete
        Show-ProgressStatus -Message "Running Custom Step 1" -Step 150 -MaxStep 300
    .EXAMPLE
        Set's "Custom Step 1" at 100 percent complete
        Show-ProgressStatus -Message "Running Custom Step 1" -Step 300 -MaxStep 300
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string] $Message,
        [Parameter(Mandatory=$true)]
        [int]$Step,
        [Parameter(Mandatory=$true)]
        [int]$MaxStep,
        [string]$SubMessage,
        [int]$IncrementSteps,
        [switch]$Outhost
    )

    Begin{

        If($SubMessage){
            $StatusMessage = ("{0} [{1}]" -f $Message,$SubMessage)
        }
        Else{
            $StatusMessage = $Message

        }
    }
    Process
    {
        If($Script:tsenv){
            $Script:TSProgressUi.ShowActionProgress(`
                $Script:tsenv.Value("_SMSTSOrgName"),`
                $Script:tsenv.Value("_SMSTSPackageName"),`
                $Script:tsenv.Value("_SMSTSCustomProgressDialogMessage"),`
                $Script:tsenv.Value("_SMSTSCurrentActionName"),`
                [Convert]::ToUInt32($Script:tsenv.Value("_SMSTSNextInstructionPointer")),`
                [Convert]::ToUInt32($Script:tsenv.Value("_SMSTSInstructionTableSize")),`
                $StatusMessage,`
                $Step,`
                $Maxstep)
        }
        Else{
            Write-Progress -Activity "$Message ($Step of $Maxstep)" -Status $StatusMessage -PercentComplete (($Step / $Maxstep) * 100) -id 1
        }
    }
    End{
        Write-LogEntry $Message -Severity 1 -Outhost:$Outhost
    }
}


##*===========================================================================
##* VARIABLES
##*===========================================================================
## Instead fo using $PSScriptRoot variable, use the custom InvocationInfo for ISE runs
If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation } Else { $InvocationInfo = $MyInvocation }
[string]$scriptDirectory = Split-Path $MyInvocation.MyCommand.Path -Parent
[string]$scriptName = Split-Path $MyInvocation.MyCommand.Path -Leaf
[string]$scriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($scriptName)

#un-comment $scriptDirectory for ISE testing
#$scriptDirectory = 'D:\DeploymentShare\Automate'

#build local paths
[string]$ConfigPath = Join-Path -Path $scriptDirectory -ChildPath 'Configs'


#Try to Import SMSTSEnv
Import-SMSTSENV

#Get required folder and File paths
$MDTSharePath = Split-Path $scriptDirectory -Parent

#Grab Variables from MDT's Control folder
$MDTSettings = [Xml] (Get-Content "$MDTSharePath\Control\Settings.xml")
[string]$MDT_Physical_Path  = $MDTSettings.Settings.PhysicalPath
[string]$MDT_UNC_Path    = $MDTSettings.Settings.UNCPath
[string]$MDT_Server    = $MDTSettings.Settings.MonitorHost
[string]$ModulesPath = Join-Path -Path "$MDTSharePath\Tools" -ChildPath 'Modules'
$MDTDrive = New-PSDrive -Name DS001 -PSProvider mdtprovider -Root $MDTSharePath

#remote into MDT server to import MDT module
Enter-PSSession $MDT_Server 

#import TaskSequence Module
Import-Module $ModulesPath\ZTIUtility -ErrorAction SilentlyContinue
Import-Module "C:\Program Files\Microsoft Deployment Toolkit\bin\MicrosoftDeploymentToolkit.psd1"


$PatchedMDTOSFolder = 'Patched'
$CurrentMDTOSFolder = 'Production'
$LatestMDTOSFolder = 'Latest'


#Start Transcript Logging
Start-Transcript -path $LogFile -Force


# grab current deployments
#Get-MDTMonitorData -Path "$($MDTDrive.Name):"

#Grabbing MDT folders under Operating System to determine import path for WIM from patched version
$MDTOSGroupsFile = [Xml] (Get-Content "$MDTSharePath\Control\OperatingSystemGroups.xml")
#find Patched folder
$MDT_Dest_Folder = ($MDTOSGroupsFile.groups.group | Where{$_.Name -match $PatchedMDTOSFolder}).Name | Select -First 1
#create directory if it doesn't exist in MDT
If(!$MDT_Dest_Folder){New-Item -Path "$($MDTDrive.Name):\Operating Systems\$PatchedMDTOSFolder" -ItemType Directory -Force | Out-Null}


If($Global:tsenv){
    $MDT_Working_Dir = $tsenv.Value("BackupDir")
    $BackupFile = $tsenv.Value("BackupFile")
}
Else{
    $MDT_Working_Dir = "Captures"
    $BackupFile = "WIN101803X64OFF16_2018-10-18_1455.wim"
    $BackupFile = "WIN101803X64_2018-10-18_1500.wim"
}

#Build both Current name and latest build names.
#pattern grabs the string before and after an 8 dights string (eg. WIN101803X64_20181018.wim)
#$pattern = "(?<text>.*)(?<date>\d{8})(?<text2>.*)"
#$BackupFile -match $pattern | Out-Null
#if match is found, trim the date off
#$WIMTrimDate = $BackupFile -replace ('_' + $matches.date),""
#$WIMPatchedBaseName = [io.path]::GetFileNameWithoutExtension($WIMTrimDate)
#Drop everything after the first _ to ger base name 
#remove patch to get true base name
#$WIMName = ($WIMTrimDate -replace ('_' + $PatchedMDTOSFolder),"")
#$WIMBaseName = [io.path]::GetFileNameWithoutExtension($WIMName)

#build captured wim using variables in running tasksequence
$WIMPatchedBaseName = $BackupFile.Split("_")[0]
$WIMTrimDate = $BackupFile.Split("_")[1]
$Build_Captures_Path = "" + $MDT_Physical_Path + "\" + $MDT_Working_Dir + "\" + $BackupFile + ""
$SourceWIM = Get-ChildItem $Build_Captures_Path
$MDTPath = "$($MDTDrive.Name):\Operating Systems\$MDT_Dest_Folder"
$BuildDestFolder = $PatchedMDTOSFolder + '_' + $WIMPatchedBaseName + '_' + $WIMTrimDate

# move captured WIM into MDT operating systems to default Patched branch
$newCapture = Import-MDTOperatingSystem -Path $MDTPath -SourceFile $SourceWIM.FullName -DestinationFolder $BuildDestFolder -Move


# Grab updated Operating System content and group (folder) membership
$MDTOSGroupsFile = [Xml] (Get-Content "$MDTSharePath\Control\OperatingSystemGroups.xml")
$MDTOSsFile = [Xml](Get-Content "$MDTSharePath\Control\OperatingSystems.xml")

# grab GUID and Build of uploaded WIM
$NewMDTOS = $MDTOSsFile.oss.os | Where {($_.Name -match "$($newCapture.Name)")}
[string]$NewMDTOSBuild = $NewMDTOS.Build
[string]$NewMDTOSGUID = $NewMDTOS.guid

# Since there could be other builds with different Operating Systems, find the default description of Windows IBS image (custom built) and OS build number
# Filter out current uploaded image
# Filter out base name from wim file (task sequence name)
$OtherMDTOSs = $MDTOSsFile.oss.os | Where {($_.Name -like "*$WIMPatchedBaseName*") -and ($_.Description -eq "Windows IBS image") -and ($_.Build -eq $NewMDTOSBuild) -and ($_.Name -ne "$($newCapture.Name)")}
Foreach ($OtherOS in $OtherMDTOSs){
    [datetime]$OSCreationDate = (Get-Date $OtherOS.CreatedTime -Format MMddyyyy)
    If($OSCreationDate -le $NewMDTOSDate.AddDays(-30)){
        #with mdt module imports and psdrive mapperd, remove item will remvoe if from both xml and directory
        Remove-Item -path "$($MDTDrive.Name):\Operating Systems\$MDT_Dest_Folder\$BackupFile" -Force -Confirm:$false
        #eg. Remove-Item -path "DS001:\Operating Systems\Patched\WIN101803_BC01CDrive in WIN101803X64_Patched_10-18-2018 WIN101803X64_Patched_10-18-2018.wim" -Verbose
    }
    Else{
        Write-Host ("WIM's [{0}] isn't older than [{1}], skipping..." -f $OtherOS.Name, $NewMDTOSDate.AddDays(-30)) -ForegroundColor Gray
    }
}

#Update Tasksequences with new wim 
$MDTTSsFile = [Xml] (Get-Content "$MDTSharePath\Control\TaskSequences.xml")
Foreach ($ts in $MDTTSSFile.tss.ts){
    $tsguid = $ts.guid
    $tsname = $ts.ID
    $TSXML = [Xml] (Get-Content "$MDTSharePath\Control\$tsname\ts.xml")
	$OSGUID = (($TSXML.sequence.globalVarList | Select -ExpandProperty variable) | Where {$_.name -eq "OSGUID"} | Select -First 1).'#text'
    If($OSGUID -eq $deletedGUIDS){
        
    }



    #If($tsenv.Value("TaskSequenceID") -eq $ts.ID){
        
	#}


}




 
# End user defined variables
#$MDT_Physical_Path = "D:\MDT_Build"
#$MDT_Dest_Folder = "D:\MDT_Production"
$RefTaskID = "REFW10-X64-001"
$ProdTaskID = "PRODW10-X64-001"
$TestTaskID = "TESTW10-X64-001"
$OS = "Windows 10"
 

$Prod_Control_Path = $MDT_Dest_Folder + "\Control"
 
# Get the file name of the most recent capture
$RefImg = Get-ChildItem $Build_Captures_Path -Filter ($RefTaskID + "*") | Sort LastWriteTime | Select -Last 1
 
# Drop .wim from name and save to separate variable
$RefImgOSName = $RefImg.Name.Replace('.wim','')
 
# Set up PSDrive to Production Deployment Share
New-PSDrive -Name "DS002" -PSProvider MDTProvider -Root $MDT_Dest_Folder
 
# Import RefImg to Production DS Operating Systems
Import-MDTOperatingSystem -Path "DS002:\Operating Systems\$($OS)" -SourceFile $RefImg.FullName -DestinationFolder $RefImgOSName
 
# Rename the Operating System to clean it up
$RefImgOSCurrentName = Get-ChildItem "DS002:\Operating Systems\$($OS)" | Where {$_.Name -like "* in *"}
Rename-Item -Path "DS002:\Operating Systems\$($OS)\$($RefImgOSCurrentName.Name)" -NewName $RefImgOSName
 
# Get GUID of OS
$RefImgOSguid = (Get-ItemProperty "DS002:\Operating Systems\$($OS)\$($RefImgOSName)").guid
 
# Delete RefImg Capture from Build DS
Remove-Item $RefImg.FullName -Force
 
# Copy Production TS to Test TS
Copy-Item "$Prod_Control_Path\$($ProdTaskID)\*" "$Prod_Control_Path\$($TestTaskID)" -Force
 
# Modify Test Task Sequence to use new RefImg OS
$TSPath = "$Prod_Control_Path\$TestTaskID\ts.xml"
$TSXML = [xml](Get-Content $TSPath)
$TSXML.sequence.globalVarList.variable | Where {$_.name -eq "OSGUID"} | ForEach-Object {$_."#text" = $RefImgOSguid}
$TSXML.sequence.group | Where {$_.Name -eq "Install"} | ForEach-Object {$_.step} | Where {
    $_.Name -eq "Install Operating System"} | ForEach-Object {$_.defaultVarList.variable} | Where {
    $_.name -eq "OSGUID"} | ForEach-Object {$_."#text" = $RefImgOSguid}
$TSXML.Save($TSPath)