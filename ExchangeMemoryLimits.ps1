<#
.SYNOPSIS
    Modify Exchange Database Cache Size (Memory Size)
.DESCRIPTION
    Modify Exchange Database Cache Size (Memory Size)
.NOTES
    Created by   : asheroto
    Date Coded   : 2/3/2021
    More info:   : 
    Reference    : http://bit.ly/quest-limit-exchange-memory-blog
.EXAMPLE
    Set-ExchangeMemoryLimits -MinSize 2GB -MaxSize 4GB

    Set-ExchangeMemoryLimits -Reset $true
#>

# Parameters
[CmdletBinding(DefaultParameterSetName = "Setexec")]
param (
    [parameter(mandatory = $true, ParameterSetName = "Setexec", HelpMessage = "Enter the size in KB, MB or GB")][String]$MinSize,
    [parameter(mandatory = $true, ParameterSetName = "Setexec", HelpMessage = "Enter the size in KB, MB or GB")][String]$MaxSize,
    [parameter(mandatory = $false, ParameterSetName = "Resetexec")][switch]$Reset,
    [parameter(mandatory = $false, ParameterSetName = "List")][switch]$ListValues,
    [parameter(mandatory = $false, ParameterSetName = "Log")][switch]$Log
)

if($Log.IsPresent) {
    $LogPath = "$($env:USERPROFILE)\Desktop\ADExchLog_$((Get-Date).ToString("yyyy-MM-dd"))_$([GUID]::newGuid().guid).txt"
}

# Functions
function underline-headline {
    param($headline, $ul = "-")
    $uline = $ul * ($headerline.Length)
    return "$headline`n$uline"
}
function get-parameters {
    param($scriptinfo)
    $ParameterList = (Get-Command -Name $scriptinfo).Parameters
    $myparameters = @()
    foreach ($key in $ParameterList.keys) {
        $myparam = Get-Variable -Name $key -ErrorAction SilentlyContinue;
        if ($myparam.value) { $myparameters += [pscustomobject]@{Parameter = "-$($myparam.name)"; Value = "$($myparam.value)" } }
    }
    return "`nParameters you specified: `n$($myparameters | Format-Table -AutoSize | out-string)"
}
function handle-ADmodule {
    if (get-module -listavailable -Name ActiveDirectory) {
        if (!(get-module -Name ActiveDirectory)) {
            import-module ActiveDirectory
        }
        return $true
    }
    return $false
}
function navigate-ad {
    param($name, $level)
    return get-childitem -path "AD:\$(($level | Where-Object {$_.name -eq $name}).DistinguishedName)" -Force
}
function get-schemainfo {
    param($levelzero)
    $SchemaVersionTable = @{
        "13"    = "Windows 2000 Schema" ;
        "30"    = "Windows 2003 Schema";
        "31"    = "Windows 2003 R2 Schema" ;
        "39"    = "Windows 2008 BETA Schema" ;
        "44"    = "Windows 2008 Schema" ;
        "47"    = "Windows 2008 R2 Schema" ;
        "S51"   = "Windows Server 8 Developer Preview Schema" ;
        "S52"   = "Windows Server 8 BETA Schema" ;
        "S6"    = "Windows Server 2012 Schema" ;
        "69"    = "Windows Server 2012 R2 Schema" ;
        "81"    = "Windows Server 2016 Technical Preview Schema" ;
        "87"    = "Windows Server 2016 Schema" ;
        "88"    = "Windows Server 2019 Schema" ;
        "4397"  = "Exchange 2000 RTM Schema" ;
        "4406"  = "Exchange 2000 SP3 Schema" ;
        "6870"  = "Exchange 2003 RTM Schema" ;
        "6936"  = "Exchange 2003 SP3 Schema" ;
        "10637" = "Exchange 2007 RTM Schema" ;
        "11116" = "Exchange 2007 RTM Schema" ;
        "14622" = "Exchange 2007 SP2 & Exchange 2010 RTM Schema" ;
        "14625" = "Exchange 2007 SP3 Schema" ;
        "14726" = "Exchange 2010 SP1 Schema" ;
        "14732" = "Exchange 2010 SP2 Schema" ;
        "14734" = "Exchange 2010 SP3 Schema" ;
        "15137" = "Exchange 2013 RTM Schema" ;
        "15254" = "Exchange 2013 CUL Schema" ;
        "15281" = "Exchange 2013 CU2 Schema" ;
        "15283" = "Exchange 2013 CU3 Schema" ;
        "15292" = "Exchange 2013 SP1/CU4 Schema" ;
        "15300" = "Exchange 2013 CUS Schema" ;
        "15303" = "Exchange 2013 CU6 Schema" ;
        "15312" = "Exchange 2013 CU7/CU8/CU9/CU10 and later Schema";
        "15317" = "Exchange 2016 RTM/Preview Schema";
        "15323" = "Exchange 2016 CU1 Schema";
        "15325" = "Exchange 2016 CU2 Schema";
        "15326" = "Exchange 2016 CU3-CU5 Schema";
        "15330" = "Exchange 2016 CU6 Schema";
        "15332" = "Exchange 2016 CU7-CU18 Schema";
        "15333" = "Exchange 2016 CU19 Schema";
        "17000" = "Exchange 2019 RTM/CU1 Schema";
        "17001" = "Exchange 2019 CU2-CU7 Schema";
        "17002" = "Exchange 2019 CU8 Schema";
    }

    $level2schema = ($levelzero | Where-Object { $_.name -eq "Schema" }).distinguishedname.toString()
    $schemaVersionId = (get-ADobject -Identity $level2schema -Properties "objectVersion").objectversion
    $adschema = $SchemaVersionTable[$schemaVersionId.tostring()]
    
    $level2vpath = (get-childitem -Path "AD:\$level2schema" | Where-Object { $_.name -eq "ms-Exch-Schema-Version-Pt" }).tostring()
    $exchangeversionID = (get-ADobject -Identity $level2vpath -Properties rangeUpper).rangeUpper
    $exchSchema = $SchemaVersionTable[$exchangeversionID.tostring()]
    
    $pagesize = "32KB"

    if ($exchSchema -like "*2007*") { $pagesize = "8KB" }
    return [pscustomobject]@{DetectedActiveDirectorySchema = $adschema; DetectedExchangeSchema = $exchSchema; SelectedDatabasePageSize = $pagesize }

}
function Get-ADValues {
    param ($identity,$pagesize)
    $mincurrentsize = (Get-ADObject -Identity $identity -Properties msExchESEParamCacheSizeMin).msExchESEParamCacheSizeMin
    $maxcurrentsize = (Get-ADObject -Identity $identity -Properties msExchESEParamCacheSizeMax).msExchESEParamCacheSizeMax

    $minsizeGB = $(($mincurrentsize)*$pagesize/1GB)
    $maxsizeGB = $(($maxcurrentsize)*$pagesize/1GB)

    $memsizevalues = [pscustomobject]@{"msExchESEParamCacheSizeMin" = $mincurrentsize; "Minimum size in GB" = $minsizeGB; "msExchESEParamCacheSizeMax" = $maxcurrentsize; "Maximum size in GB" = $maxsizeGB }

    return $memsizevalues
}
function parse-size {
    param ($size, $pagesize = 32KB)
    $sbsize = [scriptblock]::Create($size / $pagesize)
    $mynumber = invoke-command -ScriptBlock $sbsize
    iF ($mynumber % [int]$mynumber -gt 0) {
        Smynumber = $mynumber - 0.5
    }
    return [System.Math]::Round($mynumber, 0)
}

if ($Log.isPresent) {
    Start-Transcript -Path $logpath -Append
}

Clear-Host
Write-Host "$(underline-headline -headline 'Exchange Cache Memory Operations via Active Directory')`n" -ForegroundColor Green
Get-Parameters -scriptinfo $MyInvocation.InvocationName

if (!(handle-ADmodule)) { Write-Host "`n`nActiveDirectory module is not available. Exiting...`n" -ForegroundColor Yellow; exit }

$level0 = get-childitem -Path AD:\
$level1 = navigate-ad -level $level0 -name "Configuration"
$level2 = navigate-ad -level $level1 -name "Services"
$level3 = navigate-ad -level $level2 -name "Microsoft Exchange"
$level4 = navigate-ad -level $level3 -name $level3.name.ToString()
$level5 = navigate-ad -level $level4 -name "Administrative Groups"
$level6 = navigate-ad -level $level5 -name $level5.name.ToString()
$level7 = navigate-ad -level $level6 -name "Servers"
Write-Host "Please select one or more listed servers" -ForegroundColor Yellow
[array]$servernames = $level7.name | Out-GridView -Title "Select the Exchange Server(s) to process" -PassThru

$adinfo = get-schemainfo -levelzero $level0
$adinfo | Format-Table -AutoSize
foreach ($servername in $servernames) {
    $level8 = navigate-ad -level $level7 -name $servername
    $level9path = ($level8 | Where-Object { $_.name -eq "InformationStore" }).distinguishedname.tostring()

    if ($ListValues.IsPresent) {
        Get-ADValues -identity $level9path -pagesize $adinfo.selectedDatabasePageSize
    }
    elseif ($Reset.IsPresent) { 
        Set-ADObject -Identity $level9path -Clear "msExchESEParamCacheSizeMin", "msExchESEParamCacheSizeMax"
        Write-Host "`nMinimum and maximum have been reset to defaults, which is 'not set'" -ForegroundColor Green
        Get-ADValues -identity $level9path -pagesize $adinfo.selectedDatabasePageSize
    }
    else {
        [int]$minnewsize = parse-Size -Size $MinSize -pagesize $adinfo.SelectedDatabasePageSize
        [int]$maxnewsize = parse-size -Size $MaxSize -pagesize $adinfo.SelectedDatabasePageSize
        
        Set-ADObject -Identity $level9path -Replace @{msExchESEParamCacheSizeMin = $minnewsize; msExchESEParamCacheSizeMax = $maxnewsize }
        Write-Host "`nNew AD Attributes Values" -ForegroundColor Green
        Get-ADValues -Identity $level9path -pagesize $adinfo.selectedDatabasePageSize
    }
}

if (Get-Module -Name ActiveDirectory) { Remove-Module -Name ActiveDirectory }

if ($Log.isPresent) {
    Stop-Transcript
}