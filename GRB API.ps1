 <# ==================================================================
    Author: Jan Speecke
    Last Updated : 2024 03 20 - 08:41
    Version	:  1.2
    Comments :
    Changes:
        22/10/2023 - Added function to Unzip the downloaded ZIP-file
        22/10/2023 - Added some logging and reporting
        06/03/2024 - Changed URL to production
        06/03/2024 - Delete downloaded file after unzipping
        20/03/2024 - Delete base folder before the rest of the script
    TO DO:
        -
        -
    ==================================================================

<#

.SYNOPSIS
This Powershell will be used to:
    - Use API from Digitaal Vlaanderen
    - Place orders via the API
    - Check if the file is ready to download
    - Download the files

.DESCRIPTION

.EXAMPLE

.NOTES
    Information can be found here: https://www.vlaanderen.be/digitaal-vlaanderen/onze-oplossingen/downloadtoepassing/download-api-v2

    Testing: https://download.api.beta-vlaanderen.be/
    Production: https://download.api.vlaanderen.be/

    We're using the V2 API
    Documentation on this API can be found on https://download.api.beta-vlaanderen.be/docs/v2/api-documentation.html

.LINK
#>

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

# PARAMETERS
# ====================================================================
    $URIBase    = "https://download.api.vlaanderen.be/v2/"          # BaseURL - replace with "https://download.api.vlaanderen.be/" for production environment
    $APIKey     = "your ipakey here"                                # APIKey requested form Vlaanderen
    $header     = @{"x-api-key"=$APIKey}                            # Create Header containing APIKey
    $Script:Path       = "\\Fileserver\Drive-Z\DIPS\Download\"      # Path to download resulting file(s)
    $Script:UnZIPPath  = "\\Fileserver\Drive-Z\DIPS\Gemeentes\"
    $Script:BaseFolder = "\\Fileserver\Drive-Z\DIPS"


    $Municipalities = @{
31003 = 'Beernem'
31004 = 'Blankenberge'
31005 = 'Brugge'
31006 = 'Damme'
31012 = 'Jabbeke'
31022 = 'Oostkamp'
31033 = 'Torhout'
31040 = 'Zedelgem'
31042 = 'Zuienkerke'
31043 = 'Knokke-Heist'
32003 = 'Diksmuide'
32006 = 'Houthulst'
32010 = 'Koekelare'
32011 = 'Kortemark'
32030 = 'Lo-Reninge'
33011 = 'Ieper'
33016 = 'Mesen'
33021 = 'Poperinge'
33029 = 'Wervik'
33037 = 'Zonnebeke'
33039 = 'Heuvelland'
33040 = 'Langemark-Poelkapelle'
33041 = 'Vleteren'
34027 = 'Menen'
34040 = 'Waregem'
35002 = 'Bredene'
35005 = 'Gistel'
35006 = 'Ichtegem'
35011 = 'Middelkerke'
35013 = 'Oostende'
35014 = 'Oudenburg'
35029 = 'De Haan'
36006 = 'Hooglede'
36007 = 'Ingelmunster'
36008 = 'Izegem'
36010 = 'Ledegem'
36011 = 'Lichtervelde'
36012 = 'Moorslede'
36015 = 'Roeselare'
36019 = 'Staden'
37002 = 'Dentergem'
37007 = 'Meulebeke'
37010 = 'Oostrozebeke'
37011 = 'Pittem'
37012 = 'Ruiselede'
37015 = 'Tielt'
37017 = 'Wielsbeke'
37018 = 'Wingene'
37020 = 'Ardooie'
38002 = 'Alveringem'
38008 = 'De Panne'
38014 = 'Koksijde'
38016 = 'Nieuwpoort'
38025 = 'Veurne'
    }


    # LOG
    $Script:Logfile = $Script:Path  + "WVI_API_Vlaanderen_Download.log"


    # EMAIL
    $Mailsubject = "[GRB] - Automatic download GRB via API..."
    $Mailbody = ""
    $SMTPDestination = @("j.speecke@wvi.be","m.decraemer@wvi.be")
    $SMTPSource = "jobserver@wvi.be"
    $SMTPServer = "mail.wvi.be"

# FUNCTIONS
# ====================================================================

# OUTPUT FUNCTIONS
function WriteTitle($message){
    $message = $message.ToUpper()
    Write-Host "`n`n`n============== $message ==============`n"  -ForegroundColor Cyan
    Write-Log "`n============== $message =============="
}

function WriteSubTitle($message){
    Write-Host "`n$message`n-----------------------------------------------------"  -ForegroundColor Cyan
    Write-Log "$message`n-----------------------------------------------------"
}

function WriteInfo($message){
$DateNow = Get-Date -Format 'yyyyMMdd HH:mm:ss'
Write-Host "    $DateNow - [+] $message"
Write-Log "    $DateNow - [+] $message"
}

function WriteInfoHighlighted($message){
$DateNow = Get-Date -Format 'yyyyMMdd HH:mm:ss'
Write-Host "    $DateNow - [+] $message" -ForegroundColor Yellow
Write-Log "    $DateNow - [+] $message"
}

function WriteSuccess($message){
$DateNow = Get-Date -Format 'yyyyMMdd HH:mm:ss'
Write-Host "    $DateNow - [OK] $message" -ForegroundColor Green
Write-Log "    $DateNow - [OK] $message"
}

function WriteError($message){
$DateNow = Get-Date -Format 'yyyyMMdd HH:mm:ss'
Write-Host "    $DateNow - [NOK] $message" -ForegroundColor Red
Write-Log "    $DateNow - [NOK] $message"
}

function WriteErrorAndExit($message){
Write-Host $message -ForegroundColor Red
Write-Host "    $DateNow - [NOK] Press enter to continue ..."
Write-Log
Stop-Transcript
Read-Host | Out-Null
Exit
}

function Start-Logging {
    Remove-Item $Script:Logfile -Force
    New-Item $Script:Logfile -ItemType File
}

function Write-Log {
    param
    (
        [Parameter(ValueFromPipeline)]
        [string]$content
    )
    Add-Content -Path $script:Logfile -Value $content # + "`n"
    # Usage:
    #  - Write-Output "hello world" | Write-Log
    #  - Write-Log -content "Just another line of content"
}

function Get-TotalJobTime {
    # $start_time = Get-Date
    $JobTimeDays = (Get-Date).Subtract($start_time).Days
    $JobTimeDays = If ($JobTimeDays -eq 0) { "" } else { " $JobTimeDays days" }
    $JobTimeHours = (Get-Date).Subtract($start_time).Hours
    $JobTimeHours = If ($JobTimeHours -eq 0) { "" } else { " $JobTimeHours hours" }
    $JobTimeMinutes = (Get-Date).Subtract($start_time).Minutes
    $JobTimeMinutes = If ($JobTimeMinutes -eq 0) { "" } else { " $JobTimeMinutes minutes" }
    $JobTimeseconds = (Get-Date).Subtract($start_time).Seconds
    $JobTimeseconds = If ($JobTimeseconds -eq 0) { "" } else { " $JobTimeseconds seconds" }
    $TotalJobTime = "This action took$JobTimeDays$JobTimeHours$JobTimeMinutes$JobTimeseconds to complete."
    $TotalJobTime
}

Function Delete-BaseFolder {
    Remove-Item $Script:BaseFolder -Recurse -Force
    }

Function Create-folder {
    New-Item $Script:BaseFolder -ItemType Directory -Force
    If (!(Test-Path $Script:Path)) {
    New-Item $Script:Path -ItemType Directory -Force
    }

    If (!(Test-Path $Script:UnZIPPath)) {
    New-Item $Script:UnZIPPath -ItemType Directory -Force
    }
}

function PlaceOrder{                    # Place order One municipality
    param (
    [Parameter()] [string] $NISCode
    )
    $Endpoint = "Orders"
    $Method = "POST"
    $Entity = "Gemeente"
    $EntityCode = @($NISCode)
    $body = @{
            "productId" = 1;
            "format" = "Shapefile"
            "geographicalCrop" = @{
                "selectionEntityCode"= $Entity
                "selectionEntityCodeValue" = $EntityCode
            }
        } | ConvertTo-Json
    $Uri = $URIBase + $Endpoint

    WriteInfo "An order will be placed for NISCode $NISCode..."
    $Script:Orderfeedback = Invoke-Restmethod -Uri $URI -Method $Method -Headers $header -Body $body -ContentType 'application/json'

    # Get OrderID from the feedback
    $Script:OrderID = $Script:Orderfeedback.member.orderID
    }

function GetOrderInformation{           # Get Order information for the specified order
        param (
        [Parameter()] [string] $Script:OrderID
        )

        $Download_Name = $null
        $Download_Volume = $null
        $Download_ID = $null

        $Endpoint = "orders/$Script:OrderID"
        $Method = "GET"
        $Uri = $URIBase + $Endpoint
        $Orderdetail = Invoke-Restmethod -Uri $URI -Method $Method -Headers $header -ContentType 'application/json'
        $script:Download_Name = $Orderdetail.downloads.name
        $script:Download_Volume = $Orderdetail.downloads.volume
        $script:Download_ID = $Orderdetail.downloads.FileID

    }

function GetDownloadInformation{        # Wait until order is completed...  If not, the script will loop till ready and then continue
    $count = 0
    $success = $null
    $retries = 20
    $RetryInterval = 60

    do{
        If ($Null -ne $script:Download_ID) {
            $success = $true
        #    Write-Host "   Bingo!  Seems like we have a winner...The script will continue." -ForegroundColor Green
        }
        else {
        #    Write-Output "   Did not succeed... Next attempt in $RetryInterval seconds"
            GetOrderInformation $Script:OrderID
            Start-sleep -Seconds $RetryInterval        }
        $count++
    }until($count -eq $retries -or $success)
    #if(-not($success)){exit}
    WriteInfo "File $Script:Download_Name from order $Script:OrderID is ready to download..."
    }

function DownloadFile{
    $Endpoint = "orders/$Script:OrderID/download/$script:Download_ID"
    $Method = "GET"
    $Uri = $URIBase + $Endpoint
    $FileName = $script:Download_Name
    $Script:FileNameFull = $Script:Path + $FileName
    WriteInfo "The order will be downloaded to $Script:FileNameFull..."
    Invoke-WebRequest -Uri $URI -Method $Method -Headers $header -OutFile $Script:FileNameFull -TimeoutSec 10000
    WriteSuccess "Downloaded completed.  Resulting file: $Script:FileNameFull"
}


Add-Type -AssemblyName System.IO.Compression.FileSystem
Function Unzip {
    param([string]$zipfile, [string]$outpath)
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}

Function Unzip-Downloaded-File {
    WriteInfo "Unzipping $Script:FileNameFull to folder $Outpath ..."
    $outpath = $UnZIPPath+"$Municipality_Name"
    WriteInfo "Deleting destination folder: $Script:FileNameFull ..."
    Remove-Item $outpath -Force -Recurse
    Unzip $Script:FileNameFull $outpath
    WriteSuccess "Archive $Script:FileNameFull unzipped to folder $Outpath "
    Remove-Item $Script:FileNameFull -Force
}

# FUNCTION - Send-mail
Function Send-EMail {
    $script:Mailbody = @"
    <p>Hey there,</p>
    <p>I downloaded some freshly baked GRB files for you from $URIBase.</p>
    <p>Please check out if everything looks fine...&nbsp;If not so, start yelling and crying... (or try to pinpoint and fix the issue)</p>
    <h2>Resulting files:</h2>
    <p>Downloaded files can be found here: $Script:UnZIPPath</p>
    <p><a href="$Script:UnZIPPath">Open the folder</a></p>
    <p>&nbsp;</p>
    <p>Everything looks fine? Well, have a nice day!<br />If not so, have a nice day to!</p>
    <p>Your humble and servant<br />SRV16 (a.k.a. jobserver)</p>
"@
    Send-MailMessage -To $SMTPDestination -From $SMTPSource -Subject $Mailsubject -Body ($script:Mailbody | out-String) -SmtpServer $SMTPServer -BodyAsHtml -Attachments $Script:Logfile
    }



# PROGRAM
# ====================================================================


    Start-Logging
    Delete-BaseFolder
    Create-folder


Foreach ($Municipality in $Municipalities.GetEnumerator()) {
    $Municipality_Name = $Municipality.value
    $Municipality_NIS = $Municipality.Name
    $NISCode = $Municipality.Name

    $start_time = Get-Date
    WriteSubTitle "NIS CODE $NISCode ($Municipality_Name) WILL BE PROCESSED..."
    PlaceOrder $NISCode                 # Place the order
    GetOrderInformation $Script:OrderID # Get order information
    GetDownloadInformation              # Get Download information for the order
    DownloadFile                        # Download the file
    Unzip-Downloaded-File               # Unzip the downloaded file
    WriteSuccess "NIS CODE $NISCode ($Municipality_Name) COMPLETED"
    Get-TotalJobTime #  Put following command at beginning of the job execution:   $start_time = Get-Date
}
    Send-EMail





