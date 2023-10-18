 <# ==================================================================
    Author: Jan Speecke
    Last Updated :17/10/2023 - 5:49:19
    Version	:  0.9
    Comments :
    Changes:
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

# PARAMETERS
# ====================================================================
    $URIBase    = "https://download.api.beta-vlaanderen.be/v2/"     # BaseURL - replace with "https://download.api.vlaanderen.be/" for production environment
    $APIKey     = "%%%YOUR-API-CODE-HERE%"                          # APIKey requested form Vlaanderen
    $header     = @{"x-api-key"=$APIKey}                            # Create Header containing APIKey
    $Path       = "Z:\DIPS\Download\"                               # Path to download resulting file(s)

    $NISCodes = @(
        31003
        32010
        32011
        32030
        33011
        33016
        37007
        37010
        37011
        )



# FUNCTIONS
# ====================================================================
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

    Write-Host "   --> An order will be places for NISCode $NISCode..."
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
    Write-Host "   --> File $Script:Download_Name from order $Script:OrderID is ready to download..."
    }

function DownloadFile{
    $Endpoint = "orders/$Script:OrderID/download/$script:Download_ID"
    $Method = "GET"
    $Uri = $URIBase + $Endpoint
    $FileName = $script:Download_Name
    $FileNameFull = $Path + $FileName
    Write-Host "   --> The order will be downloaded to $FileNameFull..."
    Invoke-WebRequest -Uri $URI -Method $Method -Headers $header -OutFile $FileNameFull -TimeoutSec 10000
    Write-Host "   --> Downloaded completed.  Resulting file: $FileNameFull"
}

# PROGRAM
# ====================================================================

ForEach ($NISCode in $NISCodes) {

    Write-Host "NIS CODE $NISCode WILL BE PROCESSED...`n-------------------------------------------" -ForegroundColor Green
    PlaceOrder $NISCode                 # Place the order
    GetOrderInformation $Script:OrderID # Get order information
    GetDownloadInformation              # Get Download information for the order
    DownloadFile                        # Download the file
}

