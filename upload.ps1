<#
   .SYNOPSIS
      Upload all images from a given directory to Transkribus and create a document
   .NOTES
      Author: Dario Kampkaspar <dario.kampkaspar@oeaw.ac.at>
   .PARAMETER $processId
      ID of the current goobi process
   .PARAMETER $imagePath
      ID of the current goobi process
   .PARAMETER $collection
      The ID of the Transkribus collection to use
   .PARAMETER $documentName
      The name for the document within Transkribus
   .PARAMETER $transkribusSecrets
      A path to a secrets file to use when logging in to Transkribus
#>

param(
   [Parameter(Mandatory=$true)] [string] $processId,
   [Parameter(Mandatory=$true)] [string] $imagePath,
   [Parameter(Mandatory=$true)] [string] $collection,
   [Parameter(Mandatory=$true)] [string] $documentName,
   $transkribusSecrets
)

##################################   STEP 0: Initialise variables and connection   #####################################
If ( $transkribusSecrets.IsPresent ) {
   $secrets = Get-Content $transkribusSecrets | ConvertFrom-Json
} Else {
   $secrets = Get-Content $imagePath + "../transkribusSecrets" | ConvertFrom-Json
}

If ( $secrets.user -eq "" -or $secrets.pass -eq "" ) {
   Throw "No credentials supplied in $transkribusSecrets"
}

# log in to Transkribus REST service
# $session will hold the session information needed for subsequent calls to the API
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$parms = @{ "Uri" = "https://transkribus.eu/TrpServer/rest/auth/login";
            "Body" = "user=$user&pw=$pass";
            "Method" = "Post";
            "SessionVariable" = "session"
         }
Invoke-RestMethod @parms | Out-Null

$jobList = "https://transkribus.eu/TrpServer/rest/jobs/list?collId=$collectionID"
$jobRequest = "https://transkribus.eu/TrpServer/rest/jobs/"
########################################################################################################################

##########################################   STEP 1: upload images via FTP   ###########################################
# Create directory on Transkribus FTP
$dir = "ftp://transkribus.eu/$documentName"

$request = [Net.WebRequest]::Create($dir)
$request.Credentials = New-Object System.Net.NetworkCredential($user, $pass)

# if the directory already exists, we simply upload again, no need to throw an error
try {
   $request.Method = [System.Net.WebRequestMethods+FTP]::MakeDirectory
   $resp = $request.GetResponse()
   $resp.Close()
} Catch {}

# Upload to Transkribus via FTP
$imageDirectory = $imagePath + '/master_abc_media'
$items = Get-ChildItem $imageDirectory

$webclient = New-Object System.Net.WebClient
$webclient.Credentials = New-Object System.Net.NetworkCredential($secrets.user, $secrets.pass)

ForEach ( $item in $items ) {
   $webclient.UploadFile($uri, $item.FullName) | Out-Null
}
########################################################################################################################

##########################################   STEP 1: upload images via FTP   ###########################################
# start import in Transkribus
$parms = @{ "Uri" = "https://transkribus.eu/TrpServer/rest/collections/$collectionID/ingest?fileName=$documentName";
            "Method " = "Post";
            "WebSession" = $session
         }
$importJob = Invoke-RestMethod @params

# check every 60 seconds if document creation was successful
$documentJob = $jobRequest + $importJob
do {
   Start-Sleep 60
   $importStatus = Invoke-RestMethod -Uri $documentJob -Method Get -WebSession $session
} While ( $importStatus.state -ne 'FAILED' -and $importStatus.state -ne 'FAILED' )
########################################################################################################################

Exit 0
