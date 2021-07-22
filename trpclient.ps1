<# 
.PARAMETER $collection
  The Transkribus collection ID
.PARAMETER $j
  List all jobs running on documents in the collection
.PARAMETER $p
  Whether to start printed block detection
.PARAMETER $c
  Whether to start CITlab Advanced LA
.PARAMETER $h
  ID of the HTR model to be used
.PARAMETER $e
  export collection to basic TEI
.PARAMETER $from
  Lowest ID of document in $collection to be used
.PARAMETER $to
  Highest ID of document in $collection to be used
.PARAMETER $user
  Transkribus user (usually the eMail used to sign up)
.PARAMETER $pass
  Transkribus password – we do not expect it to contain white space!
#>
param(
  [Parameter(Mandatory)]
    [int] $collection,
  [switch] $j,
  [switch] $p,
  [switch] $c,
  [switch] $e,
  [int] $h,
  [int] $from,
  [int] $to,
  [Parameter(Mandatory)]
    [string] $user,
  [Parameter(Mandatory)]
    [string] $pass
)

########################################################################################################################
##### Functions
function Request-Status {
    param (
      [int[]] $jobIds,
      $session
    )
  
    $openJobs = @()
    $jobIds | ForEach-Object {
      $statusRequest = "https://transkribus.eu/TrpServer/rest/jobs/" + $_
      try {
        $md = Invoke-RestMethod -Uri $statusRequest -Method Get -WebSession $session
      } catch {
        Invoke-RestMethod -Uri https://transkribus.eu/TrpServer/rest/auth/login -Body "user=$user&pw=$pass" -Method Post -SessionVariable session | Out-Null
        $md = Invoke-RestMethod -Uri $statusRequest -Method Get -WebSession $session
      }
  
      if ( $md.state -ne "FINISHED" -and $md.state -ne "FAILED" ) {
        $openJobs += $_
      }
    }
  }

########################################################################################################################
$startTime = $(Get-Date)

"Starting batch processing for collection $collection at $startTime"

# log in to Transkribus REST service
# $session will hold the session information needed for subsequent calls to the API
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  Invoke-RestMethod -Uri https://transkribus.eu/TrpServer/rest/auth/login -Body "user=$user&pw=$pass" -Method Post -SessionVariable session | Out-Null

########################################################################################################################
####### Get a list of all document IDs in the collection

$mdReq = "https://transkribus.eu/TrpServer/rest/collections/$collection/list"
$md = Invoke-RestMethod -Uri $mdReq -Method Get -WebSession $session

$documents = $md | ForEach-Object { $_ }
$errors = @()
$jobList = @()
########################################################################################################################

########################################################################################################################
##### get a list of all unfinished jobs

if ( $j.IsPresent ) {
  ("CREATED", "WAITING", "RUNNING") | ForEach-Object {
    $statusRequest = "https://transkribus.eu/TrpServer/rest/jobs/list?collId=95080&status=" + $_
    $list = Invoke-RestMethod -Method Get -Uri $statusRequest -WebSession $session

    $list | ForEach-Object {
      $jobList += $_.jobId
    }
  }

  $jobList
  $jobList.Length.ToString() + " open jobs in collection " + $collection
}

########################################################################################################################
##### printed block detection

if ( $p.IsPresent ) {
  $pJobs = @()

  $documents | ForEach-Object {
    $docId = $_.docId
    $numP = $_.numberOfPages
    $request = "https://transkribus.eu/TrpServer/rest/recognition/ocr?collId=$collection&id=$docId&pages=1-$numP&doBlockSegOnly=true"

    "Starting PBD for $docId"
    try {
      $rOcr = Invoke-RestMethod -Uri $request -Method Post -WebSession $session
      $pJobs +=  $rOcr
    } catch {
      $_.Exception.Response.StatusCode.value__
    }
  }

  $wait = 20 * $documents.Length
  $ts = New-TimeSpan -Seconds $wait
  "Waiting to allow printed block detection to finish until " + ((Get-Date) + $ts)
  Start-Sleep -Seconds $wait

  $openJobs = Request-Status $pJobs $session
  do {
    $wait = 15 * $openJobs.length
    "Waiting for another $wait seconds to allow " + $openJobs.length + " PBD jobs to finish"
    Start-Sleep -Seconds $wait
  } while ( $openJobs.length -gt 0 )
}

########################################################################################################################
##### Line Recognition (CitLab Advanced LayoutAnalysis)

if ( $c.IsPresent ) {
  $LARequest = "https://transkribus.eu/TrpServer/rest/LA?collId=$collection"
  Invoke-RestMethod -Uri https://transkribus.eu/TrpServer/rest/auth/login -Body "user=$user&pw=$pass" -Method Post -SessionVariable session | Out-Null

  $cJobs = @()
  $documents | ForEach-Object {
    $docId = $_.docId
    
    # Get list of pages for recognition job parameters
    $pagesReq = "https://transkribus.eu/TrpServer/rest/collections/$collection/$docId/fulldoc"
    $pa = Invoke-RestMethod -Uri $pagesReq -Method Get -WebSession $session
    $ps = $pa.pageList.pages | ForEach-Object {
      $pageId = $_.pageId
      
      "<pages><pageId>$pageId</pageId></pages>"
    }
    $jobParams = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
      <jobParameters>
        <docList>
          <docs>
            <docId>$docId</docId>
            <pageList>$ps</pageList>
          </docs>
        </docList>
        <params>
          <entry>
            <key>modelName</key>
            <value>LA_news_onb_att_newseye.pb</value>
          </entry>
        </params>
      </jobParameters>"

    "Starting line detection for $docId"
    try {
      $newJob = Invoke-RestMethod -Uri $LARequest -Method Post -WebSession $session -Body $jobParams -ContentType application/xml
      
      $cJobs += $newJob.trpJobStatuses.trpJobStatus.jobId
    } catch {
        $errors | Add-Member -NotePropertyName $docId -NotePropertyValue $_.Exception
        $_
    }
  }

  $wait = 20 * $documents.Length
  $ts = New-TimeSpan -Seconds $wait
  "Waiting to allow line detection to finish until " + ((Get-Date) + $ts)
  #Start-Sleep -Seconds $wait
  timeout /t $wait

  $openJobs = Request-Status $cJobs $session
  do {
    $wait = 15 * $openJobs.length
    "Waiting for another $wait seconds to allow " + $openJobs.length + " LR jobs to finish"
    #Start-Sleep -Seconds $wait
    timeout /t $wait
  } while ( $openJobs.length -gt 0 )
}

########################################################################################################################
##### Text Recognition

if ( $h.length -eq 1 -and $h -gt 0 ) {
  Invoke-RestMethod -Uri https://transkribus.eu/TrpServer/rest/auth/login -Body "user=$user&pw=$pass" -Method Post -SessionVariable session | Out-Null

  $numAllPages = 0
  $hJobs = @()
  $documents | ForEach-Object {
    $docId = $_.docId
    "Starting HTR for $docId"

    $pagesReq = "https://transkribus.eu/TrpServer/rest/collections/$collection/$docId/fulldoc"
    $pa = Invoke-RestMethod -Uri $pagesReq -Method Get -WebSession $session
    $numPages = $pa.pageList.pages.Count
    $numAllPages += $numPages
    $ps = $pa.pageList.pages | ForEach-Object {
      $pageId = $_.pageId
      
      "<pages><pageId>$pageId</pageId></pages>"
    }
    $jobParams = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
    <documentSelectionDescriptor>
        <docId>$docId</docId>
        <pageList>$ps</pageList>
    </documentSelectionDescriptor>"
    $jobRequest = "https://transkribus.eu/TrpServer/rest/recognition/$collection/$h/htrCITlab"

    try {
      $newJob = Invoke-RestMethod -Uri $jobRequest -Method Post -WebSession $session -Body $jobParams -ContentType application/xml
      $hJobs += $newJob
    } catch {
        $jobRequest
        $_.Exception.Response.StatusCode.value__
    }
  }

  $wait = 30 * $documents.Length
  $ts = New-TimeSpan -Seconds $wait
  "Waiting to allow text recognition to finish until " + ((Get-Date) + $ts)
  Start-Sleep -Seconds $wait

  $openJobs = Request-Status $hJobs $session
  do {
    $wait = 15 * $openJobs.length
    "Waiting for another $wait seconds to allow " + $openJobs.length + " TR jobs to finish"
    Start-Sleep -Seconds $wait
  } while ( $openJobs.length -gt 0 )
}

########################################################################################################################
##### PAGE Export and TEI-Konversion
if ( $e.isPresent ) {
  $exportPath = "export/$collection"
  New-Item -ItemType Directory -Force -Path $exportPath | Out-Null
  $counter = 0

  $documents | ForEach-Object {
    $counter++
    $docId = $_.docId
    $tempPath = "temp/$docId"

    # Info for progress bar
    $status = "Exporting file $counter of " + $documents.length + ": $docId"
    $pct = 100 * ($counter / $documents.Length)
    $act = "Exporting files…"
    
    New-Item -ItemType Directory -Force -Path $tempPath | Out-Null
    
    # get a list of pages
    $pagesReq = "https://transkribus.eu/TrpServer/rest/collections/$collection/$docId/fulldoc"
    Try {
      $pa = Invoke-RestMethod -Uri $pagesReq -Method Get -WebSession $session
    } Catch {
      $pagesReq
      $_
      Return
    }
    $numPages = $pa.pageList.pages.Count
  
    # last modification on server
    $modifiedServer = ($pa.pageList.pages.tagsStored | Measure-Object -Maximum).Maximum
  
    # only export if we do not have the file or the server version is newer
    if ( 
         (Test-Path $tempPath) -eq $False -or
         (Get-ChildItem $tempPath -Filter "*.xml").Length -eq 0 -or
         (Get-ChildItem $tempPath -Filter "*.xml" | Test-Path -OlderThan $modifiedServer)
    ) {
      $req = "https://transkribus.eu/TrpServer/rest/collections/$collection/$docId/export"
      $exportParams = @{
        "commonPars"= @{
          "pages"= "1-$numPages"
          "doExportDocMetadata"= "false"
          "doWriteMets"= "true"
          "doWriteImages"= "false"
          "doExportPageXml"= "true"
          "doExportAltoXml"= "false"
          "doWritePdf"= "false"
          "doWriteTei"= "true"
          "doWriteDocx"= "false"
          "doWriteTxt"= "false"
          "doWriteTagsXlsx"= "false"
          "doWriteTagsIob"= "false"
          "doWriteTablesXlsx"= "false"
          "doCreateTitle"= "false"
          "useVersionStatus"= "Latest version"
          "writeTextOnWordLevel"= "false"
          "doBlackening"= "false"
          "selectedTags"= @(
            "add",
            "date",
            "Address",
            "Antiqua",
            "supplied",
            "work",
            "unclear",
            "sic",
            "div",
            "regionType",
            "speech",
            "person",
            "gap",
            "organization",
            "comment",
            "abbrev",
            "place"
          )
          "font"= "FreeSerif"
          "splitIntoWordsInAltoXml"= "false"
          "pageDirName"= "page"
          "fileNamePattern"= "${filename}"
          "useHttps"= "true"
          "remoteImgQuality"= "orig"
          "doOverwrite"= "true"
          "useOcrMasterDir"= "true"
          "exportTranscriptMetadata"= "true"
        }
      } | ConvertTo-Json -Depth 3
      
      try {
        Write-Progress -Activity $act -CurrentOperation "triggering export" -PercentComplete $pct -Status $status
        $pa = Invoke-RestMethod -Uri $req -Method Post -WebSession $session -Body $exportParams -ContentType application/json
        
        $jobreq = "https://transkribus.eu/TrpServer/rest/jobs/$pa"
        $jo = Invoke-RestMethod -Uri $jobreq -Method Get -WebSession $session -Headers @{"Accept"="application/json"}
        
        Write-Progress -Activity $act -CurrentOperation "waiting for export to finish..." -PercentComplete $pct -Status $status
        while ($jo.state -ne "FINISHED") {
          Start-Sleep -Seconds 20 
          $jo = Invoke-RestMethod -Uri $jobreq -Method Get -WebSession $session -Headers @{"Accept"="application/json"}
        }
      
        $link = $jo.result
        Write-Progress -Activity $act -CurrentOperation "downloading from $link" -PercentComplete $pct -Status $status
        $global:progressPreference = 'SilentlyContinue'
        $get = Invoke-WebRequest -Uri $link -OutFile "temp/temp.zip" | Out-Null
        $global:progressPreference = 'Continue'
        
        Write-Progress -Activity $act -CurrentOperation "Expanding archive…" -PercentComplete $pct -Status $status
        # Force to overwrite log.txt...
        $get = Expand-Archive "temp/temp.zip" -DestinationPath temp -Force | Out-Null
        
        $name = (Get-ChildItem -Path $tempPath -Filter "*.xml").Name
        $xml = $tempPath + '/' + $name
        
        Write-Progress -Activity $act -CurrentOperation "getting or creating TEI" -PercentComplete $pct -Status $status
        # Check whether TEI file is > 0 Bytes; if = 0 Bytes, transform from PAGE
        If ((Get-Item $xml).length -eq 0kb) {
          $mets = $tempPath + "/" + $name.Substring(0, 10) + "/mets.xml"
          java -jar Saxon-HE-9.9.1-2.jar -xsl:page2tei-0.xsl -s:$mets -o:$xml
        }

        Copy-Item -Path $xml -Destination $exportPath
      } catch {
        "    Error downloading"
        $req
        $get
        $_
      }
    } Else {
      Write-Progress -Activity $act -CurrentOperation "already up-to-date (last modified on server: $modifiedServer)" -PercentComplete $pct -Status $status
    }
  }
}


$errors | ConvertTo-Json -Depth 3 | Out-File -FilePath errors.json

Get-Date -Format HH:mm:ss
$elapsedTime = $(get-date) - $startTime
"
Done processing $collectionID, this took"
"{0:HH:mm:ss}" -f ([datetime]$elapsedTime.Ticks)
