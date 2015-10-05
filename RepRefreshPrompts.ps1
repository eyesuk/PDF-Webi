#
# Retrieves a specified Web Intelligence document as PDF.
#
############################################################################################################################
# Input: $logonInfo, $hostUrl, $locale, $documentId and $folderPath to suite your preferences
$logonInfo = @{}
$logonInfo.userName = "youruser"           #user name
$logonInfo.password = "yourpassword"       #password 
$logonInfo.auth     = "AuthType"           #secTypes = secEnterprise, secLDAP, secWinAD, secSAPR3  
$hostUrl = "http://YOURURL:YOURPORT/biprws/" # your environment RESTful services.

# Mapping prompt to prompt value. Change the labels with the actual prompt names (PromptName) from the report and wanted values (StringVal/NumericVal) 
$promptValueMap = @{"PromptName1" = @("StrsingVal1") ;
                    "PromptName2" = @("StrsingVal2") ;
                    "PromptName3" = NumericVal ;
                    "PromptNameNTH" = NumericVal }
$documentId = 2316769  # SI_ID for the document NOT CUID
$locale = "en-US"         # Product Language
$contentLocale = "en-US"  # Document Content Formatting Language

# Folder where PDF file will be saved.  File name will be report_name.pdf
$folderPath = "Y:\Downloads" 
############################################################################################################################

# Logon and retrieve Logon Token and add to the HTTP Header we send out on subsequent calls.
$headers = @{"Accept"       = "application/json" ; 
             "Content-Type" = "application/json"
              }
$result = Invoke-RestMethod -Method Post -Uri ($hostUrl + "/logon/long") -Headers $headers -Body (ConvertTo-Json($logonInfo))
$logonToken =  "`"" + $result.logonToken + "`""  # The logon token must be delimited by double-quotes.

# Get document information and use the document name as file name.
$headers = @{ "X-SAP-LogonToken" = $logonToken ;
              "Accept"           = "application/json" ;
              "Content-Type"     = "application/json" ;
              "Accept-Language"  = $locale    ;
              "X-SAP-PVL" = $contentLocale  
           }
$documentUrl = $hostUrl + "/raylight/v1/documents/" + $documentId 
$result = Invoke-RestMethod -Method Get -Uri $documentUrl -Headers $headers
$document = $result.document

# Retrieve and set parameters.
$headers = @{ "X-SAP-LogonToken" = $logonToken ;
              "Accept"           = "application/json" ;
              "Content-Type"     = "application/json" ;
              "X-SAP-PVL" = $contentLocale  
              }

# Retrieve the parameters to specify.
$parametersUrl = $documentUrl + "/parameters"
$parametersResult = Invoke-RestMethod -Method Get -Uri $parametersUrl -Headers $headers

# Given each parameter in the parameters collection, specify the parameter value accorinding to the promptValueMap. 
$parametersResult.parameters.parameter | Where-Object {$_ -ne $null} | ForEach-Object {
    $promptValue = $promptValueMap[$_.name]
    if($promptValue){
        $_.answer.values.value = $promptValue
    }
}
$headers = @{ "X-SAP-LogonToken" = $logonToken ;
              "Accept"           = "application/json" ;
              "Content-Type"     = "application/json"   ;
              "X-SAP-PVL" = $contentLocale  
              } 
$result = Invoke-RestMethod -Method Put -Uri $parametersUrl -Headers $headers -Body (ConvertTo-Json $parametersResult -Depth 10)

# Get PDF and save to file (only if the file path is valid)
$filePath = $folderPath + "/" + $document.name + ".pdf"
if(Test-Path $filePath -isValid) {
    # Get PDF and save to file
    $headers = @{ "X-SAP-LogonToken" = $logonToken ;
                  "Accept"           = "application/pdf";
                  "X-SAP-PVL" = $contentLocale  
               }
               
    Invoke-RestMethod -Method Get -Uri ($documentUrl + "/pages") -Headers $headers -OutFile $filePath
} else {
    Write-Error "Invalid file path " + $filePath
}

# Unload document from Raylight
$headers = @{ "X-SAP-LogonToken" = $logonToken ;
              "Accept"           = "application/json" ;
              "Content-Type"     = "application/json" ;
              "X-SAP-PVL"        = $contentLocale
              }
$result = Invoke-RestMethod -Method Put -Uri $documentUrl -Headers $headers -Body (ConvertTo-Json(@{"document"=@{"state"="Unused"}}))

# Log off the Session identified by the X-SAP-LogonToken HTTP Header
$headers = @{ "X-SAP-LogonToken" = $logonToken ;
              "Accept"           = "application/json" ;
              "Content-Type"     = "application/json" }
$logoffUrl = $hostUrl + "/logoff"
Invoke-RestMethod -Method Post -Uri $logoffUrl -Headers $headers