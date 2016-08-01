function ConvertTo-URLEncode{
  [cmdletbinding()]
  PARAM
  (
    $String
  )
  Add-Type -AssemblyName System.Web
  [system.web.httputility]::urlencode($String)
}
 
function Invoke-BingWebRequest {
  [cmdletbinding()]
  PARAM
  (
    $uri,
    $body,
    $Access_Token
  )
  Write-Debug $uri
  Write-Verbose $uri
  Try {
    if ($Access_Token) { 
      $HeaderValue = "Bearer $Access_Token"
      $result = Invoke-RestMethod -Uri $uri -Headers @{Authorization = $HeaderValue} -body $body -ErrorAction SilentlyContinue -SessionVariable SessionError -ErrorVariable Caught 
      $result  | add-member StatusCode 200
    } Else {
      $result = Invoke-RestMethod -Uri $uri -body $body -ErrorAction SilentlyContinue -SessionVariable SessionError -ErrorVariable Caught
      $result  | add-member StatusCode 200
    }
  } Catch [System.Net.WebException],[System.IO.IOException] {
    $result = '' | Select-Object StatusCode,Method,Message
    $result.StatusCode = 404
    $result.Message = $caught.Message
  } Catch {
    $result = '' | Select-Object StatusCode,Method,Message
    $result.StatusCode = 503
    $result.Message = ($caught.Message.split("`n") | Select-String -SimpleMatch 'Message:').tostring().replace('Message: ','')
    $result.Method = ($caught.Message.split("`n") | Select-String -SimpleMatch 'Method:').tostring().replace('Method: ','') 
      
  }

  $result

}

function Get-BingMTAccessToken {
  [cmdletbinding()]
  PARAM
  (
    $ClientID,
    $Client_Secret
  )
  $uri = 'https://datamarket.accesscontrol.windows.net/v2/OAuth2-13'

  # If ClientId or Client_Secret has special characters, UrlEncode before sending request
  $Body = 'grant_type=client_credentials&client_id=' + $(convertto-UrlEncode($ClientID)) 
  $Body += '&client_secret=' + $(convertto-UrlEncode($Client_Secret)) + '&scope=http://api.microsofttranslator.com'

  $admAuth = Invoke-RestMethod -Uri $uri -Body ($Body) -ContentType 'application/x-www-form-urlencoded' -Method Post

  #$admAuth = ($token.Content | ConvertFrom-Json)
  $admAuth.access_token
  Write-Verbose $admAuth.access_token
  Write-Debug $admAuth.access_token

}

Function Invoke-GetLanguagesForTranslate{
  [cmdletbinding()]
  PARAM
  (
    $AppID,
    $Server='api.microsofttranslator.com',
    [switch]$Secure,
  $Access_Token)

  $uri = 'http'
  if ($Secure) {$uri = 'https'}
  $uri += "://$Server/v2/Http.svc/GetLanguagesForTranslate"
  if ($AppID) { $uri += "?AppId=$AppId"}

  Invoke-BingWebRequest -uri $uri -Access_Token $Access_Token
}

Function Invoke-TranslateMethod{
  [cmdletbinding()]
  PARAM
  (
    $AppId,
    $Server='api.microsofttranslator.com',
    $text,
    $from,
    $to,
    [switch]$Secure,
    $Access_Token
  )
  PROCESS {
    $encodedText = ConvertTo-URLEncode($text)
    $uri = 'http'
    if ($Secure) {$uri = 'https'}
    $uri += "://$Server/v2/Http.svc/Translate?"
    if ($AppId) {$uri += "AppId=$AppId&" }
    $uri += "text=$encodedText&from=$from&to=$to"

    Invoke-BingWebRequest -uri $uri -Access_Token $Access_Token

  }
}

Function Invoke-DetectLanguage{
  [cmdletbinding()]
  PARAM
  (
    $AppId,
    $Server='api.microsofttranslator.com',
    $text,
    [switch]$Secure,
    $Access_Token
  )
  $encodedText = ConvertTo-URLEncode($text)
  $uri = 'http'
  if ($Secure) {$uri = 'https'}
  $uri += "://$Server/v2/Http.svc/Detect?"
  if ($AppId) {$uri += "AppId=$AppId&" }

  $uri += "Text=$EncodedText"  

  Invoke-BingWebRequest -uri $uri -Access_Token $Access_Token
} 

Function Invoke-DetectArrayLanguage{
  [cmdletbinding()]
  PARAM
  (
    $AppId,
    $Server='api.microsofttranslator.com',
    $text,
    [switch]$Secure,
    $Access_Token
  )
  $encodedText = ConvertTo-URLEncode($text)
  $uri = "http://$Server/V2/Http.svc/DetectArray?AppId=$AppId"

  Invoke-BingWebRequest -uri $uri -Access_Token $Access_Token
}

Function Invoke-GetLanguageNames{
  [cmdletbinding()]
  PARAM
  (
    $AppId,
    $Server='api.microsofttranslator.com',
    $locale,
    [switch]$Secure,
    $Access_Token
  )

  $uri = 'http'
  if ($Secure) {$uri = 'https'}
  $uri += "://$Server/v1/Http.svc/GetLanguageNames?"
  if ($AppId) {$uri += "AppId=$AppId&" }

  $uri += "locale=$locale"

  Write-Verbose $uri

  Invoke-BingWebRequest -uri $uri -Access_Token $Access_Token
}

Function Invoke-GetLanguagesForSpeak{
  [cmdletbinding()]
  PARAM
  (
    $AppId,
    $Server='api.microsofttranslator.com',
    [switch]$Secure,
    $Access_Token
  )
  $uri = 'http'
  if ($Secure) {$uri = 'https'}
  $uri += "://$Server/v2/Http.svc/GetLanguagesForSpeak?"
  if ($AppId) {$uri += "AppId=$AppId" }

  Invoke-BingWebRequest -uri $uri -Access_Token $Access_Token
}

Function Invoke-GetTranslationsArray{
  [cmdletbinding()]
  PARAM
  (
    $AppId,
    $Server='api.microsofttranslator.com',
    [switch]$Secure,
    $Access_Token
  )
  $encodedText = ConvertTo-URLEncode($text)
  $uri = 'http'
  if ($Secure) {$uri = 'https'}
  $uri += "://$Server/v2/Http.svc/GetTranslationsArray?"
  if ($AppId) {$uri += "AppId=$AppId" }

  Invoke-BingWebRequest -uri $uri -Access_Token $Access_Token
}

Function Invoke-Speak{
  [cmdletbinding()]
  PARAM
  (
    $AppId,
    $Server='api.microsofttranslator.com',
    $text,
    $language,
    [switch]$Secure,
    $Access_Token
  )
  $encodedText = ConvertTo-URLEncode($text)

  $uri = 'http'
  if ($Secure) {$uri = 'https'}
  $uri += "://$Server/v2/Http.svc/Speak?"
  if ($AppId) {$uri += "AppId=$AppId&" }

  $uri += "text=$encodedText&language=$language&format=" + (ConvertTo-URLEncode('audio/wav')) + '&options=MaxQuality' 

  Invoke-BingWebRequest -uri $uri -Access_Token $Access_Token
}

Function Invoke-TranslateArray{
  [cmdletbinding()]
  PARAM
  (
    $AppId,
    $Server='api.microsofttranslator.com',
    [switch]$Secure,
    $Access_Token
  )
  $encodedText = ConvertTo-URLEncode($text)

  $uri = 'http'
  if ($Secure) {$uri = 'https'}
  $uri += "://$Server/v2/Http.svc/TranslateArray"
  if ($AppId) {$uri += "AppId=$AppId&" }

  $ContentType = 'text/plain'
  [string] $Body = '<TranslateArrayRequest>' +
  '<AppId />' +
  "<From>$from</From>" +
  '<Options>' +
  '<Category xmlns="http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2" />' +
  "<ContentType xmlns=""http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2"">$ContentType</ContentType>" +
  '<ReservedFlags xmlns="http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2" />' +
  '<State xmlns="http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2" />' +
  '<Uri xmlns="http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2" />' +
  '<User xmlns="http://schemas.datacontract.org/2004/07/Microsoft.MT.Web.Service.V2" />' +
  '</Options>' +
  '<Texts>' +
  '<string xmlns="http://schemas.microsoft.com/2003/10/Serialization/Arrays">This is string 1</string>' +
  '<string xmlns="http://schemas.microsoft.com/2003/10/Serialization/Arrays">This is string 1</string>' +
  '<string xmlns="http://schemas.microsoft.com/2003/10/Serialization/Arrays">This is string 1</string>' +
  '</Texts>' +
  "<To>$to</To>" +
  '</TranslateArrayRequest>';


  Invoke-BingWebRequest -uri $uri -body $Body -Access_Token $Access_Token
}

Function Invoke-BreakSentences{
  [cmdletbinding()]
  PARAM
  (
    $AppId,
    $Server='api.microsofttranslator.com',
    $text,
    $Language,        
    [switch]$Secure,
    $Access_Token
  )

  $encodedText = ConvertTo-URLEncode($text)

  $uri = 'http'
  if ($Secure) {$uri = 'https'}
  $uri += "://$Server/v2/Http.svc/BreakSentences?"
  if ($AppId) {$uri += "AppId=$AppId&" }

  $uri += "Text=$encodedText&Language=$Language"

  Invoke-BingWebRequest -uri $uri -Access_Token $Access_Token
}

Function Invoke-TransformTextMethod{
  [cmdletbinding()]
  PARAM
  (
    $AppId,
    $Server='api.microsofttranslator.com',
    $Sentence,
    $Language='en',
    [switch]$Secure,
    $Access_Token
  )
  PROCESS {
    $encodedText = ConvertTo-URLEncode($Sentence)
    $uri = 'http'
    if ($Secure) {$uri = 'https'}
    $uri += "://$Server/v3/json/TransformText?"
    if ($AppId) {$uri += "AppId=$AppId&" }
    $uri += "sentence=$encodedText&language=$Language"

    $json = Invoke-BingWebRequest -uri $uri -Access_Token $Access_Token
    if ($json.StatusCode -ne 404) {
      $json = $json.Substring(1,$json.Length-1)
      ConvertFrom-Json $json
    } else {
      New-Object psobject -Property @{ec=404;em='Not Responding'}
    }


  }
}
