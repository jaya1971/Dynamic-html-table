<# 
.SYNOPSIS
  1. Monitor OnPrem certificates and alert systems team, through email, on any certificates that will expire
	within 30 days.

.INPUTS
  1. Modify the variable section.  Make sure to order the $columnList in the order you want the columns to appear in the email.
  2. Modify the "n=" in the "Select-Object" section of the certificate query to match column list names above and the way you want the headers to read on your table
  3. Modify any section with <> to match your environment.
   
.OUTPUTS
    1. Email with a table of certificates that will expire within 30 days.  

.NOTES
    Author:         Alex Jaya
    Creation Date:  02/18/2022
    Modified Date:  05/17/23
#>
#region Function section ################################
function Get-GraphToken {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$ClientID,
    [Parameter(Mandatory = $true)]
    [string]$ClientSecret,
    [Parameter(Mandatory = $true)]
    [string]$TenantID
  )
  $request = @{
    Method = 'POST'
    URI    = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
    body   = @{
      grant_type    = "client_credentials"
      scope         = "https://graph.microsoft.com/.default"
      client_id     = $ClientID
      client_secret = $ClientSecret
    }
  }
  try {
    # Get the access token
    $token = (Invoke-RestMethod @request).access_token
    return $token
  }
  catch {
    Write-Host "****** Failed to obtain a token ******" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.message)" -ForegroundColor Red
    return $false
  }
}
function Send-eMail {
  Param(
    [Parameter(Mandatory = $true)][string] $Sender,
    [Parameter(Mandatory = $true)][string] $UserEmail,
    [Parameter(Mandatory = $true)][string] $Subject,
    [Parameter(Mandatory = $true)][string] $EmailBody,
    [Parameter(Mandatory = $true)][string] $tenantid,
    [Parameter(Mandatory = $true)][string] $clientid,
    [Parameter(Mandatory = $true)][string] $clientsecret
  )

  # Get the access token
  $tokenParams = @{
    clientId = $clientId
    clientSecret = $clientSecret
    tenantId= $tenantId
  }
  $token = Get-GraphToken @tokenParams
  if($token){
    # DO NOT CHANGE ANYTHING BELOW THIS LINE
    # Build the Microsoft Graph API request
    $params = @{
      "URI" = "https://graph.microsoft.com/v1.0/users/$Sender/sendMail"
      "Headers" = @{
        "Authorization" = ("Bearer {0}" -F $token)
      }
      "Method" = "POST"
      "ContentType" = 'application/json'
      "Body" = (@{
        "message" = @{
          "subject" = $Subject
          "body" = @{
            "contentType" = 'HTML'
            "content" = $EmailBody
          }
          "toRecipients" = @(
            @{
              "emailAddress" = @{
                "address" = $UserEmail
              }
            }
          )
        }
        SaveToSentItems = "false"
      }) | ConvertTo-JSON -Depth 10
    }

    # Send the message
    Invoke-RestMethod @params -Verbose
  }
}
#endregion Function section ################################

#region Variable section ################################
$clientID = <CLIENTID>
$ClientSecret = <CLIENTSECRET>
$tenantid= <TENANTID>
$mailTo = <GroupEmail>
$subject = "OnPrem Expired Certificates"
$TableName = "OnPrem Expired Certficates"
$ColumnList = @('Computer Name','Friendly Name','Thumbprint','Subject','Expiration Date')
$Sortby = "Computer Name"
#endregion Variable section ################################
Import-Module Active-Directory
Import-Module .\CreateTable -Force
$SYScred = Get-Credential

#Get all servers in the domain
$Servers = Get-ADComputer -Filter "OperatingSystem -like 'Windows Server*'" -SearchBase "OU=Domain Member Servers,DC=<DOMAIN>,DC=com"
$count = 0
$Certificates = @()
$Certificates += $servers | ForEach-Object{
  $percentage = ($count++/$servers.count)*100
  $percentage = [math]::Floor($percentage)
  Write-Progress -Activity "Checking $_" -Status "$percentage % completed" -PercentComplete $percentage

  if(Test-Connection -ComputerName $_.DNSHostName -Count 2 -Quiet){
      Invoke-Command -ComputerName $_.DNSHostName -Credential $SYSCred -ErrorAction SilentlyContinue {Get-ChildItem -Path Cert:\LocalMachine\my | Where-Object  {$_.NotAfter -lt (Get-Date).AddDays(30) -and `
          ($_.Subject -Match "<SearchString1>|<SearchString2>|<SearchString3>")} | Select-Object @{n="Computer Name";e={$env:COMPUTERNAME}}, @{n="Friendly Name";e={$_.friendlyName}},Thumbprint,Subject,@{n="Expiration Date";e={$_.NotAfter}}}
  }
}

if($Certificates)
{
  $Table = New-Table -TableName $TableName -tblArray $Certificates -ColumnArray $ColumnList

  ###############
  # Build Email #
  ###############
  # Creating head style
  $Head = @"
    
  <style>
    body {
      font-family: "Arial";
      font-size: 8pt;
      color: #4C607B;
      }
    th, td { 
      border: 1px solid #e57300;
      border-collapse: collapse;
      padding: 5px;
      }
    th {
      font-size: 1.2em;
      text-align: left;
      background-color: #003366;
      color: #ffffff;
      }
    td {
      color: #000000;
      }
    .even { background-color: #ffffff; }
    .odd { background-color: #bfbfbf; }
  </style>    
"@
  [string]$tblBody = [PSCustomObject]$Table | Select-Object -Property $ColumnList | Sort-Object -Property $Sortby | ConvertTo-Html `
  -Head $Head -Body "<font color =`"Black`"><h4>OnPrem Certificate Expiration Report</h4></font>"

  $MailArguments = @{
    Sender = $mailTo
    UserEmail = $mailTo 
    Subject = $subject 
    EmailBody = $tblBody
    tenantid = $tenantid
    clientid = $clientid
    clientsecret = $clientsecret
  }
  Try
  {    
    Send-eMail @MailArguments
    Write-Output "Mail has been sent!"
  }
  Catch
  {
    $ExcMessage = $_.Exception.Message
    throw "Error: Can not send email!. Exception: $ExcMessage"
  }
  Clear-Variable Certificates
}