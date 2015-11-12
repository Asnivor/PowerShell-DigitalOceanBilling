####################################################################################################################################
#--------------------------------------------------------[Script Information]------------------------------------------------------#
####################################################################################################################################
<#
.SYNOPSIS
  			DigitalOcean Invoice Download Script
 
.DESCRIPTION
  			This script connects logs into the DigitalOcean website with your details and downloads all invoice PDFs to a folder that
			you specify (skipping downloads if the file already exists).
			
			This is a small snippet of a larger script I developed that pulls all billing data (not just the PDFs) and puts this into
			MSSQL Server for display on a company intranet.
	 
.OUTPUTS
  			Invoice PDFs in the directory that you specify
 
.NOTES
  			Version:    	1.1    
			Author:         Asnivor ( coding@asnitech.co.uk )
			Creation Date:  12/11/2015
			Purpose/Change: Script creation
		
.CREDITS	
			1.	Script uses a slightly modified version of Joel Bennett's 'ConvertFrom-HTML' function: http://poshcode.org/4850
  
.EXAMPLE
			DigitalOceanBilling.ps1
#>


#################################################################################################################################### 
#----------------------------------------------------------[Declarations]----------------------------------------------------------#
####################################################################################################################################

<# 
#	Configure your variables below 	
#>

# Digital Ocean Account Credentials
$username = "someguy@somedomain.com"
$password = "your-password"

# Local folder to store downloaded invoices
$savePath = "c:\your\downloaded\invoices\location"


#################################################################################################################################### 
#-------------------------------------------------------------[Setup]--------------------------------------------------------------#
####################################################################################################################################

# Get path to this script (works with both standard PowerShell and PowerGUI
$scriptPath = Split-Path (Resolve-Path $myInvocation.MyCommand.Path)
if ($scriptPath -match 'Quest Software') {$scriptPath = [System.AppDomain]::CurrentDomain.BaseDirectory}

#################################################################################################################################### 
#--------------------------------------------------------[Generic Functions]-------------------------------------------------------#
####################################################################################################################################
 
function ConvertFrom-Html {
   #.Synopsis
   #   Convert a table from an HTML document to a PSObject
   #.Example
   #   Get-ChildItem | Where { !$_.PSIsContainer } | ConvertTo-Html | ConvertFrom-Html -TypeName Deserialized.System.IO.FileInfo
   #   Demonstrates round-triping files through HTML
   param(
      # The HTML content
      [Parameter(ValueFromPipeline=$true)]
      [string]$html,

      # A TypeName to inject to PSTypeNames 
      [string]$TypeName
   )
   begin 
   { 
   	$content = "$html" 
   }
   process { $content += "$html" }
   end {
      #
	  [xml]$table = $content -replace '(?s).*<table[^>]*>(.*)</table>.*','<table>$1</table>'
		#$save = $table.Save("$scriptPath\test.txt")
		if ($table.table.tr.Count -eq 0)
		{
			Write-Host "ERROR: Either the wrong user/pass combo was used, or the script just didn't return any data"
			Write-Host "This script will now terminate"
			exit
		}
      $header = $table.table.tr[0]  
      $data = $table.table.tr[1..1e3]

      foreach($row in $data){ 
         $item = @{}
         $h = "th"
         if(!$header.th) {
            $h = "td"
         }
         for($i=0; $i -lt $header.($h).Count; $i++){
            if($header.($h)[$i] -is [string]) {
			  $item.($header.($h)[$i]) = $row.td[$i]
            } else {

               $item.($header.($h)[$i].InnerText) = $row.td[$i]
            }
         }
         Write-Verbose ($item | Out-String)
         $object = New-Object PSCustomObject -Property $item 
         if($TypeName) {
            $Object.PSTypeNames.Insert(0,$TypeName)
         }
		 #$save2 = $object | out-file "$scriptPath\object.txt" -Append
         Write-Output $Object
		
      }
	   If($?)
		{			
		 
    	}
   }
}

#################################################################################################################################### 
#-----------------------------------------------------------[Execution]------------------------------------------------------------#
####################################################################################################################################
cls
$baseUrl = "https://cloud.digitalocean.com"		
	
# Code to ignore certificate errors
try
{
	add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    
    public class IDontCarePolicy : ICertificatePolicy {
        public IDontCarePolicy() {}
        public bool CheckValidationResult(
            ServicePoint sPoint, X509Certificate cert,
            WebRequest wRequest, int certProb) {
            return true;
        }
    }
"@
}
catch {}
[System.Net.ServicePointManager]::CertificatePolicy = new-object IDontCarePolicy 
$user_agent = "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.2; WOW64; Trident/6.0; MALNJS)"
			
# Connect to DigitalOcean Web Portal
Write-Host "Connecting to the Digital Ocean Web Portal"
$initialRequest = Invoke-WebRequest -Uri 'https://cloud.digitalocean.com/settings/billing' -SessionVariable WebSession1 -TimeoutSec 0 -MaximumRedirection 50 -UserAgent $user_agent

# Get the form fields required
$form = $initialRequest.Forms[0]

# Create an array of parameters to send as POST to the website
$postParams = @{
	utf8="&E2%9C%93"
	authenticity_token=$form.Fields.authenticity_token
	'user[email]'=$username
	'user[password]'=$password
	commit="Log In"
}

# Log in to the portal
Write-Host "Authenticating with the Digital Ocean Web Portal"
$loginRequest = Invoke-WebRequest -Uri ('https://cloud.digitalocean.com' + $form.Action) -Method Post -Body $postParams -TimeoutSec 0 -WebSession $WebSession1 -UserAgent $user_agent

# capture just the 'billing history' table from the webpage into the string '$billing'
$invoicePage = $loginRequest -split "`n"
$billing = ""
$found = 0
foreach ($line in $invoicePage)
{
	if ($line -match "listing Billing--histroy")
	{
		# data we want starts here
		$found = 1
	}
	if ($line -match "view_history")
	{
		# data we want has finished
		$found = 0
	}	
	# log data that we want
	if ($found -eq 1)
	{
		$billing += $line
	}				
}

# Clean up the HTML so that it can be parsed to the ConvertFrom-HTML function
$stringsToRemove = @("<thead>","</thead>","<tbody>","</tbody>"," class='listing Billing--histroy'","class style=`"display: table-row;`""," target=`"_blank`"","class='hidden'"
					"<a href=`"/billing/","`">View Invoice</a>")
foreach ($thing in $stringsToRemove)
{
	$billing = $billing.Replace($thing, "")
}
$billing = $billing.Replace("<th></th>", "<th>Files</th>").Replace("< tr>", "<tr>").Replace("<tr >", "<tr>")

# Convert HTML table into PowerShell Object
$htmlObject = ConvertFrom-Html -html $billing

# Download each invoice that is found

$noInvoice = 0
$localInvoice = 0
$downloadInvoice = 0

foreach ($item in $htmlObject)
{
	if ($item.Files.Length -gt 0)
	{
		# Check whether file does not exist locally
		if (!(Test-Path -Path "$($savePath)\$($item.Files.Trim())" -PathType Leaf))
		{
			# File does not exist locally - download it
			Write-Host "Downloading Invoice: $($item.Files.Trim())"
			$download = Invoke-WebRequest "https://cloud.digitalocean.com/billing/$($item.Files.Trim())" -TimeoutSec 0 -Method Get -WebSession $WebSession1 -UserAgent $user_agent -OutFile "$($savePath)\$($item.Files.Trim())"
			$downloadInvoice++
		}
		else
		{
			Write-Host "Invoice $($item.Files.Trim()) already exists locally - skipping download"
			$localInvoice++
		}		
	}
	else
	{
		$noInvoice++
	}
}

Write-Host "Script Completed"
Write-Host "$($downloadInvoice) Invoices downloaded"
Write-Host "$($localInvoice) Invoices skipped"
Write-Host "$($noInvoice) billing items detected had no downloads associated"

