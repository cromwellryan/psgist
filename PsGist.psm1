. (join-path $PSScriptRoot "/json 1.7.ps1")

function Get-Github-Credential($username) {
	$host.ui.PromptForCredential("Github Credential", "Please enter your Github user name and password.", $username, "")
}

function Clean-Username($username) {
  $username.Replace("\", "")
}

function New-DiffGist { 
<# 
	.Synopsis
	Publishes Github Gists of current git diff.

	.Description
	Publishes files as Owned or Anonymous Github Gists.

    .Parameter Name
    The name to use for the filename of the Gist (minus .diff)

	.Parameter Description 
	(optional) The Description of this Gist.

	.Parameter Username
	The Github username which will own this Gist.

	.Parameter Private
	When specified, the Gist will be made private.  Default is private.

    .Parameter Launch
    When specified, the default browser will launch to the created Gist URI

	.Example
	diffgist -Description "Hello.js greets all visitors"
	Publishing a private Gist

	.Example
	diffgist -Name "Hello" -Description "Hello.js greets all visitors" -Public
	Publishing a public Gist
#>
	Param(
		[Parameter(Position=0, ValueFromPipeline=$true)]
        [string]$Name = "current.diff",
		[string]$Description = "",
		[string]$Username = $null,
		[switch]$Public = $false,
        [switch]$Launch = $false
	)
	BEGIN {
		$files = @{}
	}
	PROCESS {

        $content = git diff | Out-String

		$content = $content -replace "\\", "\\\\"
		$content = $content -replace "`t", "\t"
		$content = $content -replace "`r", "\r"
		$content = $content -replace "`n", "\n"
		$content = $content -replace """", "\"""
		$content = $content -replace "/", "\/"

		$files.Add($Name, $content)
	}
	END {

		$apiurl = "https://api.github.com/gists"

		$request = [Net.WebRequest]::Create($apiurl)

        $credential = $(Get-Github-Credential $Username)
	
		if($credential -eq $null) {
			write-host "Github credentials are required."
			return
		}

		$username = (Clean-Username $credential.Username)
		$password = $credential.Password

		$bstrpassword= [Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
		$insecurepassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstrpassword)

		$basiccredential = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([String]::Format("{0}:{1}", $username, $insecurepassword)))
		$request.Headers.Add("Authorization", "Basic " + $basiccredential)

		$request.ContentType = "application/json"
		$request.Method = "POST"

		$files.GetEnumerator() | % { 
			$singlefilejson = """" + $_.Name + """: {
					""content"": """ + $_.Value + """
			},"
	
			$filesjson += $singlefilejson
		}

		$filesjson = $filesjson.TrimEnd(',')

		$ispublic = $Public.ToString().ToLower()
		
		$body = "{
			""description"": """ + $Description + """,
			""public"": $ispublic,
			""files"": {" + $filesjson + "}
		}"

		$bytes = [text.encoding]::Default.getbytes($body)
		$request.ContentLength = $bytes.Length

		$stream = [io.stream]$request.GetRequestStream()
		$stream.Write($bytes,0,$bytes.Length)

		try {
			$response = $request.GetResponse()
		}
		catch  [System.Net.WebException] {
			$_.Exception.Message | write-error 
			
			return
		}
		
		$responseStream = $response.GetResponseStream()
		$reader = New-Object system.io.streamreader -ArgumentList $responseStream
		$content = $reader.ReadToEnd()
		$reader.close()

		if( $response.StatusCode -ne [Net.HttpStatusCode]::Created ) {
			$content | write-error

			return
		}

		$result = convertfrom-json $content -Type PSObject

		$url = $result.html_url
	
		write-output $url
        
        if ($Launch) {
            start $url
        }
	}
}

function New-Gist { 
<# 
	.Synopsis
	Publishes Github Gists.

	.Description
	Publishes files as Owned or Anonymous Github Gists.

	.Parameter InputObject
	Accepts a series of files which will be published as a single Gist

	.Parameter File
	A single file path to be published as a Gist.

	.Parameter Description 
	(optional) The Description of this Gist.

	.Parameter Username
	The Github username which will own this Gist.

	.Parameter Public
	When specified, the Gist will be made public.  Default is private.

	.Example
	gist -File "Hello.js" -Description "Hello.js greets all visitors"
	Publishing a public Gist

	.Example
	gist -File "Hello.js" -Description "Hello.js greets all visitors" -Public
	Publishing a private Gist
#>
	Param(
		[Parameter(Position=0, ValueFromPipeline=$true)]
		[PSObject]$InputObject = $null,
		[string]$File = $null,
		[string]$Description = "",
		[string]$Username = $null,
		[switch]$Public = $false
	)
	BEGIN {
		$files = @{}
	}
	PROCESS {
		if( $InputObject -ne $null ) {
			if( $InputObject.GetType() -eq [System.IO.FileInfo] ) {
				$fileinfo = [System.IO.FileInfo]$InputObject
			} 
			else { # Ignore Directories
				return 
			}
		}
		elseif( $File -ne $null -and (Test-Path $File) ){
			$fileinfo = Get-Item $File
		}
		else {
			return
		}

		$path = $fileinfo.FullName
		$filename = $fileinfo.Name

		$content = [IO.File]::readalltext($path)

		$content = $content -replace "\\", "\\\\"
		$content = $content -replace "`t", "\t"
		$content = $content -replace "`r", "\r"
		$content = $content -replace "`n", "\n"
		$content = $content -replace """", "\"""
		$content = $content -replace "/", "\/"
		#$content = $content -replace "'", "\'"

		$files.Add($filename, $content)
	}
	END {

		$apiurl = "https://api.github.com/gists"

		$request = [Net.WebRequest]::Create($apiurl)

		$credential = $(Get-Github-Credential $Username)
	
		if($credential -eq $null) {
			write-host "Github credentials are required."
			return
		}

		$username = (Clean-Username $credential.Username)
		$password = $credential.Password

		$bstrpassword= [Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
		$insecurepassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstrpassword)

		$basiccredential = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([String]::Format("{0}:{1}", $username, $insecurepassword)))
		$request.Headers.Add("Authorization", "Basic " + $basiccredential)

		$request.ContentType = "application/json"
		$request.Method = "POST"

		$files.GetEnumerator() | % { 
			$singlefilejson = """" + $_.Name + """: {
					""content"": """ + $_.Value + """
			},"
	
			$filesjson += $singlefilejson
		}

		$filesjson = $filesjson.TrimEnd(',')

		$ispublic = $Public.ToString().ToLower()
		
		$body = "{
			""description"": """ + $Description + """,
			""public"": $ispublic,
			""files"": {" + $filesjson + "}
		}"

		$bytes = [text.encoding]::Default.getbytes($body)
		$request.ContentLength = $bytes.Length

		$stream = [io.stream]$request.GetRequestStream()
		$stream.Write($bytes,0,$bytes.Length)

		try {
			$response = $request.GetResponse()
		}
		catch  [System.Net.WebException] {
			$_.Exception.Message | write-error 
			
			return
		}
		
		$responseStream = $response.GetResponseStream()
		$reader = New-Object system.io.streamreader -ArgumentList $responseStream
		$content = $reader.ReadToEnd()
		$reader.close()

		if( $response.StatusCode -ne [Net.HttpStatusCode]::Created ) {
			$content | write-error

			return
		}

		$result = convertfrom-json $content -Type PSObject

		$url = $result.html_url
	
		write-output $url
	}
}

new-alias gist New-Gist
new-alias Create-Gist New-Gist # For those who saw my post when I used create... to be deprecated
new-alias diffgist New-DiffGist

export-modulemember -alias * -function New-Gist
export-modulemember -alias * -function New-DiffGist
