#requires -version 2.0
# Version History:
# v 0.5 - First Public version
# v 1.0 - Made ConvertFrom-Json work with arbitrary JSON 
#       - switched to xsl style sheets for ConvertTo-JSON
# v 1.1 - Changed ConvertFrom-Json to handle single item results
# v 1.2 - CodeSigned to make a fellow geek happy
# v 1.3 - Changed ConvertFrom-Json to handle zero item results (I hope)
# v 1.4 - Added -File parmeter set to ConvertFrom-Json
#       - Cleaned up some error messages
# v 1.5 - Corrected handling of arrays
# v 1.6 - Corrected pipeline binding on ConvertFrom-Json
# v 1.7 - Added a New-Json which converts simple hashtables...
#
#  There is no help (yet) because I keep forgetting that I haven't written help yet.
#  Full RoundTrip capability:
#
#  > ls | ConvertTo-Json | ConvertFrom-Json
#  > ps | ConvertTo-Json | Convert-JsonToXml | Convert-XmlToJson | convertFrom-Json
#
#  You may frequently want to use the DEPTH or NoTypeInformation parameters:
#
#  > ConvertTo-Json -Depth 2 -NoTypeInformation
#
#  But then you have to specify the type when you reimport (and you can't do that for deep objects).  
#  This problem also occurs if you convert the result of a SELECT statement (ie: PSCustomObject).
#  For Example:
#
#  > PS | Select PM, WS, CPU, ID, ProcessName |
#  >> ConvertTo-json -NoType |
#  >> convertfrom-json -Type System.Diagnostics.Process
#
#  However, you *can* use PSOjbect as your type when re-importing:
#
#  > $Json = Get-Process | 
#  >> Select PM, WS, CPU, ID, ProcessName, @{n="SnapshotTime";e={Get-Date}} | 
#  >> ConvertTo-Json -NoType 
#  
#  > $Json | ConvertFrom-json -Type PSObject


Add-Type -AssemblyName System.ServiceModel.Web, System.Runtime.Serialization
$utf8 = [System.Text.Encoding]::UTF8

function Write-Stream {
PARAM(
   [Parameter(Position=0)]$stream,
   [Parameter(ValueFromPipeline=$true)]$string
)
PROCESS {
  $bytes = $utf8.GetBytes($string)
  $stream.Write( $bytes, 0, $bytes.Length )
}  
}


function Read-Stream {
PARAM(
   [Parameter(Position=0,ValueFromPipeline=$true)]$Stream
)
process {
   $bytes = $Stream.ToArray()
   [System.Text.Encoding]::UTF8.GetString($bytes,0,$bytes.Length)
}}


function Convert-JsonToXml {
PARAM([Parameter(ValueFromPipeline=$true)][string[]]$json)
BEGIN { 
   $mStream = New-Object System.IO.MemoryStream 
}
PROCESS {
   $json | Write-Stream -Stream $mStream
}
END {
   $mStream.Position = 0
   try
   {
      $jsonReader = [System.Runtime.Serialization.Json.JsonReaderWriterFactory]::CreateJsonReader($mStream,[System.Xml.XmlDictionaryReaderQuotas]::Max)
      $xml = New-Object Xml.XmlDocument
      $xml.Load($jsonReader)
      $xml
   }
   finally
   {
      $jsonReader.Close()
      $mStream.Dispose()
   }
}
}
 
function Convert-XmlToJson {
PARAM([Parameter(ValueFromPipeline=$true)][Xml]$xml)
PROCESS {
   $mStream = New-Object System.IO.MemoryStream
   $jsonWriter = [System.Runtime.Serialization.Json.JsonReaderWriterFactory]::CreateJsonWriter($mStream)
   try
   {
     $xml.Save($jsonWriter)
     $bytes = $mStream.ToArray()
     [System.Text.Encoding]::UTF8.GetString($bytes,0,$bytes.Length)
   }
   finally
   {
     $jsonWriter.Close()
     $mStream.Dispose()
   }
}
}

function New-Json {
[CmdletBinding()]
param([Parameter(ValueFromPipeline=$true)][HashTable]$InputObject) 
begin { 
   $ser = @{}
   $jsona = @()
}
process {
   $jsoni = 
   foreach($input in $InputObject.GetEnumerator() | Where { $_.Value } ) {
      if($input.Value -is [Hashtable]) {
         '"'+$input.Key+'": ' + (New-JSon $input.Value)
      } else {
         $type = $input.Value.GetType()
         if(!$Ser.ContainsKey($Type)) {
            $Ser.($Type) = New-Object System.Runtime.Serialization.Json.DataContractJsonSerializer $type
         }
         $stream = New-Object System.IO.MemoryStream
         $Ser.($Type).WriteObject( $stream, $Input.Value )
         '"'+$input.Key+'": ' + (Read-Stream $stream)
      }
   }

   $jsona += "{`n" +($jsoni -join ",`n")+ "`n}"
}
end { 
   if($jsona.Count -gt 1) {
      "[$($jsona -join ",`n")]" 
   } else {
      $jsona
   }
}}


## Rather than rewriting ConvertTo-Xml ...
Function ConvertTo-Json {
[CmdletBinding()]
Param(
   [Parameter(Mandatory=$true,Position=1,ValueFromPipeline=$true)]$InputObject
,
   [Parameter(Mandatory=$false)][Int]$Depth=1
,
   [Switch]$NoTypeInformation
)
END { 
   ## You must output ALL the input at once 
   ## ConvertTo-Xml outputs differently if you just have one, so your results would be different
   $input | ConvertTo-Xml -Depth:$Depth -NoTypeInformation:$NoTypeInformation -As Document | Convert-CliXmlToJson
}
}

Function Convert-CliXmlToJson {
PARAM(
   [Parameter(ValueFromPipeline=$true)][Xml.XmlNode]$xml
)
BEGIN {
   $xmlToJsonXsl = @'
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<!--
  CliXmlToJson.xsl

  Copyright (c) 2006,2008 Doeke Zanstra
  Copyright (c) 2009 Joel Bennett
  All rights reserved.

  Redistribution and use in source and binary forms, with or without modification, 
  are permitted provided that the following conditions are met:

  Redistributions of source code must retain the above copyright notice, this 
  list of conditions and the following disclaimer. Redistributions in binary 
  form must reproduce the above copyright notice, this list of conditions and the 
  following disclaimer in the documentation and/or other materials provided with 
  the distribution.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND 
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
  IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
  INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, 
  BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, 
  DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF 
  LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR 
  OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF 
  THE POSSIBILITY OF SUCH DAMAGE.
-->

  <xsl:output indent="no" omit-xml-declaration="yes" method="text" encoding="UTF-8" media-type="text/x-json"/>
	<xsl:strip-space elements="*"/>
  <!--contant-->
  <xsl:variable name="d">0123456789</xsl:variable>

  <!-- ignore document text -->
  <xsl:template match="text()[preceding-sibling::node() or following-sibling::node()]"/>

  <!-- string -->
  <xsl:template match="text()">
    <xsl:call-template name="escape-string">
      <xsl:with-param name="s" select="."/>
    </xsl:call-template>
  </xsl:template>
  
  <!-- Main template for escaping strings; used by above template and for object-properties 
       Responsibilities: placed quotes around string, and chain up to next filter, escape-bs-string -->
  <xsl:template name="escape-string">
    <xsl:param name="s"/>
    <xsl:text>"</xsl:text>
    <xsl:call-template name="escape-bs-string">
      <xsl:with-param name="s" select="$s"/>
    </xsl:call-template>
    <xsl:text>"</xsl:text>
  </xsl:template>
  
  <!-- Escape the backslash (\) before everything else. -->
  <xsl:template name="escape-bs-string">
    <xsl:param name="s"/>
    <xsl:choose>
      <xsl:when test="contains($s,'\')">
        <xsl:call-template name="escape-quot-string">
          <xsl:with-param name="s" select="concat(substring-before($s,'\'),'\\')"/>
        </xsl:call-template>
        <xsl:call-template name="escape-bs-string">
          <xsl:with-param name="s" select="substring-after($s,'\')"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="escape-quot-string">
          <xsl:with-param name="s" select="$s"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- Escape the double quote ("). -->
  <xsl:template name="escape-quot-string">
    <xsl:param name="s"/>
    <xsl:choose>
      <xsl:when test="contains($s,'&quot;')">
        <xsl:call-template name="encode-string">
          <xsl:with-param name="s" select="concat(substring-before($s,'&quot;'),'\&quot;')"/>
        </xsl:call-template>
        <xsl:call-template name="escape-quot-string">
          <xsl:with-param name="s" select="substring-after($s,'&quot;')"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="encode-string">
          <xsl:with-param name="s" select="$s"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- Replace tab, line feed and/or carriage return by its matching escape code. Can't escape backslash
       or double quote here, because they don't replace characters (&#x0; becomes \t), but they prefix 
       characters (\ becomes \\). Besides, backslash should be seperate anyway, because it should be 
       processed first. This function can't do that. -->
  <xsl:template name="encode-string">
    <xsl:param name="s"/>
    <xsl:choose>
      <!-- tab -->
      <xsl:when test="contains($s,'&#x9;')">
        <xsl:call-template name="encode-string">
          <xsl:with-param name="s" select="concat(substring-before($s,'&#x9;'),'\t',substring-after($s,'&#x9;'))"/>
        </xsl:call-template>
      </xsl:when>
      <!-- line feed -->
      <xsl:when test="contains($s,'&#xA;')">
        <xsl:call-template name="encode-string">
          <xsl:with-param name="s" select="concat(substring-before($s,'&#xA;'),'\n',substring-after($s,'&#xA;'))"/>
        </xsl:call-template>
      </xsl:when>
      <!-- carriage return -->
      <xsl:when test="contains($s,'&#xD;')">
        <xsl:call-template name="encode-string">
          <xsl:with-param name="s" select="concat(substring-before($s,'&#xD;'),'\r',substring-after($s,'&#xD;'))"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise><xsl:value-of select="$s"/></xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <!-- number (no support for javascript mantise) -->
  <xsl:template match="text()[not(string(number())='NaN' or
                       (starts-with(.,'0' ) and . != '0'))]">
    <xsl:value-of select="."/>
  </xsl:template>

  <!-- boolean, case-insensitive -->
  <xsl:template match="text()[translate(.,'TRUE','true')='true']">true</xsl:template>
  <xsl:template match="text()[translate(.,'FALSE','false')='false']">false</xsl:template>

  <!-- root object(s) -->
  <xsl:template match="*" name="base">
    <xsl:if test="not(preceding-sibling::*)">
      <xsl:choose>
        <xsl:when test="count(../*)>1"><xsl:text>[</xsl:text></xsl:when>
        <xsl:otherwise><xsl:text>{</xsl:text></xsl:otherwise>
      </xsl:choose>
    </xsl:if>
    <xsl:call-template name="escape-string">
      <xsl:with-param name="s" select="name()"/>
    </xsl:call-template>
    <xsl:text>:</xsl:text>
    <!-- check type of node -->
    <xsl:choose>
      <!-- null nodes -->
      <xsl:when test="count(child::node())=0">null</xsl:when>
      <!-- other nodes -->
      <xsl:otherwise>
      	<xsl:apply-templates select="child::node()"/>
      </xsl:otherwise>
    </xsl:choose>
    <!-- end of type check -->
    <xsl:if test="following-sibling::*">,</xsl:if>
    <xsl:if test="not(following-sibling::*)">
      <xsl:choose>
        <xsl:when test="count(../*)>1"><xsl:text>]</xsl:text></xsl:when>
        <xsl:otherwise><xsl:text>}</xsl:text></xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:template>

  <!-- properties of objects -->
  <xsl:template match="*[count(../*[name(../*)=name(.)])=count(../*) and count(../*)&gt;1]">
    <xsl:variable name="inArray" select="translate(local-name(),'OBJECT','object')='object' or ../@Type[starts-with(.,'System.Collections') or contains(.,'[]') or (contains(.,'[') and contains(.,']'))]"/>
    <xsl:if test="not(preceding-sibling::*)">
       <xsl:choose>
         <xsl:when test="$inArray"><xsl:text>[</xsl:text></xsl:when>
         <xsl:otherwise>
            <xsl:text>{</xsl:text>
            <xsl:if test="../@Type">
               <xsl:text>"__type":</xsl:text>      
               <xsl:call-template name="escape-string">
                 <xsl:with-param name="s" select="../@Type"/>
               </xsl:call-template>
               <xsl:text>,</xsl:text>      
             </xsl:if>
         </xsl:otherwise>
       </xsl:choose>
    </xsl:if>
    <xsl:choose>
      <xsl:when test="not(child::node())">
        <xsl:call-template name="escape-string">
          <xsl:with-param name="s" select="@Name"/>
        </xsl:call-template>
        <xsl:text>:null</xsl:text>
      </xsl:when>
      <xsl:when test="$inArray">
        <xsl:apply-templates select="child::node()"/>
      </xsl:when>
      <!--
      <xsl:when test="not(@Name) and not(@Type)">
        <xsl:call-template name="escape-string">
          <xsl:with-param name="s" select="local-name()"/>
        </xsl:call-template>
        <xsl:text>:</xsl:text>      
        <xsl:apply-templates select="child::node()"/>
      </xsl:when>
      -->
      <xsl:when test="not(@Name)">
        <xsl:call-template name="escape-string">
          <xsl:with-param name="s" select="local-name()"/>
        </xsl:call-template>
        <xsl:text>:</xsl:text>      
        <xsl:apply-templates select="child::node()"/>
      </xsl:when> 
      <xsl:otherwise>
        <xsl:call-template name="escape-string">
          <xsl:with-param name="s" select="@Name"/>
        </xsl:call-template>
        <xsl:text>:</xsl:text>
        <xsl:apply-templates select="child::node()"/>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:if test="following-sibling::*">,</xsl:if>
    <xsl:if test="not(following-sibling::*)">       
      <xsl:choose>
        <xsl:when test="$inArray"><xsl:text>]</xsl:text></xsl:when>
        <xsl:otherwise><xsl:text>}</xsl:text></xsl:otherwise>
      </xsl:choose>
    </xsl:if>
  </xsl:template>
  
  
  <!-- convert root element to an anonymous container -->
  <xsl:template match="/">
    <xsl:apply-templates select="node()"/>
  </xsl:template>    
</xsl:stylesheet>
'@
}
PROCESS {
   if(Get-Member -InputObject $xml -Name root) {
      Write-Verbose "Ripping to Objects"
      $xml = $xml.root.Objects
   } else {
      Write-Verbose "Was already Objects"
   }
   Convert-Xml -Xml $xml -Xsl $xmlToJsonXsl
}
}

Function ConvertFrom-Xml {
   [CmdletBinding(DefaultParameterSetName="AutoType")]
   PARAM(
      [Parameter(ValueFromPipeline=$true,Mandatory=$true,Position=1)]
      [Xml.XmlNode]
      $xml
      ,
      [Parameter(Mandatory=$true,ParameterSetName="ManualType")]
      [Type]$Type
      ,
      [Switch]$ForceType
   )
   PROCESS{ 
      if($xml.Item("root") -ne $null) {
         return $xml.root.Objects | ConvertFrom-Xml
      } elseif($xml.Item("Objects") -ne $null) {
         return $xml.Objects | ConvertFrom-Xml
      }
      $propbag = @{}
      foreach($name in Get-Member -InputObject $xml -MemberType Properties | Where-Object{$_.Name -notmatch "^__|type"} | Select-Object -ExpandProperty name) {
         Write-Verbose "$Name Type: $($xml.$Name.type)"
         $propbag."$Name" = Convert-Properties $xml."$name"
      }
      if(!$Type -and $xml.HasAttribute("__type")) { $Type = $xml.__Type }
      if($ForceType -and $Type) {
         try {
            $output = New-Object $Type -Property $propbag
         } catch {
            $output = New-Object PSObject -Property $propbag
            $output.PsTypeNames.Insert(0, $xml.__type)
         }
      } else {
				if( $propbag.Count -eq 0) { 
         $output = New-Object PSObject 
				}
				else { 
         $output = New-Object PSObject -Property $propbag
				}
				
         if($Type) {
            $output.PsTypeNames.Insert(0, $Type)
         }
      }
      Write-Output $output
   }
}

Function Convert-Properties {
param($InputObject)
   switch( $InputObject.type ) {
      "object" {
         return (ConvertFrom-Xml -Xml $InputObject)
         break
      } 
      "string" {
         $MightBeADate = $InputObject.get_InnerText() -as [DateTime]
         ## Strings that are actually dates (*grumble* JSON is crap)               
         if($MightBeADate -and $propbag."$Name" -eq $MightBeADate.ToString("G")) {
            return $MightBeADate
         } else {
            return $InputObject.get_InnerText()
         }
         break
      }
      "number" {
         $number = $InputObject.get_InnerText()
         if($number -eq ($number -as [int])) {
            return $number -as [int]
         } elseif($number -eq ($number -as [double])) {
            return $number -as [double]
         } else {
            return $number -as [decimal]
         }
         break
      }
      "boolean" {
         return [bool]::parse($InputObject.get_InnerText())
      }
      "null" {
         return $null
      }
      "array" {
         [object[]]$Items = $( foreach( $item in $InputObject.GetEnumerator() ) {
            Convert-Properties $item
         } )
         return $Items
      }
      default {
         return $InputObject
         break
      }
   }

}



Function ConvertFrom-Json {
   [CmdletBinding(DefaultParameterSetName="StringInput")]
PARAM(
   [Parameter(Mandatory=$true,Position=1,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,ParameterSetName="File")]
   [Alias("PSPath")]
   [string]$File
,
   [Parameter(ValueFromPipeline=$true,Mandatory=$true,Position=1,ParameterSetName="StringInput")]
   [string]$InputObject
,
   [Parameter(Mandatory=$true)]
   [Type]$Type
   ,
   [Switch]$ForceType
)
BEGIN {
   [bool]$AsParameter = $PSBoundParameters.ContainsKey("File") -or $PSBoundParameters.ContainsKey("InputObject") 
}
PROCESS {
   if($PSCmdlet.ParameterSetName -eq "File") {
      [string]$InputObject = @(Get-Content $File) -Join "`n"
      $null = $PSBoundParameters.Remove("File")
   }
   else 
   {
      $null = $PSBoundParameters.Remove("InputObject")
   }
   [Xml.XmlElement]$xml = (Convert-JsonToXml $InputObject).Root
   if($xml) {
      if($xml.Objects) {
         $xml.Objects.Item.GetEnumerator() | ConvertFrom-Xml @PSBoundParameters
      }elseif($xml.Item -and $xml.Item -isnot [System.Management.Automation.PSParameterizedProperty]) {
         $xml.Item | ConvertFrom-Xml @PSBoundParameters
      }else {
         $xml | ConvertFrom-Xml @PSBoundParameters
      }
   } else {
      Write-Error "Failed to parse JSON with JsonReader"
   }
}
}

#########
### The JSON library is dependent on Convert-Xml from my Xml script module

function Convert-Node {
param(
[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
[System.Xml.XmlReader]$XmlReader,
[Parameter(Position=1,Mandatory=$true,ValueFromPipeline=$false)]
[System.Xml.Xsl.XslCompiledTransform]$StyleSheet
) 
PROCESS {
   $output = New-Object IO.StringWriter
   $StyleSheet.Transform( $XmlReader, $null, $output )
   Write-Output $output.ToString()
}
}
   
function Convert-Xml {
#.Synopsis
#  The Convert-XML function lets you use Xslt to transform XML strings and documents.
#.Description
#.Parameter Content
#  Specifies a string that contains the XML to search. You can also pipe strings to Select-XML.
#.Parameter Namespace
#   Specifies a hash table of the namespaces used in the XML. Use the format @{<namespaceName> = <namespaceUri>}.
#.Parameter Path
#   Specifies the path and file names of the XML files to search.  Wildcards are permitted.
#.Parameter Xml
#  Specifies one or more XML nodes to search.
#.Parameter Xsl
#  Specifies an Xml StyleSheet to transform with...
[CmdletBinding(DefaultParameterSetName="Xml")]
PARAM(
   [Parameter(Position=1,ParameterSetName="Path",Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
   [ValidateNotNullOrEmpty()]
   [Alias("PSPath")]
   [String[]]$Path
,
   [Parameter(Position=1,ParameterSetName="Xml",Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
   [ValidateNotNullOrEmpty()]
   [Alias("Node")]
   [System.Xml.XmlNode[]]$Xml
,
   [Parameter(ParameterSetName="Content",Mandatory=$true,ValueFromPipeline=$true)]
   [ValidateNotNullOrEmpty()]
   [String[]]$Content
,
   [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$false)]
   [ValidateNotNullOrEmpty()]
   [Alias("StyleSheet")]
   [String[]]$Xslt
)
BEGIN { 
   $StyleSheet = New-Object System.Xml.Xsl.XslCompiledTransform
   if(Test-Path @($Xslt)[0] -ErrorAction 0) { 
      Write-Verbose "Loading Stylesheet from $(Resolve-Path @($Xslt)[0])"
      $StyleSheet.Load( (Resolve-Path @($Xslt)[0]) )
   } else {
      Write-Verbose "$Xslt"
      $StyleSheet.Load(([System.Xml.XmlReader]::Create((New-Object System.IO.StringReader ($Xslt -join "`n")))))
   }
   [Text.StringBuilder]$XmlContent = [String]::Empty 
}
PROCESS {
   switch($PSCmdlet.ParameterSetName) {
      "Content" {
         $null = $XmlContent.AppendLine( $Content -Join "`n" )
      }
      "Path" {
         foreach($file in Get-ChildItem $Path) {
            Convert-Node -Xml ([System.Xml.XmlReader]::Create((Resolve-Path $file))) $StyleSheet
         }
      }
      "Xml" {
         foreach($node in $Xml) {
            Convert-Node -Xml (New-Object Xml.XmlNodeReader $node) $StyleSheet
         }
      }
   }
}
END {
   if($PSCmdlet.ParameterSetName -eq "Content") {
      [Xml]$Xml = $XmlContent.ToString()
      Convert-Node -Xml $Xml $StyleSheet
   }
}
}


New-Alias fromjson ConvertFrom-Json
New-Alias tojson ConvertTo-Json

#New-Alias ipjs Import-Json
#New-Alias epjs Export-Json
#Import-Json, Export-Json, 


# SIG # Begin signature block
# MIIRDAYJKoZIhvcNAQcCoIIQ/TCCEPkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUqJs83rUpkJfARRWW/QJzu773
# OXuggg5CMIIHBjCCBO6gAwIBAgIBFTANBgkqhkiG9w0BAQUFADB9MQswCQYDVQQG
# EwJJTDEWMBQGA1UEChMNU3RhcnRDb20gTHRkLjErMCkGA1UECxMiU2VjdXJlIERp
# Z2l0YWwgQ2VydGlmaWNhdGUgU2lnbmluZzEpMCcGA1UEAxMgU3RhcnRDb20gQ2Vy
# dGlmaWNhdGlvbiBBdXRob3JpdHkwHhcNMDcxMDI0MjIwMTQ1WhcNMTIxMDI0MjIw
# MTQ1WjCBjDELMAkGA1UEBhMCSUwxFjAUBgNVBAoTDVN0YXJ0Q29tIEx0ZC4xKzAp
# BgNVBAsTIlNlY3VyZSBEaWdpdGFsIENlcnRpZmljYXRlIFNpZ25pbmcxODA2BgNV
# BAMTL1N0YXJ0Q29tIENsYXNzIDIgUHJpbWFyeSBJbnRlcm1lZGlhdGUgT2JqZWN0
# IENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAyiOLIjUemqAbPJ1J
# 0D8MlzgWKbr4fYlbRVjvhHDtfhFN6RQxq0PjTQxRgWzwFQNKJCdU5ftKoM5N4YSj
# Id6ZNavcSa6/McVnhDAQm+8H3HWoD030NVOxbjgD/Ih3HaV3/z9159nnvyxQEckR
# ZfpJB2Kfk6aHqW3JnSvRe+XVZSufDVCe/vtxGSEwKCaNrsLc9pboUoYIC3oyzWoU
# TZ65+c0H4paR8c8eK/mC914mBo6N0dQ512/bkSdaeY9YaQpGtW/h/W/FkbQRT3sC
# pttLVlIjnkuY4r9+zvqhToPjxcfDYEf+XD8VGkAqle8Aa8hQ+M1qGdQjAye8OzbV
# uUOw7wIDAQABo4ICfzCCAnswDAYDVR0TBAUwAwEB/zALBgNVHQ8EBAMCAQYwHQYD
# VR0OBBYEFNBOD0CZbLhLGW87KLjg44gHNKq3MIGoBgNVHSMEgaAwgZ2AFE4L7xqk
# QFulF2mHMMo0aEPQQa7yoYGBpH8wfTELMAkGA1UEBhMCSUwxFjAUBgNVBAoTDVN0
# YXJ0Q29tIEx0ZC4xKzApBgNVBAsTIlNlY3VyZSBEaWdpdGFsIENlcnRpZmljYXRl
# IFNpZ25pbmcxKTAnBgNVBAMTIFN0YXJ0Q29tIENlcnRpZmljYXRpb24gQXV0aG9y
# aXR5ggEBMAkGA1UdEgQCMAAwPQYIKwYBBQUHAQEEMTAvMC0GCCsGAQUFBzAChiFo
# dHRwOi8vd3d3LnN0YXJ0c3NsLmNvbS9zZnNjYS5jcnQwYAYDVR0fBFkwVzAsoCqg
# KIYmaHR0cDovL2NlcnQuc3RhcnRjb20ub3JnL3Nmc2NhLWNybC5jcmwwJ6AloCOG
# IWh0dHA6Ly9jcmwuc3RhcnRzc2wuY29tL3Nmc2NhLmNybDCBggYDVR0gBHsweTB3
# BgsrBgEEAYG1NwEBBTBoMC8GCCsGAQUFBwIBFiNodHRwOi8vY2VydC5zdGFydGNv
# bS5vcmcvcG9saWN5LnBkZjA1BggrBgEFBQcCARYpaHR0cDovL2NlcnQuc3RhcnRj
# b20ub3JnL2ludGVybWVkaWF0ZS5wZGYwEQYJYIZIAYb4QgEBBAQDAgABMFAGCWCG
# SAGG+EIBDQRDFkFTdGFydENvbSBDbGFzcyAyIFByaW1hcnkgSW50ZXJtZWRpYXRl
# IE9iamVjdCBTaWduaW5nIENlcnRpZmljYXRlczANBgkqhkiG9w0BAQUFAAOCAgEA
# UKLQmPRwQHAAtm7slo01fXugNxp/gTJY3+aIhhs8Gog+IwIsT75Q1kLsnnfUQfbF
# pl/UrlB02FQSOZ+4Dn2S9l7ewXQhIXwtuwKiQg3NdD9tuA8Ohu3eY1cPl7eOaY4Q
# qvqSj8+Ol7f0Zp6qTGiRZxCv/aNPIbp0v3rD9GdhGtPvKLRS0CqKgsH2nweovk4h
# fXjRQjp5N5PnfBW1X2DCSTqmjweWhlleQ2KDg93W61Tw6M6yGJAGG3GnzbwadF9B
# UW88WcRsnOWHIu1473bNKBnf1OKxxAQ1/3WwJGZWJ5UxhCpA+wr+l+NbHP5x5XZ5
# 8xhhxu7WQ7rwIDj8d/lGU9A6EaeXv3NwwcbIo/aou5v9y94+leAYqr8bbBNAFTX1
# pTxQJylfsKrkB8EOIx+Zrlwa0WE32AgxaKhWAGho/Ph7d6UXUSn5bw2+usvhdkW4
# npUoxAk3RhT3+nupi1fic4NG7iQG84PZ2bbS5YxOmaIIsIAxclf25FwssWjieMwV
# 0k91nlzUFB1HQMuE6TurAakS7tnIKTJ+ZWJBDduUbcD1094X38OvMO/++H5S45Ki
# 3r/13YTm0AWGOvMFkEAF8LbuEyecKTaJMTiNRfBGMgnqGBfqiOnzxxRVNOw2hSQp
# 0B+C9Ij/q375z3iAIYCbKUd/5SSELcmlLl+BuNknXE0wggc0MIIGHKADAgECAgFR
# MA0GCSqGSIb3DQEBBQUAMIGMMQswCQYDVQQGEwJJTDEWMBQGA1UEChMNU3RhcnRD
# b20gTHRkLjErMCkGA1UECxMiU2VjdXJlIERpZ2l0YWwgQ2VydGlmaWNhdGUgU2ln
# bmluZzE4MDYGA1UEAxMvU3RhcnRDb20gQ2xhc3MgMiBQcmltYXJ5IEludGVybWVk
# aWF0ZSBPYmplY3QgQ0EwHhcNMDkxMTExMDAwMDAxWhcNMTExMTExMDYyODQzWjCB
# qDELMAkGA1UEBhMCVVMxETAPBgNVBAgTCE5ldyBZb3JrMRcwFQYDVQQHEw5XZXN0
# IEhlbnJpZXR0YTEtMCsGA1UECxMkU3RhcnRDb20gVmVyaWZpZWQgQ2VydGlmaWNh
# dGUgTWVtYmVyMRUwEwYDVQQDEwxKb2VsIEJlbm5ldHQxJzAlBgkqhkiG9w0BCQEW
# GEpheWt1bEBIdWRkbGVkTWFzc2VzLm9yZzCCASIwDQYJKoZIhvcNAQEBBQADggEP
# ADCCAQoCggEBAMfjItJjMWVaQTECvnV/swHQP0FTYUvRizKzUubGNDNaj7v2dAWC
# rAA+XE0lt9JBNFtCCcweDzphbWU/AAY0sEPuKobV5UGOLJvW/DcHAWdNB/wRrrUD
# dpcsapQ0IxxKqpRTrbu5UGt442+6hJReGTnHzQbX8FoGMjt7sLrHc3a4wTH3nMc0
# U/TznE13azfdtPOfrGzhyBFJw2H1g5Ag2cmWkwsQrOBU+kFbD4UjxIyus/Z9UQT2
# R7bI2R4L/vWM3UiNj4M8LIuN6UaIrh5SA8q/UvDumvMzjkxGHNpPZsAPaOS+RNmU
# Go6X83jijjbL39PJtMX+doCjS/lnclws5lUCAwEAAaOCA4EwggN9MAkGA1UdEwQC
# MAAwDgYDVR0PAQH/BAQDAgeAMDoGA1UdJQEB/wQwMC4GCCsGAQUFBwMDBgorBgEE
# AYI3AgEVBgorBgEEAYI3AgEWBgorBgEEAYI3CgMNMB0GA1UdDgQWBBR5tWPGCLNQ
# yCXI5fY5ViayKj6xATCBqAYDVR0jBIGgMIGdgBTQTg9AmWy4SxlvOyi44OOIBzSq
# t6GBgaR/MH0xCzAJBgNVBAYTAklMMRYwFAYDVQQKEw1TdGFydENvbSBMdGQuMSsw
# KQYDVQQLEyJTZWN1cmUgRGlnaXRhbCBDZXJ0aWZpY2F0ZSBTaWduaW5nMSkwJwYD
# VQQDEyBTdGFydENvbSBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eYIBFTCCAUIGA1Ud
# IASCATkwggE1MIIBMQYLKwYBBAGBtTcBAgEwggEgMC4GCCsGAQUFBwIBFiJodHRw
# Oi8vd3d3LnN0YXJ0c3NsLmNvbS9wb2xpY3kucGRmMDQGCCsGAQUFBwIBFihodHRw
# Oi8vd3d3LnN0YXJ0c3NsLmNvbS9pbnRlcm1lZGlhdGUucGRmMIG3BggrBgEFBQcC
# AjCBqjAUFg1TdGFydENvbSBMdGQuMAMCAQEagZFMaW1pdGVkIExpYWJpbGl0eSwg
# c2VlIHNlY3Rpb24gKkxlZ2FsIExpbWl0YXRpb25zKiBvZiB0aGUgU3RhcnRDb20g
# Q2VydGlmaWNhdGlvbiBBdXRob3JpdHkgUG9saWN5IGF2YWlsYWJsZSBhdCBodHRw
# Oi8vd3d3LnN0YXJ0c3NsLmNvbS9wb2xpY3kucGRmMGMGA1UdHwRcMFowK6ApoCeG
# JWh0dHA6Ly93d3cuc3RhcnRzc2wuY29tL2NydGMyLWNybC5jcmwwK6ApoCeGJWh0
# dHA6Ly9jcmwuc3RhcnRzc2wuY29tL2NydGMyLWNybC5jcmwwgYkGCCsGAQUFBwEB
# BH0wezA3BggrBgEFBQcwAYYraHR0cDovL29jc3Auc3RhcnRzc2wuY29tL3N1Yi9j
# bGFzczIvY29kZS9jYTBABggrBgEFBQcwAoY0aHR0cDovL3d3dy5zdGFydHNzbC5j
# b20vY2VydHMvc3ViLmNsYXNzMi5jb2RlLmNhLmNydDAjBgNVHRIEHDAahhhodHRw
# Oi8vd3d3LnN0YXJ0c3NsLmNvbS8wDQYJKoZIhvcNAQEFBQADggEBACY+J88ZYr5A
# 6lYz/L4OGILS7b6VQQYn2w9Wl0OEQEwlTq3bMYinNoExqCxXhFCHOi58X6r8wdHb
# E6mU8h40vNYBI9KpvLjAn6Dy1nQEwfvAfYAL8WMwyZykPYIS/y2Dq3SB2XvzFy27
# zpIdla8qIShuNlX22FQL6/FKBriy96jcdGEYF9rbsuWku04NqSLjNM47wCAzLs/n
# FXpdcBL1R6QEK4MRhcEL9Ho4hGbVvmJES64IY+P3xlV2vlEJkk3etB/FpNDOQf8j
# RTXrrBUYFvOCv20uHsRpc3kFduXt3HRV2QnAlRpG26YpZN4xvgqSGXUeqRceef7D
# dm4iTdHK5tIxggI0MIICMAIBATCBkjCBjDELMAkGA1UEBhMCSUwxFjAUBgNVBAoT
# DVN0YXJ0Q29tIEx0ZC4xKzApBgNVBAsTIlNlY3VyZSBEaWdpdGFsIENlcnRpZmlj
# YXRlIFNpZ25pbmcxODA2BgNVBAMTL1N0YXJ0Q29tIENsYXNzIDIgUHJpbWFyeSBJ
# bnRlcm1lZGlhdGUgT2JqZWN0IENBAgFRMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEWMCMGCSqGSIb3DQEJBDEWBBT6qQYgJan6
# 3I9xV61y+z9CGzM59DANBgkqhkiG9w0BAQEFAASCAQA1CQt7IXQmbYkwmEaVpyfR
# iHZYa8WfGG7nTYYPIZ7wZDV4b4SuAN+K97zx1H99JGUvB68xz4W64MnYu+JMsAL5
# mf301A7ZjD2+o18HpwLFqm0tQK4TRv2fhSQ+4uBzuaD2qcDuVhMsryGPmd9FPHwY
# 4g8LG2M9Hqb98pAbKhs5EZu3URycu9N6Z5F8+/ILZCG7FRx9/EBrR3TdxckMN9GC
# mj6kwrTae63TUxRzsVmiE8Zslar9I3A9LTvDYhXuRMMIpnw1OuCuTBErhnVaIs5Q
# MiBO4M9rNPQqF8//uNkhKJkCfbcixGC5Yz47EdyrjrdzeKO2/ECrPJbvcXBcp//M
# SIG # End signature block
