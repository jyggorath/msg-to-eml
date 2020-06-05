[CmdletBinding()]
Param(
	[String]$Path = "",
	[Parameter(Mandatory = $false)]
	[Switch]$Help
)

function B64LineBreaks {
	Param(
		[Parameter(Mandatory = $true, Position = 0)]
		[String]$Base64String
	)
	$Base64StringLineBreaks = ""
	for ($i = 0; $i -lt $Base64String.Length; $i++) {
		if ($i % 76 -eq 0 -and $i -ne 0) {
			$Base64StringLineBreaks += "`r`n"
		}
		$Base64StringLineBreaks += $Base64String[$i]
	}
	return $Base64StringLineBreaks
}

function Get-EMLFileName {
	Param(
		[Parameter(Mandatory = $true, Position = 0)]
		$MsgFile
	)
	$MSGFileName = $MsgFile.Name
	$EMLFileName = $MSGFileName.Replace($MsgFile.Extension, ".eml")
	if ($EMLFileName -eq $MSGFileName) {
		$EMLFileName = "$EMLFileName.eml"
	}
	return $EMLFileName
}

if ($Help -or $Path -eq "") {
	Write-Host "Converts provided MSG file to EML format in the current directory."
	Write-Host "Usage: MSG-to-EML.ps1 -Path .\path\to\mail.msg"
	exit
}

if (-not (Test-Path $Path)) {
	Write-Host "File not found: $Path"
	exit
}
$MsgFile = Get-Item "$Path"

# Load mimetype assembly
Try {
	Add-Type -AssemblyName "System.Web"
}
Catch {
	Write-Host "Could't add assembly: System.Web`nOld .NET version?"
	exit
}

# Setup Outlook object thingy
Try {
	$outlook = New-Object -ComObject Outlook.Application
}
Catch {
	Write-Host "Outlook must be running"
	exit
}

# Load MSG file
Try {
	$MSG = $outlook.CreateItemFromTemplate($MsgFile.FullName)
}
Catch {
	Write-Host "Couldn't load $Path, not MSG file format?"
	exit
}

$Headers = $MSG.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")

$BodyHtmlBase64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($MSG.HTMLBody))
$BodyHtmlBase64 = B64LineBreaks -Base64String "$BodyHtmlBase64"

$BodyPlainBase64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($MSG.Body))
$BodyPlainBase64 = B64LineBreaks -Base64String "$BodyPlainBase64"

$MarkOuter = "Mark=_111"
$MarkInner = "Mark=_222"

$Attachments = New-Object System.Collections.ArrayList
$ContentIdCounter = 1
foreach ($attachment in $MSG.Attachments) {
	$o = New-Object -TypeName psobject
	$TempFileName = "$($env:TEMP)\msg-to-eml.file"
	$attachment.SaveAsFile($TempFileName)
	$o | Add-Member -MemberType NoteProperty -Name Base64Content	-Value ""
	$o | Add-Member -MemberType NoteProperty -Name MimeType			-Value ""
	$o | Add-Member -MemberType NoteProperty -Name FileName			-Value ""
	$o | Add-Member -MemberType NoteProperty -Name ContentId		-Value ""
	$o.Base64Content	= [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($TempFileName))
	$o.Base64Content	= B64LineBreaks -Base64String "$($o.Base64Content)"
	$o.MimeType			= [System.Web.MimeMapping]::GetMimeMapping($TempFileName)
	$o.ContentId		= "<$ContentIdCounter@$ContentIdCounter>"
	$ContentIdCounter++
	if ($attachment.FileName) {
		$o.FileName = $attachment.FileName
	}
	else {
		$o.FileName = $attachment.DisplayName
	}
	$Attachments.Add($o) | Out-Null
}

$EMLFileName = Get-EMLFileName -MSGFile $MsgFile
$EML = ""

if ($Attachments) {
	$Headers = $Headers.Replace("Content-Transfer-Encoding: binary`r`n", "")
	$Headers = $Headers.Replace("Content-Type: application/ms-tnef; name=`"winmail.dat`"", "Content-Type: multipart/related;`r`n`tboundary=`"$MarkOuter`"")
	$EML  = "$Headers`r`n"
	$EML += "`r`n"
	$EML += "This is a multi-part message in MIME format.`r`n"
	$EML += "`r`n"
	$EML += "--$MarkOuter`r`n"
	$EML += "Content-Type: multipart/alternative;`r`n"
	$EML += "`tboundary=`"$MarkInner`"`r`n"
	$EML += "`r`n"
	$EML += "`r`n"
	$EML += "--$MarkInner`r`n"
	$EML += "Content-Type: text/plain;`r`n"
	$EML += "`tcharset=`"utf-8`"`r`n"
	$EML += "Content-Transfer-Encoding: base64`r`n"
	$EML += "`r`n"
	$EML += "$BodyPlainBase64`r`n"
	$EML += "`r`n"
	$EML += "--$MarkInner`r`n"
	$EML += "Content-Type: text/html;`r`n"
	$EML += "`tcharset=`"utf-8`"`r`n"
	$EML += "Content-Transfer-Encoding: base64`r`n"
	$EML += "`r`n"
	$EML += "$BodyHtmlBase64`r`n"
	$EML += "`r`n"
	$EML += "--$MarkInner--`r`n"
	$EML += "`r`n"
	foreach ($attachment in $Attachments) {
		$EML += "--$MarkOuter`r`n"
		$EML += "Content-Type: $($attachment.MimeType);`r`n"
		$EML += "`tname=`"$($attachment.FileName)`"`r`n"
		$EML += "Content-Transfer-Encoding: base64`r`n"
		$EML += "Content-Disposition: attachment;`r`n"
		$EML += "`tfilename=`"$($attachment.FileName)`"`r`n"
		$EML += "Content-ID: $($attachment.ContentId)`r`n"
		$EML += "`r`n"
		$EML += "$($attachment.Base64Content)`r`n"
		$EML += "`r`n"
	}
	$EML += "--$MarkOuter--`r`n"
}
else {
	$Headers = $Headers.Replace("Content-Transfer-Encoding: binary`r`n", "")
	$Headers = $Headers.Replace("Content-Type: application/ms-tnef; name=`"winmail.dat`"", "Content-Type: multipart/related;`r`n`tboundary=`"$MarkInner`"")
	$EML  = "$Headers`r`n"
	$EML += "`r`n"
	$EML += "This is a multi-part message in MIME format.`r`n"
	$EML += "`r`n"
	$EML += "--$MarkInner`r`n"
	$EML += "Content-Type: text/plain;`r`n"
	$EML += "`tcharset=`"utf-8`"`r`n"
	$EML += "Content-Transfer-Encoding: base64`r`n"
	$EML += "`r`n"
	$EML += "$BodyPlainBase64`r`n"
	$EML += "`r`n"
	$EML += "--$MarkInner`r`n"
	$EML += "Content-Type: text/html;`r`n"
	$EML += "`tcharset=`"utf-`"`r`n"
	$EML += "Content-Transfer-Encoding: base64`r`n"
	$EML += "`r`n"
	$EML += "$BodyHtmlBase64`r`n"
	$EML += "`r`n"
	$EML += "--$MarkInner--`r`n"
}

[System.IO.File]::WriteAllLines("$((Get-Location).Path)\$EMLFileName", $EML)
