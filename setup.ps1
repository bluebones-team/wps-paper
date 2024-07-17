# -*- coding: utf-8 -*-
# PowerShell >= 5.0

$FILE = $MyInvocation.MyCommand.Path
$JSADDON_DIR = Join-Path $env:APPDATA "kingsoft/wps/jsaddons"
$XML_PATH = Join-Path $JSADDON_DIR "publish.xml"
$ADDON_NAME = $null
$NAME = $null
$VERSION = $null

function Get-AddonInfo($Path) {
    $text = Get-Content -Path $Path -Raw
    $nameMatch = [regex]::Match($text, "name: '([^ ]+?)'")
    $versionMatch = [regex]::Match($text, "version: '([^ ]+?)'")
    if ($nameMatch -and $versionMatch) {
        $name = $nameMatch.Groups[1].Value
        $version = $versionMatch.Groups[1].Value
        Write-Host "Get name and version: $name, $version"
        return $name, $version
    }
    else {
        throw "Failed to get name and version"
    }
}
function Copy-Dir($OldDir, $NewDir) {
    Remove-Dir
    $ignorePatterns = @(".git\", "README.md")
    if (Test-Path ".gitignore") {
        $ignorePatterns += Get-Content ".gitignore"
    }
    Copy-Item -Path $OldDir -Destination $NewDir -Recurse -Force -Exclude $ignorePatterns
    Write-Host "Copy '$OldDir' to '$NewDir'"
}
function Add-XML() {
    if (-not (Test-Path $XML_PATH)) {
        Set-Content -Path $XML_PATH -Value "<jsplugins>`n</jsplugins>"
    }
    $xml = [xml](Get-Content $XML_PATH)
    $nodes = $xml.SelectNodes("//jsplugin[@name=""$NAME""]")
    foreach ($node in $nodes) {
        $xml.DocumentElement.RemoveChild($node)
    }
    $newNode = $xml.CreateElement("jsplugin")
    $newNode.SetAttribute("name", $NAME)
    $newNode.SetAttribute("type", "wps")
    $newNode.SetAttribute("url", "https://api.github.com/repos/Cubxx/wps-paper/zipball")
    $newNode.SetAttribute("version", $VERSION)
    $xml.DocumentElement.AppendChild($newNode)
    Write-Host "Add registration info: $($newNode.OuterXml)"
    $xml.Save($XML_PATH)
}
function Remove-Dir() {
    $dirs = Get-ChildItem -Path $JSADDON_DIR -Directory | Where-Object {
        ($_.Name -like "$($NAME)_*") -and (Test-Path $_.FullName -PathType Container)
    }
    if ($dirs.Count -eq 0) {
        Write-Host "Can't find old folder, no need to delete"
        return
    }
    foreach ($e in $dirs) {
        Remove-Item -Path $e.FullName -Recurse -Force
        Write-Host "Remove old folder: $($e.Name)"
    }
}
function Remove-XML() {
    if (-not (Test-Path $XML_PATH)) {
        Write-Host "Can't find registration info, no need to delete"
        return
    }
    $xml = [xml](Get-Content $XML_PATH)
    $nodes = $xml.SelectNodes("//jsplugin[@name=""$NAME""]")
    if ($nodes.Count -eq 0) {
        Write-Host "Can't find registration info, no need to delete"
        return
    }
    foreach ($node in $nodes) {
        $xml.DocumentElement.RemoveChild($node)
        Write-Host "Remove registration info: $($node.OuterXml)"
    }
    $xml.Save($XML_PATH)
}
function Select-Option {
    param(
        [string]$title,
        [array]$options,
        [string]$optionTitle
    )

    Write-Host $title
    for ($i = 0; $i -lt $options.Count; $i++) {
        Write-Host "  $($i + 1). $(if ($optionTitle) { $options[$i].($optionTitle) } else { $options[$i] })"
    }
    $selection = Read-Host "Enter number (1-$($options.Count))"
    $option = $options[$selection - 1]

    if ($option) {
        return $option
    }
    else {
        Write-Error "Invalid number"
        return Select-Option $title $options
    }
}
function Install-Addon($srcDir) {
    if (-not (Test-Path $JSADDON_DIR)) {
        mkdir $JSADDON_DIR
    }
    $targetDir = Join-Path $JSADDON_DIR $ADDON_NAME
    Copy-Dir $srcDir $targetDir
    Add-XML
    Write-Host "Successfully install $ADDON_NAME, the current folder can be deleted"
}
function Update-Addon() {
    $tempFile = 'temp.js'
    Invoke-RestMethod 'https://raw.kkgithub.com/Cubxx/wps-paper/main/config.js' -OutFile $tempFile 
    $NEW_NAME, $NEW_VERSION = Get-AddonInfo $tempFile
    Remove-Item $tempFile
    if ($NEW_VERSION -eq $VERSION) {
        Write-Host "Current version is latest: $VERSION"
    }
    else {
        $tempZip = 'temp.zip'
        Invoke-WebRequest 'https://api.kkgithub.com/repos/Cubxx/wps-paper/zipball' -OutFile $tempZip
        $targetDir = Join-Path $env:TEMP $NEW_NAME
        Expand-Archive $tempZip $targetDir -Force
        Remove-Item $tempZip

        $global:NAME, $global:VERSION = $NEW_NAME, $NEW_VERSION
        $global:ADDON_NAME = $global:NAME + '_' + $global:VERSION
        $srcDir = (Get-ChildItem -Path $targetDir -Directory | Select-Object -First 1).FullName 
        Install-Addon $srcDir
        
        Remove-Item $targetDir -Recurse -Force
    }
    Write-Host "Successfully update $ADDON_NAME, the current folder can be deleted"
}
function Uninstall-Addon() {
    Remove-Dir $JSADDON_DIR
    Remove-XML
    Write-Host "Successfully uninstall $ADDON_NAME, the current folder can be deleted"
}

try {
    if ($FILE -like "$JSADDON_DIR*") {
        throw "Cannot run in the current folder"
    }
    $option = Select-Option "Welcome to WPS addon installation guide" @(
        @{title = "Install"; fn = { Install-Addon (Split-Path -Path $FILE -Parent) } }, 
        @{title = "Update (Need networking)"; fn = { Update-Addon } }, 
        @{title = "Uninstall"; fn = { Uninstall-Addon } }
    ) "title"
    $NAME, $VERSION = Get-AddonInfo "config.js"
    $ADDON_NAME = $NAME + '_' + $VERSION
    & $option.fn
}
catch {
    Write-Error $_
}
Read-Host "Press any key to exit"