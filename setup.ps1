# PowerShell >= 5.0
$FILE = $MyInvocation.MyCommand.Path
$JSADDON_DIR = "$env:APPDATA/kingsoft/wps/jsaddons"
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
        Write-Host "获取名称和版本号: $name, $version"
        return $name, $version
    }
    else {
        throw "获取名称和版本号失败"
    }
}
function Copy-Dir($OldDir, $NewDir) {
    Remove-Dir
    $ignorePatterns = @(".git\", "README.md")
    if (Test-Path ".gitignore") {
        $ignorePatterns += Get-Content ".gitignore"
    }
    Copy-Item -Path $OldDir -Destination $NewDir -Recurse -Force -Exclude $ignorePatterns
    Write-Host "将 $OldDir 复制到 $NewDir"
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
    Write-Host "添加注册信息 $($newNode.OuterXml)"
    $xml.Save($XML_PATH)
}
function Remove-Dir() {
    $dirs = Get-ChildItem -Path $JSADDON_DIR -Directory | Where-Object {
        ($_.Name -like "$($NAME)_*") -and (Test-Path $_.FullName -PathType Container)
    }
    if ($dirs.Count -eq 0) {
        Write-Host "找不到旧文件夹，无需删除"
        return
    }
    foreach ($e in $dirs) {
        Remove-Item -Path $e.FullName -Recurse -Force
        Write-Host "已删除旧文件夹: $($e.Name)"
    }
}
function Remove-XML() {
    if (-not (Test-Path $XML_PATH)) {
        Write-Host "找不到注册信息，无需删除"
        return
    }
    $xml = [xml](Get-Content $XML_PATH)
    $nodes = $xml.SelectNodes("//jsplugin[@name=""$NAME""]")
    if ($nodes.Count -eq 0) {
        Write-Host "找不到注册信息，无需删除"
        return
    }
    foreach ($node in $nodes) {
        $xml.DocumentElement.RemoveChild($node)
        Write-Host "已删除注册信息 $($node.OuterXml)"
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
    Write-Host "$ADDON_NAME 安装成功, 当前文件夹可删除"
}
function Update-Addon() {
    $tempFile = 'temp.js'
    Invoke-RestMethod 'https://raw.kkgithub.com/Cubxx/wps-paper/main/config.js' -OutFile $tempFile 
    $NEW_NAME, $NEW_VERSION = Get-AddonInfo $tempFile
    if ($NEW_VERSION -eq $VERSION) {
        Write-Host "当前版本 $VERSION 已是最新版本"
    }
    else {
        $tempZip = 'temp.zip'
        Invoke-WebRequest 'https://api.kkgithub.com/repos/Cubxx/wps-paper/zipball' -OutFile $tempZip
        $targetDir = Join-Path $env:TEMP $NEW_NAME
        Expand-Archive $tempZip $targetDir -Force

        $global:NAME, $global:VERSION = $NEW_NAME, $NEW_VERSION
        $global:ADDON_NAME = $global:NAME + '_' + $global:VERSION
        $srcDir = (Get-ChildItem -Path $targetDir -Directory | Select-Object -First 1).FullName 
        Install-Addon $srcDir
        
        Remove-Item $tempZip
        Remove-Item $targetDir -Recurse -Force
    }
    Remove-Item $tempFile
    Write-Host "$ADDON_NAME 更新成功, 当前文件夹可删除"
}
function Uninstall-Addon() {
    Remove-Dir $JSADDON_DIR
    Remove-XML
    Write-Host "$ADDON_NAME 卸载成功, 当前文件夹可删除"
}

try {
    if ($JSADDON_DIR -in $FILE) {
        throw "无法在当前文件夹下操作"
    }
    $option = Select-Option "欢迎使用 WPS 加载项安装向导" @(
        @{title = "安装"; fn = { Install-Addon (Split-Path -Path $FILE -Parent) } }, 
        @{title = "更新（需要联网）"; fn = { Update-Addon } }, 
        @{title = "卸载"; fn = { Uninstall-Addon } }
    ) "title"
    $NAME, $VERSION = Get-AddonInfo "config.js"
    $ADDON_NAME = $NAME + '_' + $VERSION
    & $option.fn
}
catch {
    Write-Error $_
}
Read-Host "按任意键退出"