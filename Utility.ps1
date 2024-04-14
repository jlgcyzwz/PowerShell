<#
 ユーザーシェルフォルダ取得
#>
function global:Get-UserShellFolder
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$false, 
                   Position=0)]
        [string]$Name
    )

    if ($Name) {
        $shell = New-Object -ComObject Shell.Application
        try {
            $namespace = $shell.Namespace(('shell:{0}') -f $Name)
            if ($namespace) {
                $namespace.Self.Path
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null
                $namespace = $null
            }
        }
        finally {
            if ($shell) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
                $shell = $null
            }
        }
    }
    else {
        Get-ItemProperty 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
    }
}

<#
 フォルダパス取得
#>
function global:Get-Folder
{
    [OutputType([String])]
    Param
    (
        [Parameter(Mandatory=$false, 
                   Position=0)]
        [System.Environment+SpecialFolder]
        $SpecialFolder
    )

    [Environment]::GetFolderPath($SpecialFolder)
}

<#
 ショートカットを作成
#>
function global:New-Shortcut
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param
    (
        # ターゲットパス
        [Parameter(Mandatory=$true, 
                   Position=0)]
        [string]
        $TargetPath,

        # 名前
        [Parameter(Mandatory=$true, 
                   Position=1)]
        [string]
        $Name,

        # タイプ
        [Parameter(Mandatory=$true, 
                   Position=2)]
        [ValidateSet('Lnk', 'Url')]
        [string]
        $ShortcutType,

        # ディレクトリ
        [Parameter(Mandatory=$false)]
        [string]
        $Directory,

        # アイコンファイル、アイコンインデックス
        [Parameter(Mandatory=$false)]
        [string]
        $IconLocation,

        # 引数
        [Parameter(Mandatory=$false)]
        [string]
        $Arguments,

        # 作業フォルダー
        [Parameter(Mandatory=$false)]
        [string]
        $WorkingDirectory,

        # コメント
        [Parameter(Mandatory=$false)]
        [string]
        $Description,

        # ショートカットキー
        [Parameter(Mandatory=$false)]
        [string]
        $HotKey,

        # 実行時の大きさ
        [Parameter(Mandatory=$false)]
        [ValidateSet('通常のウィンドウ', '最小化', '最大化')]
        [string]
        $WindowStyle = '通常のウィンドウ'
    )

    if($PSCmdlet.ShouldProcess($path, 'ショートカット作成')) {
        $shell = New-Object -comObject WScript.Shell
        try {
            $path = switch($ShortcutType) { 'Lnk' { '{0}.lnk' -f $Name  } 'Url' { '{0}.url' -f $Name } }
            if ([string]::IsNullOrEmpty($Directory)) {
                $path = Join-Path (Convert-Path .) $path
            } else {
                $path = Join-Path $Directory $path
            }
            $shortcut = $shell.CreateShortcut($path)
            try {
                $shortcut.TargetPath = $TargetPath
                if ($ShortcutType -eq 'Lnk') {
                    if ($IconLocation) { $shortcut.IconLocation = $IconLocation }
                    if (![string]::IsNullOrEmpty($Arguments)) { $shortcut.Arguments = $Arguments }
                    if (![string]::IsNullOrEmpty($WorkingDirectory)) { $shortcut.WorkingDirectory = $WorkingDirectory }
                    if (![string]::IsNullOrEmpty($Description)) { $shortcut.Description = $Description }
                    switch ($WindowStyle) {
                        '通常のウィンドウ' { $Shortcut.WindowStyle = 1 }
                        '最小化' { $Shortcut.WindowStyle = 7 }
                        '最大化' { $Shortcut.WindowStyle = 3 }
                    }
                } 
                $Shortcut.Save()
            }
            finally {
                if ($shortcut) {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shortcut) | Out-Null
                    $shortcut = $null
                }
            }
        }
        finally {
            if ($shell) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
                $shell = $null
            }
        }
    }
 }

 <#
 PowerShellのパスを取得
#>
function global:Get-PowerShellPath
{
    [CmdletBinding()]
    [OutputType([string])]
    param
    (
        # ISE
        [Parameter(Mandatory=$false)]
        [switch]
        $Ise
    )

    $name = if ($Ise.IsPresent) { 'powershell_ise.exe' } else { 'powershell.exe' }
    $path = Join-Path $PSHOME $name
    if (Test-Path $path) {
        $path
    } else {
        $env:Path -split ';' | % {
            Join-Path $_ $name
        } | ? {
            Test-Path $_
        } | Select-Object -First 1
    }
}


<#
 PowerShellショートカットを作成
#>
function global:New-PowerShellShortcut
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param
    (
        # 名前
        [Parameter(Mandatory=$false, 
                   Position=0)]
        [string]
        $Name,

        # ディレクトリ
        [Parameter(Mandatory=$false, 
                   Position=1)]
        [string]
        $Directory,

        # 実行ポリシー
        # 引数
        [Parameter(Mandatory=$false)]
        [ValidateSet('AllSigned', 'Bypass', 'Default', 'RemoteSigned', 'Restricted', 'Undefined', 'Unrestricted')]
        [string]
        $ExecutionPolicy,

        # 作業フォルダー
        [Parameter(Mandatory=$false)]
        [string]
        $WorkingDirectory,

        # コメント
        [Parameter(Mandatory=$false)]
        [string]
        $Description,

        # ショートカットキー
        [Parameter(Mandatory=$false)]
        [string]
        $HotKey,

        # 実行時の大きさ
        [Parameter(Mandatory=$false)]
        [ValidateSet('通常のウィンドウ', '最小化', '最大化')]
        [string]
        $WindowStyle = '通常のウィンドウ',

        # ISE
        [Parameter(Mandatory=$false)]
        [switch]
        $Ise
    )

    if (!$Name) {
        $Name = if (!$Ise.IsPresent) { 'PowerShell' } else { 'PowerShell ISE' }
    }

    $path = Get-PowerShellPath
    if (!$Ise.IsPresent) {
        $arguments = if ($ExecutionPolicy) {
            '-ExecutionPolicy {0}' -f $ExecutionPolicy
        } 
        else {
            $null
        }
        New-Shortcut $path $Name -ShortcutType Lnk -Directory $Directory -IconLocation ('{0},0' -f $path) -Arguments $arguments -WorkingDirectory $WorkingDirectory -Description $Description -HotKey $HotKey -WindowStyle $WindowStyle
    }
    else {
        $pathIse = Get-PowerShellPath -Ise
        $arguments = if ($ExecutionPolicy) {
            '-ExecutionPolicy {0} -Command Start-Process powershell_ise' -f $ExecutionPolicy
        } 
        else {
            '-Command Start-Process powershell_ise'
        }
        New-Shortcut $path $Name -ShortcutType Lnk -Directory $Directory -IconLocation ('{0},0' -f $pathIse) -Arguments $arguments -WorkingDirectory $WorkingDirectory -Description $Description -HotKey $HotKey -WindowStyle 最小化
    }
}
