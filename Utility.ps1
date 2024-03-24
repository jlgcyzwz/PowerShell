<#
 COMオブジェクト開放
#>
function global:Release-ComObject
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([int])]
    Param
    (
        #COMオブジェクト
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   Position=0)]
        [ValidateNotNull()]
        [object]
        $Object
    )
    if($PSCmdlet.ShouldProcess($Object, 'ReleaseComObject')) {
        $count = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
        Write-Verbose ('ReleaseComObjectの戻り値={0}' -f $count)
        $count
    }
}


<#
 フォルダ取得
#>
function global:Get-Folder
{
    [CmdletBinding()]
    [OutputType([String])]
    param
    (
    )

    $shell = New-Object -ComObject Shell.Application
    try {
        Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions | Get-ItemProperty | ? {
            (($_ | Get-Member -MemberType NoteProperty) | % { $_.Name }) -contains 'Name'
        } | ? {
            $namespace = $shell.Namespace(('shell:{0}') -f $_.Name)
            if ($namespace) {
                Test-Path $namespace.Self.Path
                $namespace | Release-ComObject | Out-Null
            } else {
                $false
            }
        } | Select-Object -Property Name, Icon | Sort-Object Name
    }
    finally {
        $shell | Release-ComObject | Out-Null
        $shell = $null
    }
}

<#
 フォルダパス取得
#>
function global:Get-FolderPath
{
    [CmdletBinding(DefaultParameterSetName='Environment')]
    [OutputType([String])]
    Param
    (
        [Parameter(Mandatory=$false, 
                   Position=0,
                   ParameterSetName='Environment')]
        [System.Environment+SpecialFolder]
        $SpecialFolder,

        [Parameter(Mandatory=$false, 
                   Position=0,
                   ParameterSetName='Shell')]
        [string]
        $ShellName
    )

    switch($PSCmdlet.ParameterSetName) {
        'Environment' {
            [Environment]::GetFolderPath($SpecialFolder)
        }
        'Shell' {
            $shell = New-Object -ComObject Shell.Application
            try {
                $shell.Namespace(('shell:{0}') -f $ShellName).Self.Path
            }
            finally {
                $shell | Release-ComObject | Out-Null
                $shell = $null
            }
        }
    }
}

<#
 コマンドライン引数を作成
#>
function global:New-ComandlineArguments
{
    [CmdletBinding()]
    [OutputType([String])]
    Param
    (
        # 引数
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   Position=0)]
        [string[]]
        $Arguments
    )

    Begin
    {
        $args = @()
    }
    Process
    {
        $args += $Arguments
    }
    End
    {
        ($args | % {
            #とりあえず今は単純に空白を含むなら"で囲む
            if ($_.Contains(' ')) {
                '"{0}"' -f $_
            }
            else {
                $_
            }
        }) -join ' '
    }
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
        [Parameter(Mandatory=$false, 
                   Position=3)]
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
                    if ($IconLocation) {
                        $shortcut.IconLocation = $IconLocation
                    }
                    else {
                        $shortcut.IconLocation = '{0},0' -f $TargetPath
                    }
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
                    $shortcut | Release-ComObject | Out-Null
                    $shortcut = $null
                }
            }
        }
        finally {
            $shell | Release-ComObject | Out-Null
            $shell = $null
        }
    }
 }

 <#
 PowerShellのパスを取得
#>
function global:Get-PowerShellPath
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([string])]
    param
    (
        # ISE
        [Parameter(Mandatory=$false)]
        [switch]
        $Ise
    )

    $name = if ($Ise.IsPresent) { 'powershell_ise' } else { 'powershell' }
    $process = Get-Process -Name $name -ErrorAction SilentlyContinue | ? { ![string]::IsNullOrEmpty($_.Path) } | Select-Object -First 1
    if ($process) {
        $process.Path
    } else {
        $path = $null
        while(!$path) {
            Start-Process $name -WindowStyle Hidden
            while(($process = Get-Process -Name $name -ErrorAction SilentlyContinue | ? { ![string]::IsNullOrEmpty($_.Path) } | Select-Object -First 1) -eq $null) {
                Start-Sleep -Milliseconds 1
            }
            $path = $process.Path
            $process | Stop-Process
        }
        $path
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
        $WindowStyle = '通常のウィンドウ',

        # ISE
        [Parameter(Mandatory=$false)]
        [switch]
        $Ise
    )

    if (!$Name) {
        $Name = if (!$Ise.IsPresent) { 'PowerShell' } else { 'PowerShell ISE' }
    }

    $powerShellPath = ''
    $powerShellPath = Get-PowerShellPath

    if (!$Ise.IsPresent) {
        New-Shortcut $powerShellPath $Name -ShortcutType Lnk -Arguments '-ExecutionPolicy RemoteSigned'
    }
    else {
        $isePath = Get-PowerShellPath -Ise
        New-Shortcut $powerShellPath $Name -ShortcutType Lnk -IconLocation $isePath -Arguments '-ExecutionPolicy RemoteSigned -Command Start-Process powershell_ise' -WindowStyle 最小化
    }
}

New-PowerShellShortcut -Verbose
New-PowerShellShortcut -Ise -Verbose