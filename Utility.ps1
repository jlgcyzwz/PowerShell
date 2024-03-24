<#
 COM�I�u�W�F�N�g�J��
#>
function global:Release-ComObject
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([int])]
    Param
    (
        #COM�I�u�W�F�N�g
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   Position=0)]
        [ValidateNotNull()]
        [object]
        $Object
    )
    if($PSCmdlet.ShouldProcess($Object, 'ReleaseComObject')) {
        $count = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
        Write-Verbose ('ReleaseComObject�̖߂�l={0}' -f $count)
        $count
    }
}


<#
 �t�H���_�擾
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
 �t�H���_�p�X�擾
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
 �R�}���h���C���������쐬
#>
function global:New-ComandlineArguments
{
    [CmdletBinding()]
    [OutputType([String])]
    Param
    (
        # ����
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
            #�Ƃ肠�������͒P���ɋ󔒂��܂ނȂ�"�ň͂�
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
 �V���[�g�J�b�g���쐬
#>
function global:New-Shortcut
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param
    (
        # �^�[�Q�b�g�p�X
        [Parameter(Mandatory=$true, 
                   Position=0)]
        [string]
        $TargetPath,

        # ���O
        [Parameter(Mandatory=$true, 
                   Position=1)]
        [string]
        $Name,

        # �^�C�v
        [Parameter(Mandatory=$true, 
                   Position=2)]
        [ValidateSet('Lnk', 'Url')]
        [string]
        $ShortcutType,

        # �f�B���N�g��
        [Parameter(Mandatory=$false, 
                   Position=3)]
        [string]
        $Directory,

        # �A�C�R���t�@�C���A�A�C�R���C���f�b�N�X
        [Parameter(Mandatory=$false)]
        [string]
        $IconLocation,

        # ����
        [Parameter(Mandatory=$false)]
        [string]
        $Arguments,

        # ��ƃt�H���_�[
        [Parameter(Mandatory=$false)]
        [string]
        $WorkingDirectory,

        # �R�����g
        [Parameter(Mandatory=$false)]
        [string]
        $Description,

        # �V���[�g�J�b�g�L�[
        [Parameter(Mandatory=$false)]
        [string]
        $HotKey,

        # ���s���̑傫��
        [Parameter(Mandatory=$false)]
        [ValidateSet('�ʏ�̃E�B���h�E', '�ŏ���', '�ő剻')]
        [string]
        $WindowStyle = '�ʏ�̃E�B���h�E'
    )

    if($PSCmdlet.ShouldProcess($path, '�V���[�g�J�b�g�쐬')) {
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
                        '�ʏ�̃E�B���h�E' { $Shortcut.WindowStyle = 1 }
                        '�ŏ���' { $Shortcut.WindowStyle = 7 }
                        '�ő剻' { $Shortcut.WindowStyle = 3 }
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
 PowerShell�̃p�X���擾
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
 PowerShell�V���[�g�J�b�g���쐬
#>
function global:New-PowerShellShortcut
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param
    (
        # ���O
        [Parameter(Mandatory=$false, 
                   Position=0)]
        [string]
        $Name,

        # �f�B���N�g��
        [Parameter(Mandatory=$false, 
                   Position=1)]
        [string]
        $Directory,

        # ����
        [Parameter(Mandatory=$false)]
        [string]
        $Arguments,

        # ��ƃt�H���_�[
        [Parameter(Mandatory=$false)]
        [string]
        $WorkingDirectory,

        # �R�����g
        [Parameter(Mandatory=$false)]
        [string]
        $Description,

        # �V���[�g�J�b�g�L�[
        [Parameter(Mandatory=$false)]
        [string]
        $HotKey,

        # ���s���̑傫��
        [Parameter(Mandatory=$false)]
        [ValidateSet('�ʏ�̃E�B���h�E', '�ŏ���', '�ő剻')]
        [string]
        $WindowStyle = '�ʏ�̃E�B���h�E',

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
        New-Shortcut $powerShellPath $Name -ShortcutType Lnk -IconLocation $isePath -Arguments '-ExecutionPolicy RemoteSigned -Command Start-Process powershell_ise' -WindowStyle �ŏ���
    }
}

New-PowerShellShortcut -Verbose
New-PowerShellShortcut -Ise -Verbose