<#
 フルパス取得
#>
function global:Get-FullPath
{
    [CmdletBinding()]
    [OutputType([String])]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true)]
        [string]
        $Path
    )

    Begin
    {
    }
    Process
    {
        $Path | % {
            if (Split-Path $_ -IsAbsolute) {
                $_
            } else {
                [System.IO.Path]::GetFullPath($_)
            }
        }
    }
    End
    {
    }
}

<#
 拡張子取得
#>
function global:Get-FullPath
{
    [CmdletBinding()]
    [OutputType([String])]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true)]
        [string]
        $Path
    )

    Begin
    {
    }
    Process
    {
        $Path | % {
            [System.IO.Path]::GetExtension($_)
        }
    }
    End
    {
    }
}

<#
 拡張子を除いたファイル名取得
#>
function global:Get-FullPath
{
    [CmdletBinding()]
    [OutputType([String])]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true)]
        [string]
        $Path
    )

    Begin
    {
    }
    Process
    {
        $Path | % {
            [System.IO.Path]::GetFileNameWithoutExtension($_)
        }
    }
    End
    {
    }
}
