Add-Type -AssemblyName Microsoft.Office.Interop.Excel

<#
 エクセルを起動
#>
function global:Start-Excel
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$false, 
                   Position=0)]
        [switch]
        $Visible,

        [Parameter(Mandatory=$false, 
                   Position=1)]
        [ValidateSet('Minimized', 'Normal', 'Maximized')]
        [string]
        $WindowState
    )

    Begin
    {
    }
    Process
    {
        if ($pscmdlet.ShouldProcess("エクセル", "起動"))
        {
            $excel = New-Object -ComObject Excel.Application
            try {
                if ($Visible.IsPresent) {
                    $excel.Visible = $true
                }
                if ($WindowState) {
                    $excel.WindowState = Invoke-Expression ('[Microsoft.Office.Interop.Excel.XlWindowState]::xl{0}' -f $WindowState)
                }
                $excel
            }
            catch {
                Write-Error $_
            }
        }
    }
    End
    {
    }
}

<#
 エクセルを終了
#>
function global:Stop-Excel
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   Position=0,
                   ValueFromPipeline=$true)]
        [object]
        $Excel
    )

    Begin
    {
    }
    Process
    {
        if ($pscmdlet.ShouldProcess("エクセル", "終了"))
        {
            $Excel | % {
                try {
                    $_.Quit()
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null
                }
                catch {
                    Write-Error $_
                }
            }
        }
    }
    End
    {
    }
}

<#
 ブックを取得
#>
function global:Get-ExcelBook
{
    [CmdletBinding()]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   Position=0,
                   ValueFromPipeline=$true)]
        [object]
        $Excel
    )

    Begin
    {
    }
    Process
    {
        $Excel | % {
            $_.Workbooks
        }
    }
    End
    {
    }
}

<#
 ブックを追加
#>
function global:New-ExcelBook
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   Position=0,
                   ValueFromPipeline=$true)]
        [object]
        $Excel
    )

    Begin
    {
    }
    Process
    {
        if ($pscmdlet.ShouldProcess("エクセル", "ブックを追加"))
        {
            $Excel | % {
                $_.Workbooks.Add()
            }
        }
    }
    End
    {
    }
}

<#
 ブックを開く
#>
function global:Open-ExcelBook
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   Position=0,
                   ValueFromPipeline=$true)]
        [object]
        $Excel,

        [Parameter(Mandatory=$true,
                   Position=1)]
        [string]
        $Filename,

        [Parameter(Mandatory=$false,
                   Position=2)]
        [switch]
        $ReadOnly
    )

    Begin
    {
    }
    Process
    {
        if ($pscmdlet.ShouldProcess("エクセル", "ブックを開く"))
        {
            $Excel | % {
                $_.Workbooks.Open($Filename, 0, $ReadOnly.IsPresent)
            }
        }
    }
    End
    {
    }
}

<#
 ブックを閉じる
#>
function global:Close-ExcelBook
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   Position=0,
                   ValueFromPipeline=$true)]
        [object]
        $Book,

        [Parameter(Mandatory=$false,
                   Position=1)]
        [bool]
        $Savechanges = $false,

        [Parameter(Mandatory=$false,
                   Position=2)]
        [string]
        $FileName = $null
    )

    Begin
    {
    }
    Process
    {
        if ($pscmdlet.ShouldProcess("エクセル", "ブックを閉じる"))
        {
            $Book | % {
                if ($FileName) {
                    $_.Close($Savechanges, $FileName)
                } else {
                    $_.Close($Savechanges)
                }
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($_) | Out-Null
            }
        }
    }
    End
    {
    }
}

<#
 シートを取得
#>
function global:Get-ExcelSheet
{
    [CmdletBinding(DefaultParameterSetName='Sheets')]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   Position=0,
                   ValueFromPipeline=$true)]
        [object]
        $Excel,

        [Parameter(Mandatory=$false, 
                   Position=1)]
        [switch]
        $VisibleOnly,

        [Parameter(Mandatory=$false, 
                   Position=1)]
        [Parameter(ParameterSetName='Index')]
        [ValidateScript({$_ -gt 0})]
        [int]
        $Index,
 
        [Parameter(Mandatory=$false, 
                   Position=1)]
        [Parameter(ParameterSetName='Name')]
        [int]
        $Name
    )

    Begin
    {
    }
    Process
    {
        $Excel | % {
            $i = $_
            switch ($PSCmdlet.ParameterSetName) {
                'Sheets' {
                    if ($VisibleOnly.IsPresent) {
                        $i.Worksheets | ? { $_.Visible -eq [Microsoft.Office.Interop.Excel.XlSheetVisibility]::xlSheetVisible }
                    } else {
                        $i.Worksheets
                    }
                }
                'Index' {
                    $i.Worksheets[$Index]
                }
                'Name' {
                    $i.Worksheets[$Name]
                }
            }
        }
    }
    End
    {
    }
}

<#
 セルを取得
#>
function global:Get-ExcelCell
{
    [CmdletBinding(DefaultParameterSetName='Address')]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   Position=0,
                   ValueFromPipeline=$true)]
        [object]
        $Sheet,

        [Parameter(Mandatory=$false, 
                   Position=1)]
        [Parameter(ParameterSetName='Address')]
        [string]
        $Address,

        [Parameter(Mandatory=$false, 
                   Position=1)]
        [Parameter(ParameterSetName='Index')]
        [int]
        $Row,
 
        [Parameter(Mandatory=$false, 
                   Position=2)]
        [Parameter(ParameterSetName='Index')]
        [int]
        $Column,

        [Parameter(Mandatory=$false, 
                   Position=1)]
        [Parameter(ParameterSetName='UsedRange')]
        [switch]
        $UsedRange
    )

    Begin
    {
    }
    Process
    {
        $Sheet | % {
            $i = $_
            switch ($PSCmdlet.ParameterSetName) {
                'Address' {
                    $i.Range($Address)
                }
                'Index' {
                    $i.Cells($Row, $Column)
                }
                'UsedRange' {
                    $i.UsedRange
                }
            }
        }
    }
    End
    {
    }
}

<#
 図形を取得(グループ再起)
#>
function script:Get-ExcelShapeGroup
{
    [CmdletBinding()]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   Position=0,
                   ValueFromPipeline=$true)]
        [object]
        $GroupItems
    )

    Begin
    {
    }
    Process
    {
        $GroupItems | % {
            $_
            if ($_.GroupItems) {
                $_.GroupItems | Get-ExcelShapeGroup
            }
        }
    }
    End
    {
    }
}

<#
 図形を取得
#>
function global:Get-ExcelShape
{
    [CmdletBinding(DefaultParameterSetName='Address')]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   Position=0,
                   ValueFromPipeline=$true)]
        [object]
        $Sheet
    )

    Begin
    {
    }
    Process
    {
        $Sheet | % {
            $_.Shapes | % {
                $_
                if ($_.GroupItems) {
                    $_.GroupItems | Get-ExcelShapeGroup
                }
            }
        }
    }
    End
    {
    }
}

