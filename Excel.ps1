Add-Type -AssemblyName Microsoft.Office.Interop.Excel
Add-Type -AssemblyName Office

$Missing = [System.Type]::Missing

$MsoTriState = [Microsoft.Office.Core.MsoTriState]
$MsoShapeType = [Microsoft.Office.Core.MsoShapeType]
$MsoAutoShapeType = [Microsoft.Office.Core.MsoAutoShapeType]

$XlDirection = [Microsoft.Office.Interop.Excel.XlDirection]

<#
 boolをMsoTriStateに変換
#>
function global:ConvertTo-MsoTriState
{
    [CmdletBinding()]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true)]
        [bool]
        $Value
    )

    Begin
    {
    }
    Process
    {
        try {
            $Value | % {
                if ($Value) {
                    $MsoTriState::msoTrue
                } else {
                    $MsoTriState::msoFalse
                }
            }
        }
        catch {
            Write-Error $_
            Write-Error $_.ScriptStackTrace
        }
    }
    End
    {
    }
}


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
                    $excel.Visible = [Microsoft.Office.Interop.Excel.XlSheetVisibility]::xlSheetVisible
                }
                if ($WindowState) {
                    $excel.WindowState = Invoke-Expression ('[Microsoft.Office.Interop.Excel.XlWindowState]::xl{0}' -f $WindowState)
                }
                $excel
            }
            catch {
                Write-Error $_
                Write-Error $_.ScriptStackTrace
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
                    Write-Error $_.ScriptStackTrace
                }
            }
        }
    }
    End
    {
    }
}

<#
 ワークブックを取得
#>
function global:Get-ExcelWorkbook
{
    [CmdletBinding()]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true)]
        [object]
        $Excel
    )

    Begin
    {
    }
    Process
    {
        try {
            $Excel | % {
                $_.Workbooks
            }
        }
        catch {
            Write-Error $_
            Write-Error $_.ScriptStackTrace
        }
    }
    End
    {
    }
}

<#
 ワークブックを追加
#>
function global:New-ExcelWorkbook
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true)]
        [object]
        $Excel
    )

    Begin
    {
    }
    Process
    {
        try {
            if ($pscmdlet.ShouldProcess("エクセル", "ワークブックを追加"))
            {
                $Excel | % {
                    $_.Workbooks.Add()
                }
            }
        }
        catch {
            Write-Error $_
            Write-Error $_.ScriptStackTrace
        }
    }
    End
    {
    }
}

<#
 ワークブックを開く
#>
function global:Open-ExcelWorkbook
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true)]
        [object]
        $Excel,

        [Parameter(Mandatory=$true,
                   Position=0)]
        [string]
        $Filename,

        [Parameter(Mandatory=$false,
                   Position=1)]
        [switch]
        $ReadOnly
    )

    Begin
    {
    }
    Process
    {
        try {
            if ($pscmdlet.ShouldProcess("エクセル", "ワークブックを開く"))
            {
                $Excel | % {
                    $path = if ((Split-Path $FileName -IsAbsolute)) {
                        $FileName
                    } else {
                        Join-Path (Split-Path $FileName -Parent -Resolve) $FileName
                    }
                    $_.Workbooks.Open($path, 0, $ReadOnly.IsPresent)
                }
            }
        }
        catch {
            Write-Error $_
            Write-Error $_.ScriptStackTrace
        }
    }
    End
    {
    }
}

<#
 ワークブックを閉じる
#>
function global:Close-ExcelWorkbook
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true)]
        [object]
        $Book,

        [Parameter(Mandatory=$false,
                   Position=0)]
        [Switch]
        $Savechanges,

        [Parameter(Mandatory=$false,
                   Position=1)]
        [string]
        $FileName = $null
    )

    Begin
    {
    }
    Process
    {
        try {
            if ($pscmdlet.ShouldProcess("エクセル", "ワークブックを閉じる"))
            {
                $Book | % {
                    if ($FileName) {
                        $path = if ((Split-Path $FileName -IsAbsolute)) {
                            $FileName
                        } else {
                            Join-Path (Split-Path $FileName -Resolve) $FileName
                        }
                        $_.Close($Savechanges.IsPresent, $path)
                    } else {
                        $_.Close($Savechanges.IsPresent)
                    }
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($_) | Out-Null
                }
            }
        }
        catch {
            Write-Error $_
            Write-Error $_.ScriptStackTrace
        }
    }
    End
    {
    }
}

<#
 ワークシートを取得
#>
function global:Get-ExcelWorksheet
{
    [CmdletBinding(DefaultParameterSetName='Sheets')]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true)]
        [object]
        $Book,

        [Parameter(Mandatory=$false, 
                   Position=0)]
        [Parameter(ParameterSetName='Sheets')]
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
        [string]
        $Name
    )

    Begin
    {
    }
    Process
    {
        try {
            $Book | % {
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
        catch {
            Write-Error $_
            Write-Error $_.ScriptStackTrace
        }
    }
    End
    {
    }
}

<#
 ワークシートを追加
#>
function global:New-ExcelWorksheet
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true)]
        [object]
        $Book,

        [Parameter(Mandatory=$false)]
        [object]
        $Before,

        [Parameter(Mandatory=$false)]
        [object]
        $After,

        [Parameter(Mandatory=$false)]
        [int]
        $Count,

        [Parameter(Mandatory=$false)]
        [Microsoft.Office.Interop.Excel.XlSheetType]
        $Type
    )

    Begin
    {
    }
    Process
    {
        try {
            if ($pscmdlet.ShouldProcess($Book.Name, "ワークシートを追加")) {
                $Book | % {
                    $p1 = if ($Before) { $Before } else { $Missing }
                    $p2 = if ($After) { $After } else { $Missing }
                    $p3 = if ($Count) { $Count } else { $Missing }
                    $p4 = if ($Type) { $Type } else { $Missing }
                    $_.Worksheets.Add($p1, $p2, $p3, $p4)
                }
            }
        }
        catch {
            Write-Error $_
            Write-Error $_.ScriptStackTrace
        }
    }
    End
    {
    }
}

<#
 ワークシートを移動
#>
function global:Move-ExcelWorksheet
{
    [CmdletBinding(SupportsShouldProcess=$true,
                   DefaultParameterSetName='Before')]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true)]
        [object]
        $Sheet,

        [Parameter(Mandatory=$false, 
                   Position=0)]
        [Parameter(ParameterSetName='Before')]
        [object]
        $Before,

        [Parameter(Mandatory=$false, 
                   Position=0)]
        [Parameter(ParameterSetName='After')]
        [object]
        $After
    )

    Begin
    {
    }
    Process
    {
        try {
            if ($pscmdlet.ShouldProcess($Book.Name, "ワークシートを移動")) {
                switch ($pscmdlet.ParameterSetName) {
                    'Before' {
                        $Sheet | % {
                            $_.Move($Before)
                        }
                    }
                    'After' {
                        $Sheet | % {
                            $_.Move($Missing, $After)
                        }
                    }
                }
            }
        }
        catch {
            Write-Error $_
            Write-Error $_.ScriptStackTrace
        }
    }
    End
    {
    }
}

<#
 ワークシートをコピー
#>
function global:Copy-ExcelWorksheet
{
    [CmdletBinding(SupportsShouldProcess=$true,
                   DefaultParameterSetName='Before')]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true)]
        [object]
        $Sheet,

        [Parameter(Mandatory=$false, 
                   Position=0)]
        [Parameter(ParameterSetName='Before')]
        [object]
        $Before,

        [Parameter(Mandatory=$false, 
                   Position=0)]
        [Parameter(ParameterSetName='After')]
        [object]
        $After
    )

    Begin
    {
    }
    Process
    {
        try {
            if ($pscmdlet.ShouldProcess($Book.Name, "ワークシートをコピー")) {
                switch ($pscmdlet.ParameterSetName) {
                    'Before' {
                        $Sheet | % {
                            $_.Copy($Before)
                        }
                    }
                    'After' {
                        $Sheet | % {
                            $_.Copy($Missing, $After)
                        }
                    }
                }
            }
        }
        catch {
            Write-Error $_
            Write-Error $_.ScriptStackTrace
        }
    }
    End
    {
    }
}

<#
 グループ内の図形を取得
#>
function script:Get-ExcelShapeGroupItem
{
    [CmdletBinding()]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true)]
        [object]
        $Shape,

        [Parameter(Mandatory=$false, 
                   Position=0)]
        [switch]
        $Recurse
    )

    Begin
    {
    }
    Process
    {
        try {
            $Shape | % {
                if ($_.GroupItems) {
                    $_.GroupItems | % {
                        $_
                        if ($Recurse.IsPresent) {
                            $_ | Get-ExcelShapeGroupItem
                        }
                    }
                }
            }
        }
        catch {
            Write-Error $_
            Write-Error $_.ScriptStackTrace
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
    [CmdletBinding()]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true)]
        [object]
        $Sheet,

        [Parameter(Mandatory=$false, 
                   Position=0)]
        [switch]
        $RecurseGroup
    )

    Begin
    {
    }
    Process
    {
        try {
            $Sheet | % {
                $_.Shapes | % {
                    $_
                    if ($RecurseGroup.IsPresent) {
                        $_ | Get-ExcelShapeGroupItem -Recurse
                    }
                }
            }
        }
        catch {
            Write-Error $_
            Write-Error $_.ScriptStackTrace
        }
    }
    End
    {
    }
}

<#
 図形を追加
#>
function global:New-ExcelShape
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true)]
        [object]
        $Sheet,

        [Parameter(Mandatory=$true, 
                   Position=0)]
        [Microsoft.Office.Core.MsoAutoShapeType]
        $Type,

        [Parameter(Mandatory=$true, 
                   Position=1)]
        [float]
        $Left,

        [Parameter(Mandatory=$true, 
                   Position=2)]
        [float]
        $Top,

        [Parameter(Mandatory=$true, 
                   Position=3)]
        [float]
        $Width,

        [Parameter(Mandatory=$true, 
                   Position=4)]
        [float]
        $Height
    )

    Begin
    {
    }
    Process
    {
        try {
            if ($pscmdlet.ShouldProcess($Sheet.Name, "図形を追加")) {
                $Sheet | % {
                    $_.Shapes.AddShape($Type, $Left, $Top ,$Width, $Height)
                }
            }
        }
        catch {
            Write-Error $_
            Write-Error $_.ScriptStackTrace
        }
    }
    End
    {
    }
}

<#
 画像を追加
#>
function global:New-ExcelPicture
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true)]
        [object]
        $Sheet,

        [Parameter(Mandatory=$true, 
                   Position=0)]
        [string]
        $FileName,

        [Parameter(Mandatory=$true, 
                   Position=1)]
        [bool]
        $LinkToFile,

        [Parameter(Mandatory=$true, 
                   Position=2)]
        [bool]
        $SaveWithDocument,

        [Parameter(Mandatory=$true, 
                   Position=3)]
        [float]
        $Left,

        [Parameter(Mandatory=$true, 
                   Position=4)]
        [float]
        $Top,

        [Parameter(Mandatory=$false, 
                   Position=5)]
        [float]
        $Width = -1,

        [Parameter(Mandatory=$false, 
                   Position=6)]
        [float]
        $Height = -1
    )

    Begin
    {
    }
    Process
    {
        try {
            if ($pscmdlet.ShouldProcess($Sheet.Name, "画像を追加")) {
                $Sheet | % {
                    $path = if ((Split-Path $FileName -IsAbsolute)) {
                        $FileName
                    } else {
                        Join-Path (Split-Path $FileName -Parent -Resolve) $FileName
                    }
                    $_.Shapes.AddPicture($path, ($LinkToFile | ConvertTo-MsoTriState), ($SaveWithDocument | ConvertTo-MsoTriState), $Left, $Top, $Width, $Height)
                }
            }
        }
        catch {
            Write-Error $_
            Write-Error $_.ScriptStackTrace
        }
    }
    End
    {
    }
}

<#
 図形のTextRange取得
#>
function global:Get-ExcelShapeTextRange
{
    [CmdletBinding(DefaultParameterSetName='Address')]
    [OutputType([object])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true)]
        [object]
        $Shape
    )

    Begin
    {
    }
    Process
    {
        try {
            $Shape | % {
                if ($_.TextFrame2.HasText -eq $MsoTriState::msoTrue) {
                    $_.TextFrame2.TextRange
                }
            }
        }
        catch {
            Write-Error $_
            Write-Error $_.ScriptStackTrace
        }
    }
    End
    {
    }
}

Get-IseSnippet | ? Name -like 'Excel(ReadOnly).*' | Remove-Item
New-IseSnippet -Title 'Excel(ReadOnly)' -Description 'エクセルブックを読み取り専用で開く' -Text @'
Start-Excel -Visible -WindowState Maximized | % {
    try {
        $_ | Open-ExcelWorkbook '' -ReadOnly | % {
            try {
                $_ | Get-ExcelWorksheet | % {
                }
            }
            finally {
                $_ | Close-ExcelWorkbook
            }
        }
    }
    catch {
        Write-Error $_
        Write-Error $_.ScriptStackTrace
    }
    finally {
        $_ | Stop-Excel
    }
}
'@ -Force
