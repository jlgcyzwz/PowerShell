Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName WindowsBase

Add-Type -TypeDefinition @'
using System;
using System.Windows.Automation;

public static class UIAutoElement
{
    public static AutomationElement Root
    {
        get
        {
            return AutomationElement.RootElement;
        }
    }

    public static AutomationElement Focused
    {
        get
        {
            return AutomationElement.FocusedElement;
        }
    }

    public static AutomationElement FromHandle(IntPtr hwnd)
    {
        return AutomationElement.FromHandle(hwnd);
    }

    public static AutomationElement FromPoint(System.Windows.Point pt)
    {
        return AutomationElement.FromPoint(pt);
    }

    public static AutomationElement FindFirst(AutomationElement element, TreeScope treeScope, Condition condition)
    {
        return element.FindFirst(treeScope, condition);
    }

    public static AutomationElementCollection FindAll(AutomationElement element, TreeScope treeScope, Condition condition)
    {
        return element.FindAll(treeScope, condition);
    }
}
'@  -ReferencedAssemblies('UIAutomationClient', 'UIAutomationTypes', 'WindowsBase')

<#
 AutomationElement 取得
#>
function Get-AutoElement
{
    [CmdletBinding(DefaultParameterSetName='Find')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   Position=0,
                   ParameterSetName='Find')]
        [System.Windows.Automation.AutomationElement]
        $Element,

        [Parameter(Mandatory=$false, 
                   Position=1,
                   ParameterSetName='Find')]
        [switch]
        $First,

        [Parameter(Mandatory=$false, 
                   Position=2,
                   ParameterSetName='Find')]
        [System.Windows.Automation.ControlType]
        $ControlType = $null,

        [Parameter(Mandatory=$false, 
                   Position=3,
                   ParameterSetName='Find')]
        [string]
        $Name = $null,

        [Parameter(Mandatory=$false, 
                   Position=4,
                   ParameterSetName='Find')]
        [string]
        $AutomationId = $null,

        [Parameter(Mandatory=$false, 
                   Position=5,
                   ParameterSetName='Find')]
        [string]
        $ClassName = $null,

        [Parameter(Mandatory=$false, 
                   Position=6,
                   ParameterSetName='Find')]
        [Int32]
        $NativeWindowHandle = $null,

        [Parameter(Mandatory=$false, 
                   Position=7,
                   ParameterSetName='Find')]
        [Int32]
        $ProcessId = $null,

        [Parameter(Mandatory=$false, 
                   ParameterSetName='Root')]
        [switch]
        $Root,

        [Parameter(Mandatory=$false, 
                   ParameterSetName='Focused')]
        [switch]
        $Focused,

        [Parameter(Mandatory=$false, 
                   ParameterSetName='Handle')]
        [System.IntPtr]
        $Handle,

        [Parameter(Mandatory=$false, 
                   ParameterSetName='Point')]
        [double]
        $X,

        [Parameter(Mandatory=$false, 
                   ParameterSetName='Point')]
        [double]
        $Y
    )

    Begin
    {
    }
    Process
    {
        switch ($PSCmdlet.ParameterSetName) {
            'Find' {
                $condition = $null
                $conditions = @()
                if ($ControlType) {
                    $conditions += New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, $ControlType)
                }
                if ($Name) {
                    $conditions += New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, $Name)
                }
                if ($AutomationId) {
                    $conditions += New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::AutomationIdProperty, $AutomationId)
                }
                if ($ClassName) {
                    $conditions += New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ClassNameProperty, $ClassName)
                }
                if ($NativeWindowHandle) {
                    $conditions += New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NativeWindowHandleProperty, $NativeWindowHandle)
                }
                if ($ProcessId) {
                    $conditions += New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, $ProcessId)
                }
                if ($conditions.Count -eq 0) {
                    $condition = [System.Windows.Automation.Condition]::TrueCondition
                } elseif ($conditions.Count -eq 1) {
                    $condition = $conditions[0]
                }
                $Element | % {
                    $i = if ($_ -eq $null) { [UIAutoElement]::Root } else { $_ }
                    if ($First.IsPresent) {
                        [UIAutoElement]::FindFirst($i, [System.Windows.Automation.TreeScope]::Children, $condition)
                    }
                    else {
                        [UIAutoElement]::FindAll($i, [System.Windows.Automation.TreeScope]::Children, $condition)
                    }
                } | % {
                    $autoElement = $_
                    $autoElement.GetSupportedProperties() | % {
                        if ($_.ProgrammaticName -match '^\w+Identifiers\.(\w+)Property$') {
                            $name = $Matches[1]
                            $value = $autoElement.GetCurrentPropertyValue($_)
                            $autoElement | Add-Member NoteProperty $name $value
                            if ($name -eq 'ControlType') {
                                $autoElement | Add-Member NoteProperty 'ControlTypeName' $value.ProgrammaticName
                            }
                        }
                    }
                    $autoElement
                }
            }
            'Root' {
                [UIAutoElement]::Root
            }
            'Focused' {
                [UIAutoElement]::Focused
            }
            'Handle' {
                [UIAutoElement]::FromHandle($Handle)
            }
            'Point' {
                [UIAutoElement]::FromPoint((New-Object System.Windows.Point($X, $Y)))
            }
        }
    }
    End
    {
    }
}

<#
 パターン取得
#>
function Get-AutoPattern
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([String])]
    Param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   Position=0)]
        [System.Windows.Automation.AutomationElement]
        $Element,

        [Parameter(Mandatory=$false, 
                   Position=1)]
        [ValidateSet('Dock', 'ExpandCollapse', 'GridItem', 'Grid', 'Invoke', 'MultipleView', 'RangeValue', 'Scroll', 'ScrollItem', 'Selection', 'SelectionItem', 'SynchronizedInput', 'Text', 'Transform', 'Toggle', 'Value', 'Window', 'VirtualizedItem', 'ItemContainer')]
        [string]
        $Pattern = $null
    )

    Begin
    {
    }
    Process 
    {
        $Element | % {
            if ($Pattern) {
                $_.GetCurrentPattern((Invoke-Expression ('[System.Windows.Automation.{0}Pattern]::Pattern' -f $Pattern)))
            } else {
                $_.GetSupportedPatterns() | % {
                    if ($_.ProgrammaticName -match '(\w+)PatternIdentifiers.Pattern') {
                        $Matches[1]
                    } else {
                        $_.ProgrammaticName
                    }
                }
            }
        }
    }
    End
    {
    }
}
