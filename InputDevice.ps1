Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

/// <summary>
/// 入力送信クラス
/// </summary>
public class InputSender
{
    #region SendInput

    /// <summary>
    /// SendInput
    /// </summary>
    /// <param name="cInputs"></param>
    /// <param name="pInputs"></param>
    /// <param name="cbSize"></param>
    /// <returns></returns>
    [DllImport("user32.dll")]
    private static extern uint SendInput(uint cInputs, INPUT[] pInputs, int cbSize);

    #region INPUT

    /// <summary>
    /// INPUT
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct INPUT
    {
        public uint Type;
        public INPUT_UNION iu;
    }

    // Type
    public const uint INPUT_MOUSE = 0;
    public const uint INPUT_KEYBOARD = 1;

    /// <summary>
    /// INPUT_UNION
    /// </summary>
    [StructLayout(LayoutKind.Explicit)]
    public struct INPUT_UNION
    {
        [FieldOffset(0)]
        public MOUSEINPUT mi;

        [FieldOffset(0)]
        public KEYBDINPUT ki;
    }

    #endregion INPUT

    #region MOUSEINPUT

    /// <summary>
    /// MOUSEINPUT
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct MOUSEINPUT
    {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public UIntPtr dwExtraInfo;
    }

    // mouseData
    public const uint XBUTTON1 = 0x0001;
    public const uint XBUTTON2 = 0x0002;

    public const int WHEEL_DELTA = 120;

    // dwFlags
    public const uint MOUSEEVENTF_MOVE = 0x0001;
    public const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    public const uint MOUSEEVENTF_LEFTUP = 0x0004;
    public const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
    public const uint MOUSEEVENTF_RIGHTUP = 0x0010;
    public const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020;
    public const uint MOUSEEVENTF_MIDDLEUP = 0x0040;
    public const uint MOUSEEVENTF_XDOWN = 0x0080;
    public const uint MOUSEEVENTF_XUP = 0x0100;
    public const uint MOUSEEVENTF_WHEEL = 0x0800;
    public const uint MOUSEEVENTF_HWHEEL = 0x1000;
    public const uint MOUSEEVENTF_MOVE_NOCOALESCE = 0x2000;
    public const uint MOUSEEVENTF_VIRTUALDESK = 0x4000;
    public const uint MOUSEEVENTF_ABSOLUTE = 0x8000;

    #endregion MOUSEINPUT

    #region KEYBDINPUT

    /// <summary>
    /// KEYBDINPUT
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct KEYBDINPUT
    {
        public short wVk;
        public short wScan;
        public uint dwFlags;
        public uint time;
        public UIntPtr dwExtraInfo;
    }

    // dwFlags
    public const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
    public const uint KEYEVENTF_KEYUP = 0x0002;
    public const uint KEYEVENTF_SCANCODE = 0x0008;
    public const uint KEYEVENTF_UNICODE = 0x0004;

    #endregion KEYBDINPUT

    #endregion SendInput

    /// <summary>
    /// マウスボタン
    /// </summary>
    public enum MouseButton
    {
        Left,
        Middle,
        Right
    }

    /// <summary>
    /// マウス操作
    /// </summary>
    public enum MouseAction
    {
        Down,
        Up,
        Click,
        DoubleClick
    }

    /// <summary>
    /// マウスホイールの方向
    /// </summary>
    public enum MouseWheelDirection
    {
        Vertical,
        Horizontal
    }

    /// <summary>
    /// キー操作
    /// </summary>
    public enum KeyAction
    {
        Down,
        Up,
        Stroke
    }

    /// <summary>
    /// マウス操作
    /// </summary>
    /// <param name="button">ボタン</param>
    /// <param name="action">操作</param>
    /// <returns></returns>
    public static uint Send(MouseButton button, MouseAction action)
    {
        uint down = 0;
        uint up = 0;
        switch (button)
        {
            case MouseButton.Left:
                down = MOUSEEVENTF_LEFTDOWN;
                up = MOUSEEVENTF_LEFTUP;
                break;
            case MouseButton.Middle:
                down = MOUSEEVENTF_MIDDLEDOWN;
                up = MOUSEEVENTF_MIDDLEUP;
                break;
            case MouseButton.Right:
                down = MOUSEEVENTF_RIGHTDOWN;
                up = MOUSEEVENTF_RIGHTUP;
                break;
            default:
                throw new ArgumentOutOfRangeException(button.ToString());
        }

        INPUT[] inputs = null;
        switch (action)
        {
            case MouseAction.Down:
                inputs = new INPUT[1];
                inputs[0].iu.mi.dwFlags = down;
                break;
            case MouseAction.Up:
                inputs = new INPUT[1];
                inputs[0].iu.mi.dwFlags = up;
                break;
            case MouseAction.Click:
                inputs = new INPUT[2];
                inputs[0].iu.mi.dwFlags = down;
                inputs[1].iu.mi.dwFlags = up;
                break;
            case MouseAction.DoubleClick:
                inputs = new INPUT[4];
                inputs[0].iu.mi.dwFlags = down;
                inputs[1].iu.mi.dwFlags = up;
                inputs[2].iu.mi.dwFlags = down;
                inputs[3].iu.mi.dwFlags = up;
                break;
            default:
                throw new ArgumentOutOfRangeException(action.ToString());
        }
        for (int i = 0; i < inputs.Length; i++)
        {
            inputs[i].Type = INPUT_MOUSE;
            inputs[i].iu.mi.dx = 0;
            inputs[i].iu.mi.dy = 0;
            inputs[i].iu.mi.mouseData = 0;
            inputs[i].iu.mi.time = 0;
            inputs[i].iu.mi.dwExtraInfo = UIntPtr.Zero;
        }

        return Send(inputs);
    }

    /// <summary>
    /// マウスホイール操作
    /// </summary>
    /// <param name="direction"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public static uint Send(MouseWheelDirection direction, double value)
    {
        INPUT[] inputs = new INPUT[1];
        inputs[0].Type = INPUT_MOUSE;
        inputs[0].iu.mi.dx = 0;
        inputs[0].iu.mi.dy = 0;
        inputs[0].iu.mi.mouseData = (uint)Convert.ToInt32(Math.Round((double)WHEEL_DELTA * value, MidpointRounding.AwayFromZero));
        switch (direction)
        {
            case MouseWheelDirection.Vertical:
                inputs[0].iu.mi.dwFlags = MOUSEEVENTF_WHEEL;
                break;
            case MouseWheelDirection.Horizontal:
                inputs[0].iu.mi.dwFlags = MOUSEEVENTF_HWHEEL;
                break;
            default:
                throw new ArgumentOutOfRangeException(direction.ToString());
        }
        inputs[0].iu.mi.time = 0;
        inputs[0].iu.mi.dwExtraInfo = UIntPtr.Zero;

        return Send(inputs);
    }

    /// <summary>
    /// キー操作
    /// </summary>
    /// <param name="key"></param>
    /// <param name="action"></param>
    /// <returns></returns>
    public static uint Send(short key, KeyAction action)
    {
        INPUT[] inputs = null;
        switch (action)
        {
            case KeyAction.Down:
                inputs = new INPUT[1];
                inputs[0].iu.ki.dwFlags = 0;
                break;
            case KeyAction.Up:
                inputs = new INPUT[1];
                inputs[0].iu.ki.dwFlags = KEYEVENTF_KEYUP;
                break;
            case KeyAction.Stroke:
                inputs = new INPUT[2];
                inputs[0].iu.ki.dwFlags = 0;
                inputs[1].iu.ki.dwFlags = KEYEVENTF_KEYUP;
                break;
            default:
                throw new ArgumentOutOfRangeException(action.ToString());
        }
        for (int i = 0; i < inputs.Length; i++)
        {
            inputs[i].Type = INPUT_KEYBOARD;
            inputs[i].iu.ki.wVk = key;
            inputs[i].iu.ki.wScan = 0;
            inputs[i].iu.ki.time = 0;
            inputs[i].iu.ki.dwExtraInfo = UIntPtr.Zero;
        }

        return Send(inputs);
    }

    /// <summary>
    /// Send
    /// </summary>
    /// <param name="inputs"></param>
    /// <returns>SendInputの戻り値</returns>
    public static uint Send(INPUT[] inputs)
    {
        if (inputs != null)
        {
            return SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }
        return 0;
    }
}
'@

<#
 マウスボタンクリック
#>
function global:Click-Mouse
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter(Mandatory=$false, 
                   Position=0)]
        [ValidateSet('Left', 'Middle', 'Right')]
        $Button = 'Left',

        [Parameter(Mandatory=$false,
                   Position=1)]
        [int]
        $X = [System.Windows.Forms.Cursor]::Position.X,

        [Parameter(Mandatory=$false,
                   Position=2)]
        [int]
        $Y = [System.Windows.Forms.Cursor]::Position.Y,

        [Parameter(Mandatory=$false)]
        [Switch]
        $Double
    )
    
    if ($pscmdlet.ShouldProcess(('位置 {0},{1}' -f $X, $Y), ('{0} click' -f $Button)))
    {
        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($X, $Y)
        $action = if ($Double.IsPresent) { [InputSender+MouseAction]::DoubleClick } else { [InputSender+MouseAction]::Click }
        Switch ($Button) {
            'Left' {
                [InputSender]::Send([InputSender+MouseButton]::Left, $action) | Out-Null
            }
            'Middle' {
                [InputSender]::Send([InputSender+MouseButton]::Middle, $action) | Out-Null
            }
            'Right' {
                [InputSender]::Send([InputSender+MouseButton]::Right, $action) | Out-Null
            }
        }
    }
}

<#
 マウスホイール操作
#>
function global:Invoke-MouseWheel
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter(Mandatory=$true, 
                   Position=0)]
        [double]
        $Value,

        [Parameter(Mandatory=$false, 
                   Position=1)]
        [ValidateSet('Virtical', 'Horizontal')]
        $Direction = 'Virtical'
    )

    if ($pscmdlet.ShouldProcess(('{0}方向 {1}の量' -f $Direction, $Value), "マウスホイール操作"))
    {
        Switch ($Direction) {
            'Virtical' {
                [InputSender]::Send([InputSender+MouseWheelDirection]::Vertical, $Value) | Out-Null
            }
            'Horizontal' {
                [InputSender]::Send([InputSender+MouseWheelDirection]::Horizontal, $Value) | Out-Null
            }
        }
    }
}

<#
 送信キー取得
#>
function global:Get-SendKeys
{
    [CmdletBinding(DefaultParameterSetName='Key')]
    
    Param
    (
        [Parameter(Mandatory=$false, 
                   Position=0,
                   ParameterSetName='Key')]
        [ValidateLength(1, 1)]
        [string]
        $Key,

        [Parameter(Mandatory=$false, 
                   Position=0,
                   ParameterSetName='ActionKey')]
        [ValidateSet('BACKSPACE', 'BREAK', 'CAPSLOCK', 'DELETE', 'DOWN', 'END', 'ENTER', 'ESC', 'HELP', 'HOME', 'INSERT', 'LEFT', 'NUMLOCK', 'PGDN', 'PGUP', 'PRTSC', 'RIGHT', 'SCROLLLOCK', 'TAB', 'UP',
                     'F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8', 'F9', 'F10', 'F11', 'F12', 'F13', 'F14', 'F15', 'F16',
                     'ADD', 'SUBTRACT', 'MULTIPLY', 'DIVIDE')]
        [string]
        $ActionKey,

        [Parameter(Mandatory=$false, 
                   Position=0,
                   ParameterSetName='Keys')]
        [string]
        $Keys,

        [Parameter(Mandatory=$false, 
                   Position=1)]
        [Parameter(ParameterSetName='Key')]
        [Parameter(ParameterSetName='ActionKey')]
        [ValidateRange(1, [int]::MaxValue)]
        [int]
        $RepeatCount = 1,

        [Parameter(Mandatory=$false)]
        [Switch]
        $Ctrl,

        [Parameter(Mandatory=$false)]
        [Switch]
        $Shift,

        [Parameter(Mandatory=$false)]
        [Switch]
        $Alt,

        [Parameter(Mandatory=$false,
                    ParameterSetName='Keys')]
        [Switch]
        $Enclose
    )

    $str = switch($PSCmdlet.ParameterSetName) {
        'Key' {
            if ('+^%(){}[]'.Contains($Key)) {
                if ($RepeatCount -ge 2) {
                    '{{{0} {1}}}' -f $Key, $RepeatCount
                } else {
                    '{{{0}}}' -f $Key
                }
            } else {
                if ($RepeatCount -ge 2) {
                    '{{{0} {1}}}' -f $Key, $RepeatCount
                } else {
                    $Key
                }
            }
        }
        'ActionKey' {
            if ($RepeatCount -ge 2) {
                '{{{0} {1}}}' -f $ActionKey, $RepeatCount
            } else {
                '{{{0}}}' -f $ActionKey
            }
        }
        'keys' {
            ($Keys.ToCharArray() | % {
                $Key = $_.ToString()
                if ('+^%(){}[]'.Contains($Key)) {
                    '{{{0}}}' -f $Key
                } else {
                    $Key
                }
            }) -join ''
        }
    }
    $modifierKey = ''
    if ($Ctrl.IsPresent) {
        $modifierKey += '^'
    }
    if ($Shift.IsPresent) {
        $modifierKey += '+'
    }
    if ($Alt.IsPresent) {
        $modifierKey += '%'
    }
    if (![string]::IsNullOrEmpty($modifierKey)) {
        if ($Enclose.IsPresent) {
            '{0}({1})' -f $modifierKey, $str
        } else {
            '{0}{1}' -f $modifierKey, $str
        }
    } else {
        $str
    }
}

<#
 キー送信
#>
function global:Send-Keys
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter(Mandatory=$true, 
                   Position=0)]
        [string]
        $Keys
    )

    if ($pscmdlet.ShouldProcess('アクティブなアプリケーション', ('{0}を送信' -f $Keys)))
    {
        [System.Windows.Forms.SendKeys]::SendWait($Keys)
    }
}

<#
 Winキー
#>
function global:Send-WinKey
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter(Mandatory=$true, 
                   Position=0)]
        [System.Windows.Forms.Keys]
        $Key,

        [Parameter(Mandatory=$false)]
        [switch]
        $Ctrl,

        [Parameter(Mandatory=$false)]
        [switch]
        $Shift,

        [Parameter(Mandatory=$false)]
        [switch]
        $Alt
    )

    $operation = 'Win + '
    if ($Ctrl.IsPresent) {
        $operation += 'Ctrl + '
    }
    if ($Shift.IsPresent) {
        $operation += 'Shift + '
    }
    if ($Alt.IsPresent) {
        $operation += 'Alt + '
    }
    $operation += $Key

    if ($pscmdlet.ShouldProcess('アクティブなアプリケーション', $operation))
    {
        [InputSender]::Send([Int16][System.Windows.Forms.Keys]::LWin, [InputSender+KeyAction]::Down) | Out-Null
        if ($Ctrl.IsPresent) {
            [InputSender]::Send([Int16][System.Windows.Forms.Keys]::LControlKey, [InputSender+KeyAction]::Down) | Out-Null
        }
        if ($Shift.IsPresent) {
            [InputSender]::Send([Int16][System.Windows.Forms.Keys]::LShiftKey, [InputSender+KeyAction]::Down) | Out-Null
        }
        if ($Alt.IsPresent) {
            [InputSender]::Send([Int16][System.Windows.Forms.Keys]::LMenu, [InputSender+KeyAction]::Down) | Out-Null
        }
        [InputSender]::Send([Int16]$Key, [InputSender+KeyAction]::Stroke) | Out-Null
        if ($Alt.IsPresent) {
            [InputSender]::Send([Int16][System.Windows.Forms.Keys]::LMenu, [InputSender+KeyAction]::Up) | Out-Null
        }
        if ($Shift.IsPresent) {
            [InputSender]::Send([Int16][System.Windows.Forms.Keys]::LShiftKey, [InputSender+KeyAction]::Up) | Out-Null
        }
        if ($Ctrl.IsPresent) {
            [InputSender]::Send([Int16][System.Windows.Forms.Keys]::LControlKey, [InputSender+KeyAction]::Up) | Out-Null
        }
        [InputSender]::Send([Int16][System.Windows.Forms.Keys]::LWin, [InputSender+KeyAction]::Up) | Out-Null
    }
}

<#
 アプリケーション キー
#>
function global:Send-AppsKey
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter(Mandatory=$false)]
        [switch]
        $Ctrl,

        [Parameter(Mandatory=$false)]
        [switch]
        $Shift,

        [Parameter(Mandatory=$false)]
        [switch]
        $Alt
    )

    $operation = ''
    if ($Ctrl.IsPresent) {
        $operation += 'Ctrl + '
    }
    if ($Shift.IsPresent) {
        $operation += 'Shift + '
    }
    if ($Alt.IsPresent) {
        $operation += 'Alt + '
    }
    $operation += 'Apps'

    if ($pscmdlet.ShouldProcess('アクティブなアプリケーション', $operation))
    {
        [InputSender]::Send([Int16][System.Windows.Forms.Keys]::LWin, [InputSender+KeyAction]::Down) | Out-Null
        if ($Ctrl.IsPresent) {
            [InputSender]::Send([Int16][System.Windows.Forms.Keys]::LControlKey, [InputSender+KeyAction]::Down) | Out-Null
        }
        if ($Shift.IsPresent) {
            [InputSender]::Send([Int16][System.Windows.Forms.Keys]::LShiftKey, [InputSender+KeyAction]::Down) | Out-Null
        }
        if ($Alt.IsPresent) {
            [InputSender]::Send([Int16][System.Windows.Forms.Keys]::LMenu, [InputSender+KeyAction]::Down) | Out-Null
        }
        [InputSender]::Send([Int16][System.Windows.Forms.Keys]::Apps, [InputSender+KeyAction]::Stroke) | Out-Null
        if ($Alt.IsPresent) {
            [InputSender]::Send([Int16][System.Windows.Forms.Keys]::LMenu, [InputSender+KeyAction]::Up) | Out-Null
        }
        if ($Shift.IsPresent) {
            [InputSender]::Send([Int16][System.Windows.Forms.Keys]::LShiftKey, [InputSender+KeyAction]::Up) | Out-Null
        }
        if ($Ctrl.IsPresent) {
            [InputSender]::Send([Int16][System.Windows.Forms.Keys]::LControlKey, [InputSender+KeyAction]::Up) | Out-Null
        }
        [InputSender]::Send([Int16][System.Windows.Forms.Keys]::LWin, [InputSender+KeyAction]::Up) | Out-Null
    }
}
