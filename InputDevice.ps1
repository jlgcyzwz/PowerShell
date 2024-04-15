Add-Type -AssemblyName System.Windows.Forms

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class InputSender
{
    [DllImport("user32.dll")]
    private static extern uint SendInput(uint cInputs, INPUT[] pInputs, int cbSize);

    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT
    {
        public uint Type;
        public INPUT_UNION iu;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct INPUT_UNION
    {
        [FieldOffset(0)]
        public MOUSEINPUT mi;
        [FieldOffset(0)]
        public KEYBDINPUT ki;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MOUSEINPUT
    {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public UIntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KEYBDINPUT
    {
        public short wVk;
        public short wScan;
        public uint dwFlags;
        public uint time;
        public UIntPtr dwExtraInfo;
    }

    // INPUT
    // Type
    private const uint INPUT_MOUSE = 0;
    private const uint INPUT_KEYBOARD = 1;

    // MOUSEINPUT
    // mouseData
    private const uint XBUTTON1 = 0x0001;
    private const uint XBUTTON2 = 0x0002;
    
    // dwFlags
    private const uint MOUSEEVENTF_MOVE = 0x0001;
    private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    private const uint MOUSEEVENTF_LEFTUP = 0x0004;
    private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
    private const uint MOUSEEVENTF_RIGHTUP = 0x0010;
    private const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020;
    private const uint MOUSEEVENTF_MIDDLEUP = 0x0040;
    private const uint MOUSEEVENTF_XDOWN = 0x0080;
    private const uint MOUSEEVENTF_XUP = 0x0100;
    private const uint MOUSEEVENTF_WHEEL = 0x0800;
    private const uint MOUSEEVENTF_HWHEEL = 0x1000;
    private const uint MOUSEEVENTF_MOVE_NOCOALESCE = 0x2000;
    private const uint MOUSEEVENTF_VIRTUALDESK = 0x4000;
    private const uint MOUSEEVENTF_ABSOLUTE = 0x8000;

    // KEYBDINPUT
    // mouseData
    private const int WHEEL_DELTA = 120;
    // dwFlags
    private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
    private const uint KEYEVENTF_KEYUP = 0x0002;
    private const uint KEYEVENTF_SCANCODE = 0x0008;
    private const uint KEYEVENTF_UNICODE = 0x0004;

    public static uint MouseLeftDown()
    {
        INPUT[] input = new INPUT[1];
        input[0].Type = INPUT_MOUSE;
        input[0].iu.mi.dx = 0;
        input[0].iu.mi.dy = 0;
        input[0].iu.mi.mouseData = 0;
        input[0].iu.mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
        input[0].iu.mi.time = 0;
        input[0].iu.mi.dwExtraInfo = UIntPtr.Zero;
        return SendInput(1, input, Marshal.SizeOf(typeof(INPUT)));
    }

    public static uint MouseLeftUp()
    {
        INPUT[] input = new INPUT[1];
        input[0].Type = INPUT_MOUSE;
        input[0].iu.mi.dx = 0;
        input[0].iu.mi.dy = 0;
        input[0].iu.mi.mouseData = 0;
        input[0].iu.mi.dwFlags = MOUSEEVENTF_LEFTUP;
        input[0].iu.mi.time = 0;
        input[0].iu.mi.dwExtraInfo = UIntPtr.Zero;
        return SendInput(1, input, Marshal.SizeOf(typeof(INPUT)));
    }

    public static uint MouseRightDown()
    {
        INPUT[] input = new INPUT[1];
        input[0].Type = INPUT_MOUSE;
        input[0].iu.mi.dx = 0;
        input[0].iu.mi.dy = 0;
        input[0].iu.mi.mouseData = 0;
        input[0].iu.mi.dwFlags = MOUSEEVENTF_RIGHTDOWN;
        input[0].iu.mi.time = 0;
        input[0].iu.mi.dwExtraInfo = UIntPtr.Zero;
        return SendInput(1, input, Marshal.SizeOf(typeof(INPUT)));
    }

    public static uint MouseRightUp()
    {
        INPUT[] input = new INPUT[1];
        input[0].Type = INPUT_MOUSE;
        input[0].iu.mi.dx = 0;
        input[0].iu.mi.dy = 0;
        input[0].iu.mi.mouseData = 0;
        input[0].iu.mi.dwFlags = MOUSEEVENTF_RIGHTUP;
        input[0].iu.mi.time = 0;
        input[0].iu.mi.dwExtraInfo = UIntPtr.Zero;
        return SendInput(1, input, Marshal.SizeOf(typeof(INPUT)));
    }

    public static uint MouseMiddleDown()
    {
        INPUT[] input = new INPUT[1];
        input[0].Type = INPUT_MOUSE;
        input[0].iu.mi.dx = 0;
        input[0].iu.mi.dy = 0;
        input[0].iu.mi.mouseData = 0;
        input[0].iu.mi.dwFlags = MOUSEEVENTF_MIDDLEDOWN;
        input[0].iu.mi.time = 0;
        input[0].iu.mi.dwExtraInfo = UIntPtr.Zero;
        return SendInput(1, input, Marshal.SizeOf(typeof(INPUT)));
    }

    public static uint MouseMiddleUp()
    {
        INPUT[] input = new INPUT[1];
        input[0].Type = INPUT_MOUSE;
        input[0].iu.mi.dx = 0;
        input[0].iu.mi.dy = 0;
        input[0].iu.mi.mouseData = 0;
        input[0].iu.mi.dwFlags = MOUSEEVENTF_MIDDLEUP;
        input[0].iu.mi.time = 0;
        input[0].iu.mi.dwExtraInfo = UIntPtr.Zero;
        return SendInput(1, input, Marshal.SizeOf(typeof(INPUT)));
    }

    public static uint MouseWheel(double value)
    {
        INPUT[] input = new INPUT[1];
        input[0].Type = INPUT_MOUSE;
        input[0].iu.mi.dx = 0;
        input[0].iu.mi.dy = 0;
        input[0].iu.mi.mouseData = (uint)Convert.ToInt32(Math.Round((double)WHEEL_DELTA * value));
        input[0].iu.mi.dwFlags = MOUSEEVENTF_WHEEL;
        input[0].iu.mi.time = 0;
        input[0].iu.mi.dwExtraInfo = UIntPtr.Zero;
        return SendInput(1, input, Marshal.SizeOf(typeof(INPUT)));
    }

    public static uint MouseHWheel(double value)
    {
        INPUT[] input = new INPUT[1];
        input[0].Type = INPUT_MOUSE;
        input[0].iu.mi.dx = 0;
        input[0].iu.mi.dy = 0;
        input[0].iu.mi.mouseData = (uint)Convert.ToInt32(Math.Round((double)WHEEL_DELTA * value));
        input[0].iu.mi.dwFlags = MOUSEEVENTF_HWHEEL;
        input[0].iu.mi.time = 0;
        input[0].iu.mi.dwExtraInfo = UIntPtr.Zero;
        return SendInput(1, input, Marshal.SizeOf(typeof(INPUT)));
    }

    public static uint KeyDown(short wVk)
    {
        INPUT[] input = new INPUT[1];
        input[0].Type = INPUT_KEYBOARD;
        input[0].iu.ki.wVk = wVk;
        input[0].iu.ki.wScan = 0;
        input[0].iu.ki.dwFlags = 0;
        input[0].iu.ki.time = 0;
        input[0].iu.ki.dwExtraInfo = UIntPtr.Zero;
        return SendInput(1, input, Marshal.SizeOf(typeof(INPUT)));
    }

    public static uint KeyUp(short wVk)
    {
        INPUT[] input = new INPUT[1];
        input[0].Type = INPUT_KEYBOARD;
        input[0].iu.ki.wVk = wVk;
        input[0].iu.ki.wScan = 0;
        input[0].iu.ki.dwFlags = KEYEVENTF_KEYUP;
        input[0].iu.ki.time = 0;
        input[0].iu.ki.dwExtraInfo = UIntPtr.Zero;
        return SendInput(1, input, Marshal.SizeOf(typeof(INPUT)));
    }
}
'@ -ReferencedAssemblies 'System.Windows.Forms'

<#
 マウスボタンクリック
#>
function global:Click-Mouse
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter(Mandatory=$true, 
                   Position=0)]
        [ValidateSet("Left", "Middle", "Right")]
        $Button,

        [Parameter(Mandatory=$false, 
                   Position=1)]
        [Switch]
        $Double,

        [Parameter(Mandatory=$false)]
        [int]
        $X = [System.Windows.Forms.Cursor]::Position.X,

        [Parameter(Mandatory=$false)]
        [int]
        $Y = [System.Windows.Forms.Cursor]::Position.Y
    )

    if ($pscmdlet.ShouldProcess($Button, "Click"))
    {
        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($X, $Y)
        $count = if ($Double.IsPresent) { 2 } else { 1 }
        Switch ($Button) {
            'Left' {
                (1..$count) | % {
                    [InputSender]::MouseLeftDown() | Out-Null
                    [InputSender]::MouseLeftUp() | Out-Null
                }
            }
            'Middle' {
                (1..$count) | % {
                    [InputSender]::MouseMiddleDown() | Out-Null
                    [InputSender]::MouseMiddleUp() | Out-Null
                }
            }
            'Right' {
                (1..$count) | % {
                    [InputSender]::MouseRightDown() | Out-Null
                    [InputSender]::MouseRightUp() | Out-Null
                }
            }
        }
    }
}

