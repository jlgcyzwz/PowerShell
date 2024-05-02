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
