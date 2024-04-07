Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class InputSimulator
{
    [DllImport("user32.dll", EntryPoint = "SendInput", SetLastError = true)]
    private static extern uint SendInput(uint cInputs, INPUT[] pInputs, int cbSize);

    [StructLayout(LayoutKind.Explicit)]
    private struct INPUT
    {
        [FieldOffset(0)]
        public uint Type;
        [FieldOffset(4)]
        public MOUSEINPUT mi;
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

    // INPUT
    // Type
    private const uint INPUT_MOUSE = 0;
    private const uint INPUT_KEYBOARD = 1;

    // MOUSEINPUT
    // mouseData
    public const uint XBUTTON1 = 0x0001;	// 最初の X ボタンを押すか離した場合に設定します。
    public const uint XBUTTON2 = 0x0002;	// 2 つ目の X ボタンを押すか離すかを設定します。
    // dwFlags
    public const uint MOUSEEVENTF_MOVE = 0x0001;			// 移動が発生しました。
    public const uint MOUSEEVENTF_LEFTDOWN = 0x0002;		// 左ボタンが押されました。
    public const uint MOUSEEVENTF_LEFTUP = 0x0004;			// 左側のボタンが離されました。
    public const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;		// 右ボタンが押されました。
    public const uint MOUSEEVENTF_RIGHTUP = 0x0010;			// 右側のボタンが離されました。
    public const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020;		// 中央のボタンが押されました。
    public const uint MOUSEEVENTF_MIDDLEUP = 0x0040;		// 中央のボタンが離されました。
    public const uint MOUSEEVENTF_XDOWN = 0x0080;			// X ボタンが押されました。
    public const uint MOUSEEVENTF_XUP = 0x0100;				// X ボタンが解放されました。
    public const uint MOUSEEVENTF_WHEEL = 0x0800;			// マウスにホイールがある場合は、ホイールを移動しました。 移動の量は mouseData で指定されます。
    public const uint MOUSEEVENTF_HWHEEL = 0x1000;			// "マウスにホイールがある場合、ホイールは水平方向に移動しました。 移動の量は mouseData で指定されます。 
    public const uint MOUSEEVENTF_MOVE_NOCOALESCE = 0x2000;	// "WM_MOUSEMOVEメッセージは結合されません。 既定の動作では、 メッセージWM_MOUSEMOVE 合体します。 
    public const uint MOUSEEVENTF_VIRTUALDESK = 0x4000;		// 座標をデスクトップ全体にマップします。 MOUSEEVENTF_ABSOLUTEで使用する必要があります。
    public const uint MOUSEEVENTF_ABSOLUTE = 0x8000;		// dx メンバーと dy メンバーには、正規化された絶対座標が含まれています。 フラグが設定されていない場合、 dxと dy には相対データ (最後に報告された位置以降の位置の変化) が含まれます。 このフラグは、システムに接続されているマウスやその他のポインティング デバイスの種類に関係なく、設定することも設定することもできません。 相対的なマウスの動きの詳細については、次の「解説」セクションを参照してください。

    public static void SendMouse(int dx, int dy, uint mouseData, uint dwFlags)
    {
        INPUT[] input = new INPUT[1];
        input[0].Type = INPUT_MOUSE;
        input[0].mi.dx = dx;
        input[0].mi.dy = dy;
        input[0].mi.mouseData = mouseData;
        input[0].mi.dwFlags = dwFlags;
        SendInput(1, input, Marshal.SizeOf(typeof(INPUT)));
    }
}
'@

