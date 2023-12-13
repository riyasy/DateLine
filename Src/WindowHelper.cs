using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.Windows.Interop;

namespace DateLine;

internal class WindowHelper
{
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr FindWindowEx(IntPtr hP, IntPtr hC, string sC, string sW);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumWindows(EnumedWindow lpEnumFunc, ArrayList
        lParam);

    public delegate bool EnumedWindow(IntPtr handleWindow, ArrayList handles);

    public static bool GetWindowHandle(IntPtr windowHandle, ArrayList
        windowHandles)
    {
        windowHandles.Add(windowHandle);
        return true;
    }

    public static void SetAsDesktopChild(System.Windows.Window childWindow)
    {
        var windowHandles = new ArrayList();
        EnumedWindow callBackPtr = GetWindowHandle;
        EnumWindows(callBackPtr, windowHandles);

        foreach (IntPtr windowHandle in windowHandles)
        {
            var hNextWin = FindWindowEx(windowHandle, IntPtr.Zero,
                "SHELLDLL_DefView", null);
            if (hNextWin == IntPtr.Zero) continue;
            var interop = new WindowInteropHelper(childWindow);
            interop.EnsureHandle();
            interop.Owner = hNextWin;
        }
    }
}