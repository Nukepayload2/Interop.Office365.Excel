Imports System.Runtime.InteropServices

Class NativeMethods
    Declare Function ShowWindow Lib "user32" (
        hWnd As IntPtr, cmdShow As Integer
    ) As <MarshalAs(UnmanagedType.Bool)> Boolean

    Declare Function IsWindowVisible Lib "user32" (
        hWnd As IntPtr
    ) As <MarshalAs(UnmanagedType.Bool)> Boolean
End Class
