Imports System.Runtime.CompilerServices
Imports Microsoft.Office.Interop.Excel

''' <summary>
''' Contains extensions of <see cref="Application"/>.
''' </summary>
Public Module ExcelApplicationExtension
    ''' <summary>
    ''' Sets the specified <see cref="Application"/>'s show state.
    ''' </summary>
    ''' <param name="app">The Excel application.</param>
    ''' <param name="windowStyle">The window style.</param>
    ''' <returns><see langword="True"/> if the window was previously visible. 
    ''' Otherwise, <see langword="False"/>.
    ''' </returns>
    <Extension>
    Public Function ShowWindow(app As Application,
            Optional windowStyle As AppWinStyle = vbNormalFocus)
        Dim hwnd = app.Hwnd
        Return NativeMethods.ShowWindow(hwnd, windowStyle)
    End Function

    ''' <summary>
    ''' Checks whether the application window is visible.
    ''' </summary>
    ''' <param name="app">The Excel application to check.</param>
    ''' <returns><see langword="True"/> if the window is visible. 
    ''' Otherwise, <see langword="False"/>.
    ''' </returns>
    <Extension>
    Public Function IsWindowVisible(app As Application) As Boolean
        Dim hwnd = app.Hwnd
        Return NativeMethods.IsWindowVisible(hwnd)
    End Function
End Module
