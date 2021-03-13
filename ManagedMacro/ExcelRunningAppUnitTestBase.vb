Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' Attaches to an active and visible Excel application then run macros.
''' </summary>
<TestClass>
Public MustInherit Class ExcelRunningAppUnitTestBase
    Inherits ExcelMacroTestBase

    Private _app As Application

    <TestInitialize>
    Sub Init()
        _app = Marshal.GetActiveObject("Excel.Application")

        CloseBackgroundExcelApps()

        If _app Is Nothing Then
            Throw New InvalidOperationException("Excel is not running.")
        End If
    End Sub

    Private Sub CloseBackgroundExcelApps()
        Do While _app IsNot Nothing AndAlso Not _app.IsWindowVisible
            _app.Quit()
            _app = Marshal.GetActiveObject("Excel.Application")
        Loop
    End Sub

    Protected Overrides ReadOnly Property Application As Application
        Get
            Return _app
        End Get
    End Property
End Class
