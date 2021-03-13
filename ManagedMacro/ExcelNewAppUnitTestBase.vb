Imports Microsoft.Office.Interop.Excel
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public MustInherit Class ExcelNewAppUnitTestBase
    Inherits ExcelMacroTestBase

    Private _app As Application

    <TestInitialize>
    Sub Init()
        _app = New Application
    End Sub

    <TestCleanup>
    Sub Cleanup()
        If OnExitCloseApp Then
            _app.Quit()
        End If
    End Sub

    Protected Overrides ReadOnly Property Application As Application
        Get
            Return _app
        End Get
    End Property

    Protected Property OnExitCloseApp As Boolean = True

End Class
