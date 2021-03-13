Imports Nukepayload2.Interop.Office365.Excel

<TestClass>
Public Class NewAppSample
    Inherits ExcelNewAppUnitTestBase

    <TestMethod>
    Public Sub HelloWorld()
        OnExitCloseApp = False
        Application.ShowWindow()

        Dim wb = Workbooks.Add
        wb.ActiveSheet.Range("A1").Value = "Hello world!"
    End Sub

End Class
