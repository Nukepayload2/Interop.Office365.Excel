Imports Nukepayload2.Interop.Office365.Excel

<TestClass>
Public Class RunningAppSample
    Inherits ExcelRunningAppUnitTestBase

    <TestMethod>
    Public Sub HelloWorld()
        ' TODO: Open Excel and create workbook manually before running this macro

        Range("A1").Value = "Hello world!"
    End Sub

End Class
