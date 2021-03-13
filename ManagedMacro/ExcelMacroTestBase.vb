Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel

''' <summary>
''' The base class of .NET macros. This class is not intended to be directly in your code.
''' Your macro class should inherit
''' <see cref="ExcelNewAppUnitTestBase"/> or <see cref="ExcelRunningAppUnitTestBase"/> instead.
''' </summary>
Public MustInherit Class ExcelMacroTestBase
    Protected ReadOnly Property ActiveWorkbook As Workbook
        Get
            Return Application.ActiveWorkbook
        End Get
    End Property

    Protected ReadOnly Property Assistant As Assistant
        Get
            Return Application.Assistant
        End Get
    End Property

    Protected ReadOnly Property Cells As Range
        Get
            Return Application.Cells
        End Get
    End Property
    Protected ReadOnly Property Charts As Sheets
        Get
            Return Application.Charts
        End Get
    End Property
    Protected ReadOnly Property Columns As Range
        Get
            Return Application.Columns
        End Get
    End Property
    Protected ReadOnly Property CommandBars As CommandBars
        Get
            Return Application.CommandBars
        End Get
    End Property
    Protected ReadOnly Property DDEAppReturnCode As Integer
        Get
            Return Application.DDEAppReturnCode
        End Get
    End Property
    Protected ReadOnly Property DialogSheets As Sheets
        Get
            Return Application.DialogSheets
        End Get
    End Property
    Protected ReadOnly Property MenuBars As MenuBars
        Get
            Return Application.MenuBars
        End Get
    End Property
    Protected ReadOnly Property Modules As Modules
        Get
            Return Application.Modules
        End Get
    End Property
    Protected ReadOnly Property Names As Names
        Get
            Return Application.Names
        End Get
    End Property
    Protected ReadOnly Property Range(Cell1 As Object, Optional Cell2 As Object = Nothing) As Range
        Get
            If Cell2 Is Nothing Then
                Cell2 = Type.Missing
            End If
            Return Application.Range(Cell1, Cell2)
        End Get
    End Property
    Protected ReadOnly Property Rows As Range
        Get
            Return Application.Rows
        End Get
    End Property
    Protected ReadOnly Property Selection As Object
        Get
            Return Application.Selection
        End Get
    End Property
    Protected ReadOnly Property Sheets As Sheets
        Get
            Return Application.Sheets
        End Get
    End Property
    ReadOnly Property ShortcutMenus(Index As Integer) As Menu
        Get
            Return Application.ShortcutMenus(Index)
        End Get
    End Property
    Protected ReadOnly Property ThisWorkbook As Workbook
        Get
            Return Application.ThisWorkbook
        End Get
    End Property
    Protected ReadOnly Property Toolbars As Toolbars
        Get
            Return Application.Toolbars
        End Get
    End Property
    Protected ReadOnly Property Windows As Windows
        Get
            Return Application.Windows
        End Get
    End Property
    Protected ReadOnly Property Workbooks As Workbooks
        Get
            Return Application.Workbooks
        End Get
    End Property
    Protected ReadOnly Property WorksheetFunction As WorksheetFunction
        Get
            Return Application.WorksheetFunction
        End Get
    End Property
    Protected ReadOnly Property Worksheets As Sheets
        Get
            Return Application.Worksheets
        End Get
    End Property
    Protected ReadOnly Property AddIns As AddIns
        Get
            Return Application.AddIns
        End Get
    End Property
    Protected ReadOnly Property Excel4IntlMacroSheets As Sheets
        Get
            Return Application.Excel4IntlMacroSheets
        End Get
    End Property
    Protected ReadOnly Property Excel4MacroSheets As Sheets
        Get
            Return Application.Excel4MacroSheets
        End Get
    End Property
    Protected ReadOnly Property ActiveSheet As Worksheet
        Get
            Return Application.ActiveSheet
        End Get
    End Property
    Property ActivePrinter As String
        Get
            Return Application.ActivePrinter
        End Get
        Set(value As String)
            Application.ActivePrinter = value
        End Set
    End Property
    Protected ReadOnly Property ActiveMenuBar As MenuBar
        Get
            Return Application.ActiveMenuBar
        End Get
    End Property
    Protected ReadOnly Property ActiveDialog As DialogSheet
        Get
            Return Application.ActiveDialog
        End Get
    End Property
    Protected ReadOnly Property ActiveChart As Chart
        Get
            Return Application.ActiveChart
        End Get
    End Property
    Protected ReadOnly Property ActiveCell As Range
        Get
            Return Application.ActiveCell
        End Get
    End Property
    Protected ReadOnly Property Parent As Application
        Get
            Return Application.Parent
        End Get
    End Property
    Protected ReadOnly Property Creator As XlCreator
        Get
            Return Application.Creator
        End Get
    End Property
    Protected MustOverride ReadOnly Property Application As Application
    Protected ReadOnly Property ActiveWindow As Window
        Get
            Return Application.ActiveWindow
        End Get
    End Property

    Protected Sub DDEExecute(Channel As Integer, [String] As String)
        Application.DDEExecute(Channel, [String])
    End Sub
    Protected Sub DDEPoke(Channel As Integer, Item As Object, Data As Object)
        Application.DDEPoke(Channel, Item, Data)
    End Sub
    Protected Sub DDETerminate(Channel As Integer)
        Application.DDETerminate(Channel)
    End Sub
    Protected Sub Calculate()
        Application.Calculate()
    End Sub

    Protected Sub SendKeys(Keys As Object, Optional Wait As Object = Nothing)
        Application.SendKeys(Keys, Wait)
    End Sub

    Protected Function ExecuteExcel4Macro([String] As String) As Object
        Return Application.ExecuteExcel4Macro([String])
    End Function
    Protected Function _Evaluate(Name As Object) As Object
        Return Application._Evaluate(Name)
    End Function
    Protected Function Evaluate(Name As Object) As Object
        Return Application.Evaluate(Name)
    End Function
    Protected Function Run(Optional Macro As Object = Nothing, Optional Arg1 As Object = Nothing, Optional Arg2 As Object = Nothing, Optional Arg3 As Object = Nothing, Optional Arg4 As Object = Nothing, Optional Arg5 As Object = Nothing, Optional Arg6 As Object = Nothing, Optional Arg7 As Object = Nothing, Optional Arg8 As Object = Nothing, Optional Arg9 As Object = Nothing, Optional Arg10 As Object = Nothing, Optional Arg11 As Object = Nothing, Optional Arg12 As Object = Nothing, Optional Arg13 As Object = Nothing, Optional Arg14 As Object = Nothing, Optional Arg15 As Object = Nothing, Optional Arg16 As Object = Nothing, Optional Arg17 As Object = Nothing, Optional Arg18 As Object = Nothing, Optional Arg19 As Object = Nothing, Optional Arg20 As Object = Nothing, Optional Arg21 As Object = Nothing, Optional Arg22 As Object = Nothing, Optional Arg23 As Object = Nothing, Optional Arg24 As Object = Nothing, Optional Arg25 As Object = Nothing, Optional Arg26 As Object = Nothing, Optional Arg27 As Object = Nothing, Optional Arg28 As Object = Nothing, Optional Arg29 As Object = Nothing, Optional Arg30 As Object = Nothing) As Object
        If Macro Is Nothing Then Macro = Type.Missing
        If Arg1 Is Nothing Then Arg1 = Type.Missing
        If Arg2 Is Nothing Then Arg2 = Type.Missing
        If Arg3 Is Nothing Then Arg3 = Type.Missing
        If Arg4 Is Nothing Then Arg4 = Type.Missing
        If Arg5 Is Nothing Then Arg5 = Type.Missing
        If Arg6 Is Nothing Then Arg6 = Type.Missing
        If Arg7 Is Nothing Then Arg7 = Type.Missing
        If Arg8 Is Nothing Then Arg8 = Type.Missing
        If Arg9 Is Nothing Then Arg9 = Type.Missing
        If Arg10 Is Nothing Then Arg10 = Type.Missing
        If Arg11 Is Nothing Then Arg11 = Type.Missing
        If Arg12 Is Nothing Then Arg12 = Type.Missing
        If Arg13 Is Nothing Then Arg13 = Type.Missing
        If Arg14 Is Nothing Then Arg14 = Type.Missing
        If Arg15 Is Nothing Then Arg15 = Type.Missing
        If Arg16 Is Nothing Then Arg16 = Type.Missing
        If Arg17 Is Nothing Then Arg17 = Type.Missing
        If Arg18 Is Nothing Then Arg18 = Type.Missing
        If Arg19 Is Nothing Then Arg19 = Type.Missing
        If Arg20 Is Nothing Then Arg20 = Type.Missing
        If Arg21 Is Nothing Then Arg21 = Type.Missing
        If Arg22 Is Nothing Then Arg22 = Type.Missing
        If Arg23 Is Nothing Then Arg23 = Type.Missing
        If Arg24 Is Nothing Then Arg24 = Type.Missing
        If Arg25 Is Nothing Then Arg25 = Type.Missing
        If Arg26 Is Nothing Then Arg26 = Type.Missing
        If Arg27 Is Nothing Then Arg27 = Type.Missing
        If Arg28 Is Nothing Then Arg28 = Type.Missing
        If Arg29 Is Nothing Then Arg29 = Type.Missing
        If Arg30 Is Nothing Then Arg30 = Type.Missing
        Return Application.Run(Macro, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30)
    End Function

    Protected Function DDERequest(Channel As Integer, Item As String) As Object
        Return Application.DDERequest(Channel, Item)
    End Function
    Protected Function DDEInitiate(App As String, Topic As String) As Integer
        Return Application.DDEInitiate(App, Topic)
    End Function
    Protected Function Union(Arg1 As Range, Arg2 As Range, Optional Arg3 As Object = Nothing, Optional Arg4 As Object = Nothing, Optional Arg5 As Object = Nothing, Optional Arg6 As Object = Nothing, Optional Arg7 As Object = Nothing, Optional Arg8 As Object = Nothing, Optional Arg9 As Object = Nothing, Optional Arg10 As Object = Nothing, Optional Arg11 As Object = Nothing, Optional Arg12 As Object = Nothing, Optional Arg13 As Object = Nothing, Optional Arg14 As Object = Nothing, Optional Arg15 As Object = Nothing, Optional Arg16 As Object = Nothing, Optional Arg17 As Object = Nothing, Optional Arg18 As Object = Nothing, Optional Arg19 As Object = Nothing, Optional Arg20 As Object = Nothing, Optional Arg21 As Object = Nothing, Optional Arg22 As Object = Nothing, Optional Arg23 As Object = Nothing, Optional Arg24 As Object = Nothing, Optional Arg25 As Object = Nothing, Optional Arg26 As Object = Nothing, Optional Arg27 As Object = Nothing, Optional Arg28 As Object = Nothing, Optional Arg29 As Object = Nothing, Optional Arg30 As Object = Nothing) As Range
        If Arg1 Is Nothing Then Arg1 = Type.Missing
        If Arg2 Is Nothing Then Arg2 = Type.Missing
        If Arg3 Is Nothing Then Arg3 = Type.Missing
        If Arg4 Is Nothing Then Arg4 = Type.Missing
        If Arg5 Is Nothing Then Arg5 = Type.Missing
        If Arg6 Is Nothing Then Arg6 = Type.Missing
        If Arg7 Is Nothing Then Arg7 = Type.Missing
        If Arg8 Is Nothing Then Arg8 = Type.Missing
        If Arg9 Is Nothing Then Arg9 = Type.Missing
        If Arg10 Is Nothing Then Arg10 = Type.Missing
        If Arg11 Is Nothing Then Arg11 = Type.Missing
        If Arg12 Is Nothing Then Arg12 = Type.Missing
        If Arg13 Is Nothing Then Arg13 = Type.Missing
        If Arg14 Is Nothing Then Arg14 = Type.Missing
        If Arg15 Is Nothing Then Arg15 = Type.Missing
        If Arg16 Is Nothing Then Arg16 = Type.Missing
        If Arg17 Is Nothing Then Arg17 = Type.Missing
        If Arg18 Is Nothing Then Arg18 = Type.Missing
        If Arg19 Is Nothing Then Arg19 = Type.Missing
        If Arg20 Is Nothing Then Arg20 = Type.Missing
        If Arg21 Is Nothing Then Arg21 = Type.Missing
        If Arg22 Is Nothing Then Arg22 = Type.Missing
        If Arg23 Is Nothing Then Arg23 = Type.Missing
        If Arg24 Is Nothing Then Arg24 = Type.Missing
        If Arg25 Is Nothing Then Arg25 = Type.Missing
        If Arg26 Is Nothing Then Arg26 = Type.Missing
        If Arg27 Is Nothing Then Arg27 = Type.Missing
        If Arg28 Is Nothing Then Arg28 = Type.Missing
        If Arg29 Is Nothing Then Arg29 = Type.Missing
        If Arg30 Is Nothing Then Arg30 = Type.Missing
        Return Application.Union(Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30)
    End Function
    Protected Function Intersect(Arg1 As Range, Arg2 As Range, Optional Arg3 As Object = Nothing, Optional Arg4 As Object = Nothing, Optional Arg5 As Object = Nothing, Optional Arg6 As Object = Nothing, Optional Arg7 As Object = Nothing, Optional Arg8 As Object = Nothing, Optional Arg9 As Object = Nothing, Optional Arg10 As Object = Nothing, Optional Arg11 As Object = Nothing, Optional Arg12 As Object = Nothing, Optional Arg13 As Object = Nothing, Optional Arg14 As Object = Nothing, Optional Arg15 As Object = Nothing, Optional Arg16 As Object = Nothing, Optional Arg17 As Object = Nothing, Optional Arg18 As Object = Nothing, Optional Arg19 As Object = Nothing, Optional Arg20 As Object = Nothing, Optional Arg21 As Object = Nothing, Optional Arg22 As Object = Nothing, Optional Arg23 As Object = Nothing, Optional Arg24 As Object = Nothing, Optional Arg25 As Object = Nothing, Optional Arg26 As Object = Nothing, Optional Arg27 As Object = Nothing, Optional Arg28 As Object = Nothing, Optional Arg29 As Object = Nothing, Optional Arg30 As Object = Nothing) As Range
        If Arg1 Is Nothing Then Arg1 = Type.Missing
        If Arg2 Is Nothing Then Arg2 = Type.Missing
        If Arg3 Is Nothing Then Arg3 = Type.Missing
        If Arg4 Is Nothing Then Arg4 = Type.Missing
        If Arg5 Is Nothing Then Arg5 = Type.Missing
        If Arg6 Is Nothing Then Arg6 = Type.Missing
        If Arg7 Is Nothing Then Arg7 = Type.Missing
        If Arg8 Is Nothing Then Arg8 = Type.Missing
        If Arg9 Is Nothing Then Arg9 = Type.Missing
        If Arg10 Is Nothing Then Arg10 = Type.Missing
        If Arg11 Is Nothing Then Arg11 = Type.Missing
        If Arg12 Is Nothing Then Arg12 = Type.Missing
        If Arg13 Is Nothing Then Arg13 = Type.Missing
        If Arg14 Is Nothing Then Arg14 = Type.Missing
        If Arg15 Is Nothing Then Arg15 = Type.Missing
        If Arg16 Is Nothing Then Arg16 = Type.Missing
        If Arg17 Is Nothing Then Arg17 = Type.Missing
        If Arg18 Is Nothing Then Arg18 = Type.Missing
        If Arg19 Is Nothing Then Arg19 = Type.Missing
        If Arg20 Is Nothing Then Arg20 = Type.Missing
        If Arg21 Is Nothing Then Arg21 = Type.Missing
        If Arg22 Is Nothing Then Arg22 = Type.Missing
        If Arg23 Is Nothing Then Arg23 = Type.Missing
        If Arg24 Is Nothing Then Arg24 = Type.Missing
        If Arg25 Is Nothing Then Arg25 = Type.Missing
        If Arg26 Is Nothing Then Arg26 = Type.Missing
        If Arg27 Is Nothing Then Arg27 = Type.Missing
        If Arg28 Is Nothing Then Arg28 = Type.Missing
        If Arg29 Is Nothing Then Arg29 = Type.Missing
        If Arg30 Is Nothing Then Arg30 = Type.Missing
        Return Application.Intersect(Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30)
    End Function
    Protected Function _Run2(Optional Macro As Object = Nothing, Optional Arg1 As Object = Nothing, Optional Arg2 As Object = Nothing, Optional Arg3 As Object = Nothing, Optional Arg4 As Object = Nothing, Optional Arg5 As Object = Nothing, Optional Arg6 As Object = Nothing, Optional Arg7 As Object = Nothing, Optional Arg8 As Object = Nothing, Optional Arg9 As Object = Nothing, Optional Arg10 As Object = Nothing, Optional Arg11 As Object = Nothing, Optional Arg12 As Object = Nothing, Optional Arg13 As Object = Nothing, Optional Arg14 As Object = Nothing, Optional Arg15 As Object = Nothing, Optional Arg16 As Object = Nothing, Optional Arg17 As Object = Nothing, Optional Arg18 As Object = Nothing, Optional Arg19 As Object = Nothing, Optional Arg20 As Object = Nothing, Optional Arg21 As Object = Nothing, Optional Arg22 As Object = Nothing, Optional Arg23 As Object = Nothing, Optional Arg24 As Object = Nothing, Optional Arg25 As Object = Nothing, Optional Arg26 As Object = Nothing, Optional Arg27 As Object = Nothing, Optional Arg28 As Object = Nothing, Optional Arg29 As Object = Nothing, Optional Arg30 As Object = Nothing) As Object
        If Macro Is Nothing Then Macro = Type.Missing
        If Arg1 Is Nothing Then Arg1 = Type.Missing
        If Arg2 Is Nothing Then Arg2 = Type.Missing
        If Arg3 Is Nothing Then Arg3 = Type.Missing
        If Arg4 Is Nothing Then Arg4 = Type.Missing
        If Arg5 Is Nothing Then Arg5 = Type.Missing
        If Arg6 Is Nothing Then Arg6 = Type.Missing
        If Arg7 Is Nothing Then Arg7 = Type.Missing
        If Arg8 Is Nothing Then Arg8 = Type.Missing
        If Arg9 Is Nothing Then Arg9 = Type.Missing
        If Arg10 Is Nothing Then Arg10 = Type.Missing
        If Arg11 Is Nothing Then Arg11 = Type.Missing
        If Arg12 Is Nothing Then Arg12 = Type.Missing
        If Arg13 Is Nothing Then Arg13 = Type.Missing
        If Arg14 Is Nothing Then Arg14 = Type.Missing
        If Arg15 Is Nothing Then Arg15 = Type.Missing
        If Arg16 Is Nothing Then Arg16 = Type.Missing
        If Arg17 Is Nothing Then Arg17 = Type.Missing
        If Arg18 Is Nothing Then Arg18 = Type.Missing
        If Arg19 Is Nothing Then Arg19 = Type.Missing
        If Arg20 Is Nothing Then Arg20 = Type.Missing
        If Arg21 Is Nothing Then Arg21 = Type.Missing
        If Arg22 Is Nothing Then Arg22 = Type.Missing
        If Arg23 Is Nothing Then Arg23 = Type.Missing
        If Arg24 Is Nothing Then Arg24 = Type.Missing
        If Arg25 Is Nothing Then Arg25 = Type.Missing
        If Arg26 Is Nothing Then Arg26 = Type.Missing
        If Arg27 Is Nothing Then Arg27 = Type.Missing
        If Arg28 Is Nothing Then Arg28 = Type.Missing
        If Arg29 Is Nothing Then Arg29 = Type.Missing
        If Arg30 Is Nothing Then Arg30 = Type.Missing
        Return Application._Run2(Macro, Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30)
    End Function

End Class
