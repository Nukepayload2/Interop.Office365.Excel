Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

''' <summary>
''' Contains extensions and helpers of <see cref="XlCVError"/>.
''' </summary>
Public Module XlCVErrExtension

    ''' <summary>
    ''' Converts <see cref="XlCVError"/> to <see cref="String"/>. 
    ''' For example, <see cref="XlCVError.xlErrValue"/> is #VALUE! .
    ''' </summary>
    ''' <param name="cVErrValue">The error <see cref="XlCVError"/> to convert.</param>
    ''' <returns>The display string.</returns>
    <Extension>
    Public Function ToDisplayString(cVErrValue As XlCVError) As String
        Select Case cVErrValue
            Case XlCVError.xlErrBlocked
                Return "#BLOCKED!"
            Case XlCVError.xlErrCalc
                Return "#CALC!"
            Case XlCVError.xlErrConnect
                Return "#CONNECT!"
            Case XlCVError.xlErrDiv0
                Return "#DIV/0!"
            Case XlCVError.xlErrField
                Return "#FIELD!"
            Case XlCVError.xlErrGettingData
                Return "#GETTING_DATA"
            Case XlCVError.xlErrNA
                Return "#N/A"
            Case XlCVError.xlErrName
                Return "#NAME?"
            Case XlCVError.xlErrSpill
                Return "#SPILL!"
            Case XlCVError.xlErrNull
                Return "#NULL!"
            Case XlCVError.xlErrNum
                Return "#NUM!"
            Case XlCVError.xlErrRef
                Return "#REF!"
            Case XlCVError.xlErrUnknown
                Return "#UNKNOWN!"
            Case XlCVError.xlErrValue
                Return "#VALUE!"
        End Select
        Throw New ArgumentException
    End Function

    Private Const VErrMask = &H800A0000

    ''' <summary>
    ''' Converts <see cref="XlCVError"/> to an <see cref="ErrorWrapper"/> which can be 
    ''' set to <see cref="Range.Value2"/> or <see cref="Range.Value(Object)"/>.
    ''' </summary>
    ''' <param name="errCode">The error to convert.</param>
    ''' <returns>An <see cref="ErrorWrapper"/> which can be 
    ''' set to <see cref="Range.Value2"/> or <see cref="Range.Value(Object)"/>.</returns>
    Public Function CVErr(errCode As XlCVError) As ErrorWrapper
        Return New ErrorWrapper(VErrMask Or errCode)
    End Function

    ''' <summary>
    ''' Converts the <see cref="Integer"/> which is value of 
    ''' <see cref="Range.Value2"/> or <see cref="Range.Value(Object)"/> to <see cref="XlCVError"/>.
    ''' </summary>
    ''' <param name="errorCode">
    ''' The error code from <see cref="Range.Value2"/> or <see cref="Range.Value(Object)"/>.
    ''' </param>
    ''' <returns>The converted <see cref="XlCVError"/> value.</returns>
    Public Function ErrorCodeToXlCVErr(errorCode As Integer) As XlCVError
        Return (Not VErrMask) And errorCode
    End Function

End Module
