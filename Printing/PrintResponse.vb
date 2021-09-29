
Namespace Printing
    Public Class PrintResponse
        Public Sub New(printPassed As Boolean, errorCode As Int32)
            Passed = printPassed
            Win32ErrorCode = errorCode
        End Sub
        Public Property Passed As Boolean
        Public Property Win32ErrorCode As Int32
    End Class
End NameSpace