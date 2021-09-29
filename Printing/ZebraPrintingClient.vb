Imports System.IO
Imports System.Runtime.InteropServices

Namespace Printing

    Public Class ZebraPrintingClient
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
        Structure DocInfo
            <MarshalAs(UnmanagedType.LPWStr)> Public pDocName As String
            <MarshalAs(UnmanagedType.LPWStr)> Public pOutputFile As String
            <MarshalAs(UnmanagedType.LPWStr)> Public pDataType As String
        End Structure

        <DllImport("winspool.Drv", EntryPoint:="OpenPrinterW", _
                   SetLastError:=True, CharSet:=CharSet.Unicode, _
                   ExactSpelling:=True, CallingConvention:=CallingConvention.Cdecl)> _
        Public Shared Function OpenPrinter(ByVal src As String, ByRef hPrinter As IntPtr, ByVal pd As Int32) As Boolean
        End Function
        <DllImport("winspool.Drv", EntryPoint:="ClosePrinter", _
                   SetLastError:=True, CharSet:=CharSet.Unicode, _
                   ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
        Public Shared Function ClosePrinter(ByVal hPrinter As IntPtr) As Boolean
        End Function
        <DllImport("winspool.Drv", EntryPoint:="StartDocPrinterW", _
                   SetLastError:=True, CharSet:=CharSet.Unicode, _
                   ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
        Public Shared Function StartDocPrinter(ByVal hPrinter As IntPtr, ByVal level As Int32, ByRef pDi As DocInfo) As Boolean
        End Function
        <DllImport("winspool.Drv", EntryPoint:="EndDocPrinter", _
                   SetLastError:=True, CharSet:=CharSet.Unicode, _
                   ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
        Public Shared Function EndDocPrinter(ByVal hPrinter As IntPtr) As Boolean
        End Function
        <DllImport("winspool.Drv", EntryPoint:="StartPagePrinter", _
                   SetLastError:=True, CharSet:=CharSet.Unicode, _
                   ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
        Public Shared Function StartPagePrinter(ByVal hPrinter As IntPtr) As Boolean
        End Function
        <DllImport("winspool.Drv", EntryPoint:="EndPagePrinter", _
                   SetLastError:=True, CharSet:=CharSet.Unicode, _
                   ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
        Public Shared Function EndPagePrinter(ByVal hPrinter As IntPtr) As Boolean
        End Function
        <DllImport("winspool.Drv", EntryPoint:="WritePrinter", _
                   SetLastError:=True, CharSet:=CharSet.Unicode, _
                   ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
        Public Shared Function WritePrinter(ByVal hPrinter As IntPtr, ByVal pBytes As IntPtr, ByVal dwCount As Int32, ByRef dwWritten As Int32) As Boolean
        End Function

        Private Shared Function SendBytesToPrinter(ByVal szPrinterName As String, ByVal pBytes As IntPtr, ByVal dwCount As Int32) As PrintResponse
            Dim hPrinter As IntPtr
            Dim dwError As Int32 = 0 ' Operation completed successfully
            Dim di As DocInfo
            Dim dwWritten As Int32
            Dim bSuccess As Boolean

            di = New DocInfo
            With di
                .pDocName = "RAW DOC"
                .pDataType = "RAW"
            End With

            bSuccess = False
            If OpenPrinter(szPrinterName, hPrinter, 0) Then
                If StartDocPrinter(hPrinter, 1, di) Then
                    If StartPagePrinter(hPrinter) Then

                        bSuccess = WritePrinter(hPrinter, pBytes, dwCount, dwWritten)
                        EndPagePrinter(hPrinter)
                    End If
                    EndDocPrinter(hPrinter)
                End If
                ClosePrinter(hPrinter)
            End If

            If bSuccess = False Then
                dwError = Marshal.GetLastWin32Error()
            End If
            Return New PrintResponse(bSuccess, dwError)
        End Function

        Const X_STARTING_POSITION = 40
        Const Y_FONT_INCREMENT = 60
        Const FONT_PREFIX = "^AB,17,10^FD"
        Const Y_BARCODE_INCREMENT = 80
        Const BARCODE_PREFIX = "^BCN,40,Y,N,N,D^FD"
        Const COMMAND_HEADER = "^XA"
        Const COMMAND_FOOTER = "^XZ"
        Const FIELD_ORIGIN = "^FO"
        Const FIELD_SEPERATOR = "^FS"

        Private Shared _yPosition As Integer = 0

        Private Shared Function GetBarcodeCommand(barcodeText As String) As String
            _yPosition += Y_BARCODE_INCREMENT
            Return $"{GetOriginPosition()}{BARCODE_PREFIX}{barcodeText}{FIELD_SEPERATOR}"
        End Function

        Private Shared Function GetFontCommand(lineText as String) as String
            _yPosition += Y_FONT_INCREMENT
            Return $"{GetOriginPosition()}{FONT_PREFIX}{lineText}{FIELD_SEPERATOR}"
        End Function

        Private Shared Function GetOriginPosition() As String
            Return $"{FIELD_ORIGIN}{X_STARTING_POSITION},{_yPosition}"
        End Function

        Private Shared Function GetStringCounts(fullCommand As String) As Tuple(Of IntPtr, Int32)
            Return New Tuple(Of IntPtr, Int32)(Marshal.StringToCoTaskMemAnsi(fullCommand), fullCommand.Length())
        End Function

        ''' <summary>
        ''' Write your own ZPL command
        ''' </summary>
        ''' <param name="printerName">The name of the printer can be found in Printers and Scanners </param>
        ''' <param name="commandString">The full ZPL command that is to be printed</param>
        Public Shared Function SendStringToPrinter(ByVal printerName As String, ByVal commandString As String) As PrintResponse
           Dim counts = GetStringCounts(commandString)
            Dim response = SendBytesToPrinter(printerName, counts.Item1, counts.Item2)
            Marshal.FreeCoTaskMem(counts.Item1)
            Return response
        End Function

        ''' <summary>
        ''' Print a dictionary of labels and values with one barcode at the end.
        ''' </summary>
        ''' <param name="printerName">The name of the printer can be found in Printers and Scanners</param>
        ''' <param name="labelsAndValues">String pairs that should printed on the label</param>
        ''' <param name="barcodeValue">The value of a barcode to be printed at the end of the label</param>
        ''' <returns></returns>
        Public Shared Function SendStringToPrinter(ByVal printerName As String, ByVal labelsAndValues As Dictionary(Of String, String), Optional ByVal barcodeValue As String = "") As PrintResponse
            _yPosition = 0
            Dim command As String = COMMAND_HEADER
            For each pair In labelsAndValues
                command &= GetFontCommand($"{pair.Key}: {pair.Value}")
            Next
            If Not String.IsNullOrEmpty(barcodeValue)
                command &= GetBarcodeCommand(barcodeValue)
            End If
            command &= COMMAND_FOOTER
            Dim counts = GetStringCounts(command)
            Dim response = SendBytesToPrinter(printerName, counts.Item1, counts.Item2)
            Marshal.FreeCoTaskMem(counts.Item1)
            Return response
        End Function

        ''' <summary>
        ''' Print a dictionary of labels and values where each value can be represented by a barcode.
        ''' </summary>
        ''' <param name="printerName">The name of the printer can be found in Printers and Scanners</param>
        ''' <param name="labelsAndValuesWithOptionalBarcode">Key = label. Value1 = string to be printed. Value2 = boolean, true if value should be printed as barcode</param>
        ''' <returns></returns>
        Public Shared Function SendStringToPrinter(ByVal printerName As String, ByVal labelsAndValuesWithOptionalBarcode As Dictionary(Of String, Tuple(Of String, Boolean))) As PrintResponse
            _yPosition = 0
            Dim command As String = COMMAND_HEADER
            For Each pair In labelsAndValuesWithOptionalBarcode
                If (pair.Value.Item2) Then
                    command &= $"{pair.Key}"
                    command &= GetBarcodeCommand(pair.Value.Item1)
                Else
                    command &= GetFontCommand($"{pair.Key}: {pair.Value.Item1}")
                End If
            Next

            command &= COMMAND_FOOTER
            Dim counts = GetStringCounts(command)
            Dim response = SendBytesToPrinter(printerName, counts.Item1, counts.Item2)
            Marshal.FreeCoTaskMem(counts.Item1)
            Return response
        End Function

    End Class
End NameSpace