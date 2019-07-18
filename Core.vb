Imports unoidl.com.sun.star.lang
Imports unoidl.com.sun.star.util
Imports System.IO
Imports System.Environment
Imports unoidl.com.sun.star.uno
Imports uno.util
Imports unoidl.com.sun.star.frame
Imports unoidl.com.sun.star.sheet
Imports unoidl.com.sun.star.beans
Imports unoidl.com.sun.star.container

Module Core

    Private input As String '= "C:\Users\walfr\source\repos\LibreOfficeConvert\LibreOfficeConvert\bin\Debug\LV05340-16398-19-R0.ods"
    Private output As String '= "C:\Users\walfr\source\repos\LibreOfficeConvert\LibreOfficeConvert\bin\Debug\LV05340-16398-19-R0.ods.pdf"

    Private sheetName As String = String.Empty
    Private sheetIndex As Integer = -1

    Private Const TAB As String = Constants.vbTab

    Sub Main(args As String())
        Try
            If args.Length = 2 Then
                input = args(0)
                output = args(1)
                sheetIndex = 1
                Convert(input, output)
            ElseIf args.Length = 3 Then
                input = args(0)
                output = args(1)
                If IsNumeric(args(2)) Then sheetIndex = CInt(args(2)) Else sheetName = args(2)
                Convert(input, output)
            ElseIf args(0).ToLower.Equals("help") Then
                Console.WriteLine("{0}Converte uma planilha Calc para formato PDF.", TAB)
                Console.WriteLine("{0}{1}[...] [...] [...]", TAB, TAB)
                Console.WriteLine("{0}origem{1}Especifica o arquivo a ser convertido.", TAB, TAB)
                Console.WriteLine("{0}destino{1}Especifica o arquivo de saída.", TAB, TAB)
                Console.WriteLine("{0}sheet{1}Especifica qual sheet será convertido (padrão 1). Nome ou index.", TAB, TAB)
            Else
                Console.Write("Invalid parameter number.")
            End If
        Catch ex As Exception
            Console.Write(ex.Message)
        End Try
    End Sub

    Private aLoader As XComponentLoader
    Public ReadOnly Property Loader() As XComponentLoader
        Get
            If aLoader Is Nothing Then
                Init()
            End If
            Return aLoader
        End Get
    End Property

    Private Sub Init()
        aLoader = InitLoader()
        If aLoader Is Nothing Then
            aLoader = InitLoader()
        End If
        If aLoader Is Nothing Then
            aLoader = InitLoader()
        End If
        If aLoader Is Nothing Then
            Console.Write("Can't find OpenOffice.org or LibreOffice. Office must be installed for pdf conversion.")
        End If
    End Sub

    Private Function InitLoader() As XComponentLoader
        Try
            Dim unoPath = String.Empty
            Dim unoPathList = Configuration.ConfigurationManager.AppSettings("UNO_PATH").ToString()
            Dim urePath = String.Empty
            Dim urePathList = Configuration.ConfigurationManager.AppSettings("URE_PATH").ToString()
            Dim path As String = String.Empty

            For Each item As String In unoPathList.Split(";")
                If Directory.Exists(item) Then
                    unoPath = item
                    Exit For
                End If
            Next

            For Each item As String In urePathList.Split(";")
                If Directory.Exists(item) Then
                    urePath = item
                    Exit For
                End If
            Next

            If String.IsNullOrWhiteSpace(unoPath) Then
                Console.Write("Pasta LibreOffice não localizada")
                [Exit](1)
            End If

            If String.IsNullOrEmpty(urePath) Then
                path = String.Format("{0};{1}", Environment.GetEnvironmentVariable("PATH"), unoPath)
                Environment.SetEnvironmentVariable("PATH", path)
                Environment.SetEnvironmentVariable("UNO_PATH", unoPath)
            Else
                path = String.Format("{0};{1}", Environment.GetEnvironmentVariable("PATH"), urePath)
                Environment.SetEnvironmentVariable("PATH", path)
                Environment.SetEnvironmentVariable("UNO_PATH", unoPath)
            End If

            Dim loader As XComponentLoader = Nothing
            Dim xLocalContext = Bootstrap.bootstrap()
            Dim xRemoteFactory = CType(xLocalContext.getServiceManager(), XMultiServiceFactory)
            loader = CType(xRemoteFactory.createInstance("com.sun.star.frame.Desktop"), XComponentLoader)
            Return loader
        Catch ex As Exception
            Console.Write(ex.StackTrace)
        End Try
        Return Nothing
    End Function

    Private sync As Object = New Object()

    Public Sub Convert(ByVal from As String, ByVal toPdf As String)
        Try
            If Not File.Exists(from) Then
                Console.Write(String.Format("Can't find input file {0}", from))
                Exit Sub
            End If

            Dim pv = New PropertyValue(2) {}
            pv(0) = New PropertyValue With
            {
                .Name = "Hidden",
                .Value = New uno.Any(True)
            }

            SyncLock sync
                Dim xComponent As XComponent
                Try
                    xComponent = Loader.loadComponentFromURL("file:///" & from.Replace("\"c, "/"c), "_blank", 0, pv)
                Catch __unusedDisposedException1__ As DisposedException
                    Init()
                    xComponent = Loader.loadComponentFromURL("file:///" & from.Replace("\"c, "/"c), "_blank", 0, pv)
                End Try

                Dim doc = CType(xComponent, XSpreadsheetDocument)
                Dim ox As XSpreadsheets = CType(doc.getSheets(), XSpreadsheets)
                Dim oxx As XIndexAccess = CType(ox, XIndexAccess)
                Dim oXlsSheet As XSpreadsheet = Nothing
                If sheetIndex > -1 Then oXlsSheet = CType(oxx.getByIndex(sheetIndex).Value, XSpreadsheet)
                If Not String.IsNullOrEmpty(sheetName) Then oXlsSheet = CType(ox.getByName(sheetName).Value, XSpreadsheet)
                Dim oPrintArea = CType(oXlsSheet, XPrintAreas).getPrintAreas()
                Dim oRange = Nothing
                If oPrintArea.Count > 0 Then
                    oRange = oXlsSheet.getCellRangeByPosition(oPrintArea(0).StartColumn, oPrintArea(0).StartRow,
                                                              oPrintArea(0).EndColumn, oPrintArea(0).EndRow)
                End If

                Dim xStorable = CType(xComponent, XStorable)

                pv(0).Name = "FilterName"
                Select Case Path.GetExtension(from).ToLowerInvariant()
                    Case ".xls", ".xlsx", ".ods"
                        pv(0).Value = New uno.Any("calc_pdf_Export")
                    Case Else
                        pv(0).Value = New uno.Any("writer_pdf_Export")
                End Select

                Dim pv2 = New PropertyValue(0) {}
                If oRange IsNot Nothing Then
                    pv2(0) = New PropertyValue With
                    {
                        .Name = "Selection",
                        .Value = New uno.Any(oRange.GetType(), oRange)
                    }
                End If

                pv(1) = New PropertyValue With
                {
                   .Name = "FilterData",
                   .Value = New uno.Any(pv2.GetType(), pv2)
                }

                xStorable.storeToURL("file:///" & toPdf.Replace("\"c, "/"c), pv)
                Dim xClosable = CType(xComponent, XCloseable)
                xClosable.close(True)
                If File.Exists(toPdf) Then Process.Start(toPdf)
            End SyncLock
        Catch ex As Exception
            Console.Write(ex.Message, ex.StackTrace)
        End Try
    End Sub

End Module
