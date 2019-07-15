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

    Private input As String
    Private output As String

    Sub Main(args As String())
        Try
            If args.Length = 2 Then
                input = args(0)
                output = args(1)
                Convert(input, output)
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
            Dim urePath = String.Empty
            Dim path As String = String.Empty

            If Environment.Is64BitOperatingSystem Then
                unoPath = Configuration.ConfigurationManager.AppSettings("UNO_PATH").ToString()
                urePath = Configuration.ConfigurationManager.AppSettings("URE_PATH").ToString()
            Else
                unoPath = Configuration.ConfigurationManager.AppSettings("UNO_PATH_86x").ToString()
                urePath = Configuration.ConfigurationManager.AppSettings("URE_PATH_86x").ToString()
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
                Dim oXlsSheet As XSpreadsheet = CType(oxx.getByIndex(1).Value, XSpreadsheet)

                Dim oPrintArea = CType(oXlsSheet, XPrintAreas).getPrintAreas()
                Dim oRange = oXlsSheet.getCellRangeByPosition(oPrintArea(0).StartColumn,
                                                              oPrintArea(0).StartRow,
                                                              oPrintArea(0).EndColumn,
                                                              oPrintArea(0).EndRow)

                Dim xStorable = CType(xComponent, XStorable)

                pv(0).Name = "FilterName"
                Select Case Path.GetExtension(from).ToLowerInvariant()
                    Case ".xls", ".xlsx", ".ods"
                        pv(0).Value = New uno.Any("calc_pdf_Export")
                    Case Else
                        pv(0).Value = New uno.Any("writer_pdf_Export")
                End Select

                Dim pv2 = New PropertyValue(0) {}
                pv2(0) = New PropertyValue With
                {
                    .Name = "Selection",
                    .Value = New uno.Any(oRange.GetType(), oRange)
                }

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