Imports unoidl.com.sun.star.lang
Imports unoidl.com.sun.star.util
Imports System.IO
Imports System.Environment
Imports unoidl.com.sun.star.uno
Imports uno.util
Imports unoidl.com.sun.star.frame

Module Module1

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
            Dim confFile = Configuration.ConfigurationManager.OpenExeConfiguration(Configuration.ConfigurationUserLevel.None)
            Dim setting = confFile.AppSettings.Settings
            Dim unoPath = Configuration.ConfigurationManager.AppSettings("UNO_PATH").ToString() '"C:\Program Files (x86)\LibreOffice 3.6\program"
            Dim urePath = Configuration.ConfigurationManager.AppSettings("URE_PATH").ToString() '"C:\Program Files (x86)\LibreOffice 3.6\URE\bin"
            Dim path As String = String.Empty

            'Utilizado para versoes > 4 onde nao tem a pasta URE
            'SetEnvironmentVariable("UNO_PATH", unoPath, EnvironmentVariableTarget.Process)
            'SetEnvironmentVariable("PATH", GetEnvironmentVariable("PATH") + ";" + unoPath, EnvironmentVariableTarget.Process)

            'Utilizado para versos < 4

            'Versões que não existe a pasta 'URE' o PATH ficará na pasta 'program'
            'Tendo a pasta URE o PATH será a pasta 'URE'

            If String.IsNullOrEmpty(urePath) Then
                path = String.Format("{0};{1}", System.Environment.GetEnvironmentVariable("PATH"), unoPath)
                Environment.SetEnvironmentVariable("PATH", path)
                Environment.SetEnvironmentVariable("UNO_PATH", unoPath)
            Else
                path = String.Format("{0};{1}", System.Environment.GetEnvironmentVariable("PATH"), urePath)
                Environment.SetEnvironmentVariable("PATH", path)
                Environment.SetEnvironmentVariable("UNO_PATH", unoPath)
            End If



            Dim loader As XComponentLoader = Nothing
            Dim xLocalContext = Bootstrap.bootstrap()
            Dim xRemoteFactory = CType(xLocalContext.getServiceManager(), XMultiServiceFactory)
            loader = CType(xRemoteFactory.createInstance("com.sun.star.frame.Desktop"), XComponentLoader)
            Return loader
        Catch ex As Exception
            Console.Write(ex.Message)
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
            Dim pv = New unoidl.com.sun.star.beans.PropertyValue(0) {}
            pv(0) = New unoidl.com.sun.star.beans.PropertyValue With
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

                Dim xStorable = CType(xComponent, XStorable)
                pv(0).Name = "FilterName"

                Select Case Path.GetExtension(from).ToLowerInvariant()
                    Case ".xls", ".xlsx", ".ods"
                        pv(0).Value = New uno.Any("calc_pdf_Export")
                    Case Else
                        pv(0).Value = New uno.Any("writer_pdf_Export")
                End Select

                xStorable.storeToURL("file:///" & toPdf.Replace("\"c, "/"c), pv)
                Dim xClosable = CType(xComponent, XCloseable)
                xClosable.close(True)
                If File.Exists(toPdf) Then Process.Start(toPdf)
            End SyncLock
        Catch ex As Exception
            Console.Write(ex.Message)
        End Try
    End Sub

End Module
