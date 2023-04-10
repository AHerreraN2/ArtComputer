Public Class SEI_ARTCOMPUTER_WebServices
    Inherits ModuleBase.AbstractModuleWebServiceModel
    Implements ISEI_ARTCOMPUTER_WebServices_Contract

#Region "Constructor / Registro servicio"

    Dim ots As New TraceSwitch("TraceSwitch", "Switch del Traceador")

    Public Overrides Function registerService() As String

        Dim resourceName As String = "SEI_ARTCOMPUTER_objects.xml"

        Trace.WriteLineIf(ots.TraceInfo, "Obteniendo recurso incrustado de módulo [Módulo: " + "SEI_ARTCOMPUTER" + "] [Recurso: " + resourceName + "]")

        If Me.GetType.Assembly.GetManifestResourceInfo(Me.GetType.Assembly.GetName().Name + "." + resourceName) IsNot Nothing Then

            Dim resourceStream As System.IO.Stream = Me.GetType.Assembly.GetManifestResourceStream(Me.GetType, resourceName)
            Dim stringReader As New System.IO.StreamReader(resourceStream)
            Dim resourceContents As String = stringReader.ReadToEnd
            stringReader.Close()
            resourceStream.Close()

            Return resourceContents

            Trace.WriteLineIf(ots.TraceVerbose, "Obtenido recurso incrustado de módulo [Módulo: " + "SEI_ARTCOMPUTER" + "] [Recurso: " + resourceName + "] [Contenido: " + resourceContents + "]")

        Else

            Trace.WriteLineIf(ots.TraceError, "No puedo obtenerse recurso incrustado [Módulo: " + "SEI_ARTCOMPUTER" + "] [Recurso: " + resourceName + "]")
            Return Nothing

        End If

    End Function

    Public Sub New()

    End Sub

#End Region



End Class
