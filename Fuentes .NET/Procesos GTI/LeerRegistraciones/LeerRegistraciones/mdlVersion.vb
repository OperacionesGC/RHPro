
Module mdlVersion
    ' El ejecutable del que queremos obtener la información
    Private mLocation As String
    '
    Sub Main(ByVal FechaModificacion)
        ' Mostrar información del ejecutable actual
        mLocation = System.Reflection.Assembly.GetExecutingAssembly.Location

        mostrarInfo(FechaModificacion)
    End Sub

    'Imprime la informacion del ensamblado
    Private Sub mostrarInfo(ByVal FechaModificacion)
        FLog.EscribirLinea("-----------------------------------------------------------------")
        FLog.EscribirLinea("Información sobre la versión de: " & EXEName)
        FLog.EscribirLinea("Companía = " & Compania)
        FLog.EscribirLinea("Version = " & FileVersion)        
        FLog.EscribirLinea("Fecha = " & FechaModificacion)
        FLog.EscribirLinea("-----------------------------------------------------------------")

    End Sub

    Private ReadOnly Property FileMajorPart() As Int32
        Get
            Return System.Diagnostics.FileVersionInfo.GetVersionInfo(mLocation).FileMajorPart
        End Get
    End Property
    Private ReadOnly Property FileMinorPart() As Int32
        Get
            Return System.Diagnostics.FileVersionInfo.GetVersionInfo(mLocation).FileMinorPart
        End Get
    End Property
    Private ReadOnly Property FileBuildPart() As Int32
        Get
            Return System.Diagnostics.FileVersionInfo.GetVersionInfo(mLocation).FileBuildPart
        End Get
    End Property
    Private ReadOnly Property FilePrivatePart() As Int32
        Get
            Return System.Diagnostics.FileVersionInfo.GetVersionInfo(mLocation).FilePrivatePart
        End Get
    End Property

    Private ReadOnly Property FileVersion() As String
        Get
            Return System.Diagnostics.FileVersionInfo.GetVersionInfo(mLocation).FileVersion
        End Get
    End Property

    Private ReadOnly Property EXEName() As String
        Get
            Return System.IO.Path.GetFileName(mLocation)
        End Get
    End Property

    Private ReadOnly Property Compania() As String
        Get
            Return System.Diagnostics.FileVersionInfo.GetVersionInfo(mLocation).CompanyName

        End Get
    End Property

End Module
