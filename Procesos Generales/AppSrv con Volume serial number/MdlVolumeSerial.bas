Attribute VB_Name = "MdlVolumeSerial"
Option Explicit

Private drv As CDriveInfo

Public Sub Revisar_HDD(ByRef Configurado As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que evalua la seguridad del proceso.
' Autor      : FGZ
' Fecha      : 20/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Long
Dim m_Drive As String

Dim D_Label
Dim D_SerialNumberEx
Dim D_FormatSize_TotalSpace
Dim D_FormatSize_FreeSpace
Dim D_FormatSize_AvailableSpace
Dim D_MaxFilenameLength
Dim D_FileSystem
Dim D_DriveTypeEx
Dim D_FileSystemFlag_fsVolumeCompressed
Dim D_FileSystemFlag_fsFileBasedCompression
Dim D_FileSystemFlag_fsCasePreserved
    
    Set drv = New CDriveInfo
    
    drv.Refresh
    With drv
       If .Present Then
          D_Label = .Label
          D_SerialNumberEx = .SerialNumberEx
          D_FormatSize_TotalSpace = .FormatSize(.TotalSpace, True)
          D_FormatSize_FreeSpace = .FormatSize(.FreeSpace, True)
          D_FormatSize_AvailableSpace = .FormatSize(.AvailableSpace, True)
          D_MaxFilenameLength = .MaxFilenameLength
          D_FileSystem = .FileSystem
          D_DriveTypeEx = .DriveTypeEx
          D_FileSystemFlag_fsVolumeCompressed = .FileSystemFlag(fsVolumeCompressed)
          D_FileSystemFlag_fsFileBasedCompression = .FileSystemFlag(fsFileBasedCompression)
          D_FileSystemFlag_fsCasePreserved = .FileSystemFlag(fsCasePreserved)
       Else
         Flog.writeline "Informacion no Disponible."
         Configurado = True
       End If
    End With
    m_Drive = Left$(CurDir$, 3)
    If HayInstalacionPrevia Then
        Configurado = HDD_Configurado(D_SerialNumberEx, m_Drive)
    Else
        Call Instalar_HDD(D_SerialNumberEx, m_Drive)
        Configurado = True
    End If
    If Not Configurado Then
        Flog.writeline
        Flog.writeline "-------------------------------------------------------------------------------------------------------------"
        Flog.writeline "Instalacion de AppSrv no autorizada. Por favor comunicarse con area de soporte de RHPro. soporte@heidt.com.ar"
        Flog.writeline "-------------------------------------------------------------------------------------------------------------"
        Flog.writeline
    Else
        Flog.writeline
        Flog.writeline "-------------------------------------------------------------------------------------------------------------"
        Flog.writeline "Instalacion de AppSrv autorizada."
        Flog.writeline "-------------------------------------------------------------------------------------------------------------"
        Flog.writeline
    End If
    
    Set drv = Nothing
End Sub


Public Function HDD_Configurado(ByVal Disco, ByVal m_Drive As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que evalua la seguridad del proceso.
' Autor      : FGZ
' Fecha      : 16/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset
Dim Instalacion_Valida As Boolean

    HDD_Configurado = False
    
    StrSql = "SELECT * FROM sis_app "
    StrSql = StrSql & " WHERE snd = '" & Encrypt(EncryptionKey, Disco) & "'"
    StrSql = StrSql & " AND unidad = '" & Encrypt(EncryptionKey, m_Drive) & "'"
    StrSql = StrSql & " AND activo = -1"
    OpenRecordset StrSql, rs
    If rs.EOF Then
        HDD_Configurado = False
    Else
        HDD_Configurado = True
    End If
End Function


Public Sub Instalar_HDD(ByVal Disco, ByVal m_Drive As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que evalua existe instalacion previa del AppSrv.
' Autor      : FGZ
' Fecha      : 20/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset

    StrSql = "SELECT * FROM sis_app "
    StrSql = StrSql & " WHERE snd = '" & Encrypt(EncryptionKey, Disco) & "'"
    OpenRecordset StrSql, rs
    If rs.EOF Then
        StrSql = "INSERT INTO sis_app (snd,unidad,activo) "
        StrSql = StrSql & " VALUES ("
        StrSql = StrSql & "'" & Encrypt(EncryptionKey, Disco) & "'"
        StrSql = StrSql & ",'" & Encrypt(EncryptionKey, m_Drive) & "'"
        StrSql = StrSql & ",-1"
        StrSql = StrSql & " )"
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        StrSql = "UPDATE sis_app SET activo = -1 "
        StrSql = StrSql & ", unidad = '" & Encrypt(EncryptionKey, m_Drive) & "'"
        StrSql = StrSql & " WHERE snd = '" & Encrypt(EncryptionKey, Disco) & "'"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If

End Sub

Public Function HayInstalacionPrevia() As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que evalua existe instalacion previa del AppSrv.
' Autor      : FGZ
' Fecha      : 20/09/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset

    StrSql = "SELECT * FROM sis_app "
    StrSql = StrSql & " WHERE activo = -1"
    OpenRecordset StrSql, rs
    If rs.EOF Then
        HayInstalacionPrevia = False
    Else
        HayInstalacionPrevia = True
    End If
End Function


