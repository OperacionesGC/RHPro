VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDriveInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function SetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long

'valores retornados por GetDriveType
Private Const DRIVE_UNKNOWN       As Long = 0
Private Const DRIVE_NO_ROOT_DIR   As Long = 1
Private Const DRIVE_REMOVABLE     As Long = 2
Private Const DRIVE_FIXED         As Long = 3
Private Const DRIVE_REMOTE        As Long = 4
Private Const DRIVE_CDROM         As Long = 5
Private Const DRIVE_RAMDISK       As Long = 6

'flags del File system
Private Const FILE_CASE_SENSITIVE_SEARCH      As Long = &H1
Private Const FILE_CASE_PRESERVED_NAMES       As Long = &H2
Private Const FILE_UNICODE_ON_DISK            As Long = &H4
Private Const FILE_PERSISTENT_ACLS            As Long = &H8
Private Const FILE_FILE_COMPRESSION           As Long = &H10
Private Const FILE_VOLUME_QUOTAS              As Long = &H20
Private Const FILE_SUPPORTS_SPARSE_FILES      As Long = &H40
Private Const FILE_SUPPORTS_REPARSE_POINTS    As Long = &H80
Private Const FILE_SUPPORTS_REMOTE_STORAGE    As Long = &H100
Private Const FILE_VOLUME_IS_COMPRESSED       As Long = &H8000
Private Const FILE_SUPPORTS_OBJECT_IDS        As Long = &H10000
Private Const FILE_SUPPORTS_ENCRYPTION        As Long = &H20000

Private Const MAX_PATH As Long = 260

Public Enum DriveTypes
   diUnknown = DRIVE_UNKNOWN
   diNoRootDir = DRIVE_NO_ROOT_DIR
   diRemovable = DRIVE_REMOVABLE
   diFixed = DRIVE_FIXED
   diRemote = DRIVE_REMOTE
   diCDROM = DRIVE_CDROM
   diRAMDisk = DRIVE_RAMDISK
End Enum

Public Enum DriveFlags
   fsCaseSensitiveSearch = FILE_CASE_SENSITIVE_SEARCH
   fsCasePreserved = FILE_CASE_PRESERVED_NAMES
   fsUnicodeOnDisk = FILE_UNICODE_ON_DISK
   fsPersistantACLs = FILE_PERSISTENT_ACLS
   fsFileBasedCompression = FILE_FILE_COMPRESSION
   fsVolumeQuotas = FILE_VOLUME_QUOTAS
   fsVolumeCompressed = FILE_VOLUME_IS_COMPRESSED
   fsSupportsSparseFiles = FILE_SUPPORTS_SPARSE_FILES
   fsSupportsReparsePoints = FILE_SUPPORTS_REPARSE_POINTS
   fsSupportsRemoteStorage = FILE_SUPPORTS_REMOTE_STORAGE
   fsSupportsObjectIDs = FILE_SUPPORTS_OBJECT_IDS
   fsSupportsEncryptedFileSystem = FILE_SUPPORTS_ENCRYPTION
End Enum

'variables
Private m_Drive As String
Private m_Enhanced As Boolean
Private m_Present As Boolean
Private m_AvailSpace As Variant
Private m_FreeSpace As Variant
Private m_TotalSpace As Variant
Private m_VolName As String
Private m_VolSN As Long
Private m_MaxLen As Long
Private m_Flags As Long
Private m_FileSys As String

Private Sub Class_Initialize()
'----------------------------------------------
'   Inicializacion.
'----------------------------------------------

   ' check if enhanced function is available.
   ' only true if 95/OSR2 or NT4 or higher.
   m_Enhanced = Exported("kernel32", "GetDiskFreeSpaceExA")

   'default to current drive and get stats.
   m_Drive = Left$(CurDir$, 3)
   Me.Refresh
End Sub

' ********************************************
'  Public Read/Write Properties
' ********************************************
Public Property Let Drive(ByVal NewDrive As String)
   ' store passed value
   m_Drive = NewDrive
   If InStr(m_Drive, ":") = 2 Then
      m_Drive = UCase$(Left$(m_Drive, 2))
   End If
   
   ' should always end with a backslash, esp. UNC
   m_Drive = IIf(Right$(m_Drive, 1) = "\", m_Drive, m_Drive & "\")
   
   ' update statistics
   Me.Refresh
End Property

Public Property Get Drive() As String
   Drive = m_Drive
End Property

Public Property Let Label(ByVal NewLabel As String)
   ' passing NULL will delete volume label
   If NewLabel = "" Then NewLabel = vbNullString
   
   ' set new label and update stats
   If SetVolumeLabel(m_Drive, NewLabel) Then
      Call UpdateVolumeInfo
   End If
End Property

Public Property Get Label() As String
   Label = m_VolName
End Property

' ********************************************
'  Public Read-Only Properties
' ********************************************
Public Property Get AvailableSpace() As Variant
   ' this value could be critical!
   ' make sure it's updated before return.
   Call UpdateFreeSpace
   AvailableSpace = m_AvailSpace
End Property

Public Property Get DriveType() As DriveTypes
   ' simply return API result.
   DriveType = GetDriveType(m_Drive)
End Property

Public Property Get DriveTypeEx() As String
   ' provide easy display value.
   Select Case GetDriveType(m_Drive)
      Case diUnknown:   DriveTypeEx = "Unknown"
      Case diNoRootDir: DriveTypeEx = "No Root Directory"
      Case diRemovable: DriveTypeEx = "Removable"
      Case diFixed:     DriveTypeEx = "Fixed"
      Case diRemote:    DriveTypeEx = "Remote"
      Case diCDROM:     DriveTypeEx = "CD-ROM"
      Case diRAMDisk:   DriveTypeEx = "RAM Disk"
      Case Else:        DriveTypeEx = "Unknown"
   End Select
End Property

Public Property Get FileSystem() As String
   FileSystem = m_FileSys
End Property

Public Property Get FileSystemFlag(ByVal Flag As DriveFlags) As Boolean
   FileSystemFlag = ((m_Flags And Flag) = Flag)
End Property

Public Property Get FileSystemFlags() As Long
   FileSystemFlags = m_Flags
End Property

Public Property Get FreeSpace() As Variant
   ' this value could be critical!
   ' make sure it's updated before return.
   Call UpdateFreeSpace
   FreeSpace = m_FreeSpace
End Property

Public Property Get MaxFilenameLength() As Long
   MaxFilenameLength = m_MaxLen
End Property

Public Property Get Present() As Boolean
   Present = m_Present
End Property

Public Property Get SerialNumber() As String
   SerialNumber = m_VolSN
End Property

Public Property Get SerialNumberEx() As String
   Dim sTemp As String
   ' provide easy display value.
   If m_VolSN <> 0 Then
      sTemp = Right$("00000000" & Hex(m_VolSN), 8)
      SerialNumberEx = Left$(sTemp, 4) & "-" & Right$(sTemp, 4)
   End If
End Property

Public Property Get TotalSpace() As Variant
   TotalSpace = m_TotalSpace
End Property

' ********************************************
'  Public Methods
' ********************************************
Public Function FormatSize(ByVal Size As Variant, Optional ByVal LongDisplay As Boolean = False) As String
   Dim sRet As String
   Dim nSize As Currency
   Const KB& = 1024
   Const MB& = KB * KB
   Const GB@ = MB * KB
   Const TB@ = GB * KB
   
   ' Return size of file in kilobytes.
   If Size < KB Then
      sRet = Format(Size, "#,##0") & " byte"
      If Size <> 1 Then sRet = sRet & "s"
   Else
      ' Divide to obtain base number.
      If Size >= TB Then
         nSize = Size / TB
         sRet = " TB"
      ElseIf Size >= GB Then
         nSize = Size / GB
         sRet = " GB"
      ElseIf Size >= MB Then
         nSize = Size / MB
         sRet = " MB"
      Else
         nSize = Size / KB
         sRet = " KB"
      End If
      ' Display with suitable number of decimals.
      If nSize >= 100 Then
         sRet = Format(nSize, "0") & sRet
      ElseIf nSize >= 10 Then
         sRet = Format(nSize, "0.0") & sRet
      Else
         sRet = Format(nSize, "0.00") & sRet
      End If
   End If
      
   ' Return more detailed string on request.
   If LongDisplay Then
      If Size >= KB Then
         sRet = sRet & " (" & Format(Size, "#,##0") & " bytes)"
      End If
   End If
   FormatSize = sRet
End Function

Public Sub Refresh()
   Call UpdateFreeSpace
   If m_Present Then Call UpdateVolumeInfo
End Sub

' ********************************************
'  Private Methods
' ********************************************
Private Function Exported(ByVal ModuleName As String, ByVal ProcName As String) As Boolean
   Dim hModule As Long
   Dim lpProc As Long
   Dim FreeLib As Boolean
   
   ' check first to see if the module is already
   ' mapped into this process.
   hModule = GetModuleHandle(ModuleName)
   If hModule = 0 Then
      ' need to load module into this process.
      hModule = LoadLibrary(ModuleName)
      FreeLib = True
   End If
   
   ' if the module is mapped, check procedure
   ' address to verify it's exported.
   If hModule Then
      lpProc = GetProcAddress(hModule, ProcName)
      Exported = (lpProc <> 0)
   End If
   
   ' unload library if we loaded it here.
   If FreeLib Then Call FreeLibrary(hModule)
End Function

Private Function TrimNull(ByVal StrIn As String) As String
   Dim nul As Long
   '
   ' Truncate input string at first null.
   ' If no nulls, perform ordinary Trim.
   '
   nul = InStr(StrIn, vbNullChar)
   Select Case nul
      Case Is > 1
         TrimNull = Left(StrIn, nul - 1)
      Case 1
         TrimNull = ""
      Case 0
         TrimNull = Trim(StrIn)
   End Select
End Function

Private Sub UpdateFreeSpace()
   ' reset values in case drive is missing
   m_AvailSpace = CDec(0)
   m_TotalSpace = CDec(0)
   m_FreeSpace = CDec(0)
   m_Present = False
   
   ' check if enhanced function is available.
   If m_Enhanced Then
      Dim cAvail As Currency
      Dim cTotal As Currency
      Dim cFree As Currency
      ' return available bytes, as that's more important to know
      ' than total free bytes if they differ.
      If GetDiskFreeSpaceEx(m_Drive, cAvail, cTotal, cFree) Then
         m_AvailSpace = CDec(cAvail * 10000)
         m_TotalSpace = CDec(cTotal * 10000)
         m_FreeSpace = CDec(cFree * 10000)
         m_Present = True
      End If
   
   Else  ' enhanced function not exported.
      Dim nSecPerClus As Long
      Dim nBytPerSec As Long
      Dim nFreeClus As Long
      Dim nTotalClus As Long
      ' do the math to return total free bytes.
      If GetDiskFreeSpace(m_Drive, nSecPerClus, nBytPerSec, nFreeClus, nTotalClus) Then
         m_TotalSpace = CDec(nSecPerClus * nBytPerSec * nTotalClus)
         m_FreeSpace = CDec(nSecPerClus * nBytPerSec * nFreeClus)
         m_AvailSpace = m_FreeSpace
         m_Present = True
      End If
   End If
End Sub

Private Sub UpdateVolumeInfo()
   ' resize buffers to max
   m_VolName = Space(MAX_PATH)
   m_FileSys = Space(MAX_PATH)
   
   ' make call, and clean up returns
   If GetVolumeInformation(m_Drive, m_VolName, MAX_PATH, m_VolSN, m_MaxLen, m_Flags, m_FileSys, MAX_PATH) Then
      m_VolName = TrimNull(m_VolName)
      m_FileSys = TrimNull(m_FileSys)
      m_Present = True
   Else
      m_VolName = ""
      m_VolSN = 0
      m_MaxLen = 0
      m_Flags = 0
      m_FileSys = ""
      m_Present = False
   End If
End Sub

