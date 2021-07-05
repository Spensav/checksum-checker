Attribute VB_Name = "Module1"
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Public Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function MoveFile Lib "kernel32.dll" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Public Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function VirtualAlloc Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const OPEN_EXISTING = 3
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_END = 2
Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1
Public Const LWA_COLORKEY = &H1
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const MAX_PATH = 260
Public Const SW_SHOWNORMAL = 1

Public Type FileTime
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As FileTime
    ftLastAccessTime    As FileTime
    ftLastWriteTime     As FileTime
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH
    cAlternate          As String * 14
End Type

Type BROWSEINFO
     hOwner As Long
     pidlRoot As Long
     pszDisplayName As String
     lpszTitle As String
     ulFlags As Long
     lpfn As Long
     lParam As Long
     iImage As Long
End Type
Public Function BrowseFolder(ByVal aTitle As String, ByVal aForm As Form) As String
Dim bInfo As BROWSEINFO
Dim rtn&, pidl&, path$, pos%
Dim BrowsePath As String
bInfo.hOwner = aForm.hWnd
bInfo.lpszTitle = aTitle
bInfo.ulFlags = &H1
pidl& = SHBrowseForFolder(bInfo)
path = Space(512)
t = SHGetPathFromIDList(ByVal pidl&, ByVal path)
pos% = InStr(path$, Chr$(0))
BrowseFolder = Left(path$, pos - 1)
If Right$(Browse, 1) = "\" Then
    BrowseFolder = BrowseFolder
    Else
    BrowseFolder = BrowseFolder + "\"
End If
If Right(BrowseFolder, 2) = "\\" Then BrowseFolder = Left(BrowseFolder, Len(BrowseFolder) - 1)
If BrowseFolder = "\" Then BrowseFolder = ""
End Function
Public Function StripNulls(ByVal OriginalStr As String) As String
    If (InStr(OriginalStr, Chr$(0)) > 0) Then
        OriginalStr = Left$(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function
Public Function IsScript(Filename As String) As Boolean
IsScript = False
ext = Split("|vbs|vbe", "|")
For i = 1 To UBound(ext)
If LCase(Right(Filename, 3)) = LCase(ext(i)) Then IsScript = True
Next
End Function

