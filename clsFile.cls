VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FO_DELETE = &H3

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare Function GetFileType Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long

Public Key As String
Public FileName As String
Public path As String
Public Attrib As VbFileAttribute
Public Size As Long
Public LastMod As String
Public LastAccess As String
Public Created As String
Public Directory As Boolean
Public ReadOnly As Boolean
Public Archive As Boolean
Public System As Boolean
Public Hidden As Boolean


Public Function SetFile(fName As String, Optional fPath As String = "") As Boolean
    On Error GoTo ErrorCode
    Dim hFile As Long, nSize As Long, sSave As String
    Dim FCT As FILETIME, FAT As FILETIME, LWT As FILETIME, ST As SYSTEMTIME
    'get the details about the file.
    If fPath = "" Then fPath = App.path
    fPath = Trim(Replace(fPath & "\ ", "\\ ", "\"))
    
    'open the file
    hFile = CreateFile(fPath & fName, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    
    
    FileName = fName
    path = fPath
    Key = fPath & FileName
    
    'get the attributes
    Attrib = GetFileAttributes(fPath & fName)
    Directory = CBool(Attrib And vbDirectory)
    ReadOnly = CBool(Attrib And vbReadOnly)
    Archive = CBool(Attrib And vbArchive)
    System = CBool(Attrib And vbSystem)
    Hidden = CBool(Attrib And vbHidden)
    
    'get the filesize
    nSize = GetFileSize(hFile, nSize)
    If nSize < 0 Then nSize = 0
    Size = nSize
    
    'get the file times
    GetFileTime hFile, FCT, FAT, LWT
    
    FileTimeToSystemTime FCT, ST
    Created = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & " " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond & "." & ST.wMilliseconds
    FileTimeToSystemTime LWT, ST
    LastMod = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & " " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond & "." & ST.wMilliseconds
    FileTimeToSystemTime FAT, ST
    LastAccess = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & " " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond & "." & ST.wMilliseconds
    
    CloseHandle hFile
    SetFile = True
    
    Exit Function
ErrorCode:
    SetFile = False
    CloseHandle hFile
End Function

Public Function IsAttrib(TestAttrib As VbFileAttribute) As Boolean
    'Check if a file has the specified attribute
    IsAttrib = CBool(Attrib And TestAttrib)
End Function
