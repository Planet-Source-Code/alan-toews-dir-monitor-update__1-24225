VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


Private cFile As clsFile
Private mCol As Collection
Private cUpdating As Boolean
Private mPath As String
Private mSubDirs As Boolean
Private mFileCount As Long
Private mDirCount As Long
Private mFilter As String

Public Event OnError(Number As Long, Description As String)
Public Property Get Busy() As Boolean
    Busy = cUpdating
End Property
Public Property Get Files() As Integer
    Files = mFileCount
End Property
Public Property Get Dirs() As Integer
    Dirs = mDirCount
End Property
Public Property Get SubDirs() As Boolean
    SubDirs = mSubDirs
End Property
Public Property Let SubDirs(newval As Boolean)
    mSubDirs = newval
End Property
Public Property Get Filter() As String
    Filter = mFilter
End Property
Public Property Let Filter(newval As String)
    mFilter = newval
End Property

Public Property Get path() As String
    path = mPath
End Property
Public Property Let path(newval As String)
    Dim tmp As Integer, tmp2 As Integer
    'load the new path into the file collection
    If Not mPath = newval Then
        mPath = newval
        LoadPath mPath, mFilter, mFileCount, mDirCount
    End If
End Property
Public Property Get RecurseSubDirs() As Boolean
    RecurseSubDirs = mSubDirs
End Property
Public Property Let RecurseSubDirs(newval As Boolean)
    mSubDirs = newval
End Property

Function TrimNulls(ByVal txt As String) As String
    If (InStr(txt, vbNullChar) > 0) Then
        txt = Left(txt, InStr(txt, vbNullChar) - 1)
    End If
    TrimNulls = txt
End Function

Public Function Count() As Integer
    Count = mCol.Count
End Function
Public Function Item(Index As Variant) As clsFile
    On Error Resume Next
    Set Item = mCol(Index)
End Function
Public Function Refresh()
    LoadPath mPath, mFilter, mFileCount, mDirCount
End Function
Private Function LoadPath(ByVal path As String, SearchStr As String, FileCount As Long, DirCount As Long, Optional Recursive As Boolean = False)
    'KPD-Team 1999
    'E-Mail: KPDTeam@Allapi.net
    Dim cFil As clsFile
    Dim FileName As String ' Walking filename variable...
    Dim DirName As String ' SubDirectory Name
    Dim dirNames() As String ' Buffer for directory name entries
    Dim nDir As Integer ' Number of directories in this path
    Dim i As Integer ' For-loop counter...
    Dim hSearch As Long ' Search Handle
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Integer
    If Right(path, 1) <> "\" Then path = path & "\"
    ' Search for subdirectories.
    nDir = 0
    ReDim dirNames(nDir)
    Cont = True
    hSearch = FindFirstFile(path & "*", WFD)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
        
        DirName = TrimNulls(WFD.cFileName)
        ' Ignore the current and encompassing directories.
        If (DirName <> ".") And (DirName <> "..") Then
            ' Check for directory with bitwise comparison.
            If GetFileAttributes(path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
                dirNames(nDir) = DirName
                DirCount = DirCount + 1
                nDir = nDir + 1
                ReDim Preserve dirNames(nDir)
            End If
        End If
        Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
        Loop
        Cont = FindClose(hSearch)
    End If
    
    ' Walk through this directory and sum file sizes.
    hSearch = FindFirstFile(path & SearchStr, WFD)
    Cont = True
    If Not Recursive Then Clear_mCol
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            Set cFil = New clsFile
            FileName = TrimNulls(WFD.cFileName)
            If (FileName <> ".") And (FileName <> "..") Then
                LoadPath = LoadPath & (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
                FileCount = FileCount + 1
                cFil.SetFile FileName, path
                mCol.Add cFil, path & FileName
            End If
            Cont = FindNextFile(hSearch, WFD) ' Get next file
        Wend
        Cont = FindClose(hSearch)
    End If
    
    ' If there are sub-directories...
    If nDir > 0 And mSubDirs Then
        ' Recursively walk into them...
        For i = 0 To nDir - 1
            LoadPath = LoadPath + LoadPath(path & dirNames(i) & "\", SearchStr, FileCount, DirCount, True)
        Next
    End If

End Function

Private Function Clear_mCol()
    'clear the collection.
    Do While mCol.Count > 0
        mCol.Remove 1
    Loop
End Function

Private Sub Class_Initialize()
    'initialize the declared objects
    Set mCol = New Collection
    Set cFile = New clsFile
    mFilter = "*.*"
End Sub

Private Sub Class_Terminate()
    'clean up behind us
    Set mCol = Nothing
    Set cFile = Nothing
End Sub
