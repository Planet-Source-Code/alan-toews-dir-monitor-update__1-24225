VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Monitor"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum cEnumChangeType
    cChangesInAny = &HFF            'all changes
    cChangeInFileSize = &H1         'file size change
    cChangeInAttributes = &H2       'change in any attribute
    cChangeFilesAdded = &H4         'file added
    cChangeFilesRemoved = &H8       'file removed
    cChangeInFileName = &H10        'file name change
    cChangeInFileDateCreated = &H20 'date created change
    cChangeInFileDateAccessed = &H40 'file accessed
    cChangeInFileDateModified = &H80 'file modified
    cChangeInFileDate = &HE0        'all date changes
End Enum

Private WithEvents cFilNew As clsFiles, _
        WithEvents cFilComp As clsFiles
Attribute cFilNew.VB_VarHelpID = -1
Attribute cFilComp.VB_VarHelpID = -1
        
'Privately, I began using the Common Control Replacement Project's High resolution timer.
'You can download this control freely from www.mvps.org/ccrp/.
'while you're there, check out their other controls.
'They've got some excellent stuff.

'For distribution purposes, I've switched to a slightly modified
'timer class from vbaccellerator.com, since the source code was availabe for it.
'I changed the event name, and added an Enabled property to make it more logical
'to my way of thinking.

'I've left the references to the ccrp timer in case you want to try it,
'but the vbAccelerator timer does everything this project needs.
Private WithEvents CTimer As CTimer 'ccrpTimer
Attribute CTimer.VB_VarHelpID = -1

Private cPath As String
Private cChanges As cEnumChangeType
Private cInitializing As Boolean
Private cEnabled As Boolean
Private cInterval As Long
Private mFilter As String
Private mSubDirs As Boolean


Public Event OnError(ByVal Number As Integer, ByVal Description As String)
Public Event OnChange(FileName As String, ChangeType As cEnumChangeType)
Public Property Get ChangeType() As cEnumChangeType
    ChangeType = cChanges
End Property
Public Property Get Filter() As String
    Filter = mFilter
End Property
Public Property Let Filter(newval As String)
    mFilter = newval
    cFilComp.Filter = mFilter
    cFilNew.Filter = mFilter
End Property
Public Property Get SubDirs() As Boolean
    SubDirs = mSubDirs
    cFilComp.SubDirs = mSubDirs
    cFilNew.SubDirs = mSubDirs
End Property
Public Property Let SubDirs(newval As Boolean)
    Dim tmp As Boolean
    If mSubDirs <> newval Then
        tmp = CTimer.Enabled
        mSubDirs = newval
        CTimer.Enabled = False
        Set cFilComp = New clsFiles
        Set cFilNew = New clsFiles
        cFilComp.SubDirs = mSubDirs
        cFilNew.SubDirs = mSubDirs
        cFilComp.Filter = mFilter
        cFilNew.Filter = mFilter
        cFilComp.path = cPath
        cFilNew.path = cPath
        CTimer.Enabled = tmp
    End If
    
End Property

Public Property Let ChangeType(newval As cEnumChangeType)
    cChanges = newval
End Property

Public Property Get path() As String
    path = cPath
End Property
Public Property Let path(newval As String)
    
    cPath = Trim(Replace(newval & "\ ", "\\ ", "\"))
    cFilComp.path = cPath
    cFilNew.path = cPath
End Property
Public Property Get Files() As clsFiles
    Set Files = cFilComp
End Property
Public Property Get Interval() As Long
    
    Interval = cInterval
     
End Property
Public Property Let Interval(newval As Long)
    cInterval = newval
    If cEnabled Then CTimer.Interval = cInterval
    
End Property
Public Property Get Enabled() As Boolean
    Enabled = CTimer.Enabled
End Property
Public Property Let Enabled(newval As Boolean)

    If path = "" Then
        CTimer.Enabled = False
    Else
        CTimer.Enabled = newval
    End If
End Property
'ccrp timer declare-Private Sub cTimer_Timer(ByVal Milliseconds As Long)


Private Sub cTimer_Timer()
    cFilNew.Refresh

    On Error GoTo ErrorCode

    Dim fAttribChange As Integer, _
        fnCount As Integer, _
        fcCount As Integer, _
        CheckSize As Boolean, _
        CheckAttrib As Boolean, _
        CheckAdded As Boolean, _
        CheckRemoved As Boolean, _
        CheckingForRemoved As Boolean, _
        CheckingForAdded As Boolean, _
        AddedOrRemoved As Boolean, _
        CheckFileName As Boolean, _
        CheckDateCreated As Boolean, _
        CheckDateModified As Boolean, _
        CheckDateAccessed As Boolean, _
        EventRaised As Boolean, _
        WasRemoved As Boolean, _
        fName As String, _
        x As Integer
    
    'determine which changes we are looking for
    CheckSize = CBool(cChanges And cChangeInFileSize)
    CheckAttrib = CBool(cChanges And cChangeInAttributes)
    CheckAdded = CBool(cChanges And cChangeFilesAdded)
    CheckRemoved = CBool(cChanges And cChangeFilesRemoved)
    CheckSize = CBool(cChanges And cChangeInFileDate)
    CheckFileName = CBool(cChanges And cChangeInFileName)
    fnCount = cFilNew.Count
    fcCount = cFilComp.Count

    'we are checking if any files have been removed
    CheckingForRemoved = True
    'test for changes
    For x = 1 To fcCount
        fName = cFilComp.Item(x).Key
        
        
        'we want to check for all changes, even if the user hasn't asked for it.
        'If any change ocurs, we want to update out compare list.
        'events will only fire for the requested changes
        'check for changes in attributes (also checks for removed files)
        
            'check for changes in attributes (also checks for removed files)
            If cFilComp.Item(fName).Attrib <> cFilNew.Item(fName).Attrib Then
                If Not AddedOrRemoved Then 'raise event, if requested
                    If CheckAttrib Then RaiseEvent OnChange(fName, cChangeInAttributes)
                    EventRaised = True      'set flag so we can update the cFilComp later
                End If
            End If
            'check for changes in size
            If cFilComp.Item(fName).Size <> cFilNew.Item(fName).Size Then
                If Not AddedOrRemoved Then  'raise event, if requested
                    If CheckSize Then RaiseEvent OnChange(fName, cChangeInFileSize)
                    EventRaised = True      'set flag so we can update the cFilComp later
                End If
            End If
            'check for changes in date
            If cFilComp.Item(fName).LastAccess <> cFilNew.Item(fName).LastAccess Then
                If Not AddedOrRemoved Then  'raise event, if requested
                    If CheckSize Then RaiseEvent OnChange(fName, cChangeInFileDateAccessed)
                    EventRaised = True      'set flag so we can update the cFilComp later
                End If
            End If
            If cFilComp.Item(fName).LastMod <> cFilNew.Item(fName).LastMod Then
                If Not AddedOrRemoved Then  'raise event, if requested
                    If CheckSize Then RaiseEvent OnChange(fName, cChangeInFileDateModified)
                    EventRaised = True      'set flag so we can update the cFilComp later
                End If
            End If
            If cFilComp.Item(fName).Created <> cFilNew.Item(fName).Created Then
                If Not AddedOrRemoved Then  'raise event, if requested
                    If CheckSize Then RaiseEvent OnChange(fName, cChangeInFileDateCreated)
                    EventRaised = True      'set flag so we can update the cFilComp later
                End If
            End If
        
        
        'reset marker variables.
        'these prevent false events such as
        'cChangeInAttributes, if file is deleted
        'and also prevent cChangeFilesRemoved from occuring twice
        CheckingForRemoved = True
        AddedOrRemoved = False
    Next
    
    'check for new files
    CheckingForRemoved = False
    CheckingForAdded = True
    For x = 1 To fnCount
        fName = cFilNew.Item(x).Key
        If cFilNew.Item(fName).FileName <> cFilComp.Item(fName).FileName Then _
            If CheckFileName Then RaiseEvent OnChange(fName, cChangeInFileName)
    Next
    Debug.Print CTimer.Interval
    'refresh the comparison list if any changes occured
    If EventRaised Then cFilComp.Refresh
    Exit Sub
ErrorCode:
    If Err.Number = 91 Then
        'if a file is not found, then it has either been added or deleted
        'We can raise the appropriate event depending on our marker variables
        If CheckingForRemoved Then
            If CheckRemoved Then RaiseEvent OnChange(fName, cChangeFilesRemoved)
            'WasRemoved = True
            CheckingForRemoved = False
        ElseIf CheckingForAdded Then
            'If WasRemoved And CheckFileName Then RaiseEvent OnChange(fName, cChangeInFileName)
            If CheckAdded Then RaiseEvent OnChange(fName, cChangeFilesAdded)
            
        End If
        'set flags so we refresh cFilComp later,
        'and so we don't misfire an attribute or size event
        'on the resume command.
        EventRaised = True
        AddedOrRemoved = True
        Debug.Print "add remove"
        Resume Next
    End If
    RaiseEvent OnError(Err.Number, Err.Description)

End Sub

Private Sub cFilComp_OnError(Number As Long, Description As String)
    RaiseEvent OnError(Number, Description)
End Sub

Private Sub cFilNew_OnError(Number As Long, Description As String)
    RaiseEvent OnError(Number, Description)
End Sub

Private Sub Class_Initialize()

    'initialize all our called objects.
    Set cFilNew = New clsFiles
    Set cFilComp = New clsFiles

    'I've switched to an api timer class, and I've put it in another dll
    'for portability.
    Set CTimer = New CTimer 'ccrpTimer
    
    
    mFilter = cFilComp.Filter
    'initially, the class will look for all possible changes.
    'cChanges = cChangesInAny
    
    'make sure the interval is set to a useable default value
    CTimer.Enabled = False
    CTimer.Interval = 2000
    
    
End Sub

Private Sub Class_Terminate()

    'clean up behind us.
    Set cFilNew = Nothing
    Set cFilComp = Nothing

    Set CTimer = Nothing
    

End Sub


