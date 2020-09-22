VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Directory Monitor"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "File Date Changed"
      Height          =   195
      Index           =   7
      Left            =   3060
      TabIndex        =   15
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1755
   End
   Begin VB.CheckBox Check2 
      Caption         =   "File Accessed"
      Height          =   195
      Index           =   6
      Left            =   3060
      TabIndex        =   14
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Watch Subdirectories"
      Height          =   195
      Left            =   5040
      TabIndex        =   13
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   60
      TabIndex        =   11
      Top             =   3720
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Name Changed"
      Height          =   195
      Index           =   5
      Left            =   1500
      TabIndex        =   10
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Caption         =   "File Modified"
      Height          =   195
      Index           =   4
      Left            =   3060
      TabIndex        =   9
      Top             =   3240
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Files Attributes Changed"
      Height          =   195
      Index           =   3
      Left            =   5040
      TabIndex        =   6
      Top             =   3480
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Files Size Changed"
      Height          =   195
      Index           =   2
      Left            =   5040
      TabIndex        =   5
      Top             =   3240
      Value           =   1  'Checked
      Width           =   1875
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Files Removed"
      Height          =   195
      Index           =   1
      Left            =   1500
      TabIndex        =   4
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1515
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Files Added"
      Height          =   195
      Index           =   0
      Left            =   1500
      TabIndex        =   3
      Top             =   3240
      Value           =   1  'Checked
      Width           =   1515
   End
   Begin VB.ListBox List1 
      Height          =   2400
      ItemData        =   "frmTest.frx":0000
      Left            =   60
      List            =   "frmTest.frx":0002
      TabIndex        =   2
      Top             =   840
      Width           =   7695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Enabled"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Text            =   "c:\documents and settings\alan\desktop\"
      Top             =   240
      Width           =   7035
   End
   Begin VB.Label Label3 
      Caption         =   "Filter"
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   3540
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Directory to watch:"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Change Events:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   660
      Width           =   1215
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents dMon As DirMonDll.Monitor
Attribute dMon.VB_VarHelpID = -1

Private Sub Check1_Click()
    dMon.path = Text1.Text
    dMon.Interval = 3000
    dMon.Enabled = CBool(Check1.Value)
    Check2_Click 0
End Sub

Private Sub Check2_Click(Index As Integer)
    Dim NewChangeType As cEnumChangeType, x As Integer
    
    NewChangeType = NewChangeType + (cChangeFilesAdded * Check2(0).Value)
    NewChangeType = NewChangeType + (cChangeFilesRemoved * Check2(1).Value)
    NewChangeType = NewChangeType + (cChangeInFileSize * Check2(2).Value)
    NewChangeType = NewChangeType + (cChangeInAttributes * Check2(3).Value)
    NewChangeType = NewChangeType + (cChangeInFileDateModified * Check2(4).Value)
    NewChangeType = NewChangeType + (cChangeInFileName * Check2(5).Value)
    NewChangeType = NewChangeType + (cChangeInFileDateAccessed * Check2(6).Value)
    NewChangeType = NewChangeType + (cChangeInFileDateCreated * Check2(7).Value)
    
    dMon.ChangeType = NewChangeType
End Sub

Private Sub Check3_Click()
    If Check3.Value = vbChecked Then
        If MsgBox("If there is a large number of files contained in the" & vbCrLf & _
                  "subdirectories, then dir monitor will not work properly, or at all." & vbCrLf & _
                  "Do you want to continue anyway?", vbQuestion + vbYesNo, "Please Respond") = vbYes Then
            dMon.SubDirs = CBool(Check3.Value)
        Else
            Check3.Value = vbUnchecked
        End If
    End If
End Sub

Private Sub dMon_OnChange(FileName As String, ChangeType As DirMonDll.cEnumChangeType)
    Select Case ChangeType
        Case cChangeFilesAdded
            List1.AddItem FileName & " Added"
        Case cChangeFilesRemoved
            List1.AddItem FileName & " Removed"
        Case cChangeInAttributes
            List1.AddItem FileName & " Attributes Changed"
        Case cChangeInFileSize
            List1.AddItem FileName & " Size Changed"
        Case cChangeInFileDateAccessed
            List1.AddItem FileName & " File Accessed"
        Case cChangeInFileDateCreated
            List1.AddItem FileName & " Date Changed"
        Case cChangeInFileDateModified
            List1.AddItem FileName & " File Modified"
        Case cChangeInFileName
            List1.AddItem FileName & " Name Changed"
    End Select
End Sub

Private Sub Form_Load()
    Set dMon = New DirMonDll.Monitor
    Text1.Text = GetSpecialfolder(CSIDL_DESKTOP)
    Text2.Text = dMon.Filter
End Sub



Private Sub Text2_LostFocus()
    dMon.Filter = Text2.Text
End Sub


