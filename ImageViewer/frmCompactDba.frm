VERSION 5.00
Begin VB.Form frmCompactDba 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compact Database"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6570
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   3480
         Width           =   1575
      End
      Begin VB.CommandButton cmdCompactDba 
         Caption         =   "Compact Database"
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton cmdBackupdba 
         Caption         =   "Backup Database"
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   3480
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "frmCompactDba.frx":0000
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   3120
         Width           =   5655
      End
      Begin VB.Label lblNewSize 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   2760
         Width           =   5295
      End
      Begin VB.Label lblDBSize 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   1920
         Width           =   4575
      End
      Begin VB.Label lblFreeSpace 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   1560
         Width           =   4695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Compact Database"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmCompactDba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbaSize As Long


Private Sub cmdBackupDba_Click()
frmBackupDba.Show
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdCompactDba_Click()
Dim NewPath As String
Dim StringPos As Long

On Error GoTo err

If SecLevel >= SecSldLvl(18) Then

StringPos = InStr(1, dbPath, "\imagedba.mdb", vbTextCompare)
NewPath = Left(dbPath, StringPos - 1)

If MsgBox("Are you sure", vbYesNo) = vbYes Then
Screen.MousePointer = vbHourglass
lblStatus.Caption = "Started Compacting"
DoEvents
      DB.Close
      'DEnv.conImage.Close
    If Dir(NewPath & "\CompactDBA.mdb") <> "" Then
      Kill NewPath & "\CompactDBA.mdb"
    End If
    DBEngine.CompactDatabase NewPath & "\Imagedba.mdb", NewPath & "\CompactDBA.mdb", , , ";pwd=" & dbPass
    Kill NewPath & "\ImageDBA.mdb"
    Name NewPath & "\CompactDBA.mdb" As NewPath & "\ImageDBA.mdb"
    'PathName = App.Path & "\Contract.MDB"
    'On Error GoTo err
    dbaSize = FileLen(dbPath)
    lblNewSize = "Compacted Database size : " & FormatNumber((dbaSize / 1024) / 1024, 3) & "MB."
    OpenDB
    'ConnectDenv
 Screen.MousePointer = vbDefault
 lblStatus.Caption = "Compacting Completed"
 End If

ElseIf SecLevel < SecSldLvl(18) Then
 MsgBox "Security Level to low to perform this action", vbExclamation, "Security Warning"
End If

err:
 If err.Number = 3356 Then
   Screen.MousePointer = vbDefault
   lblStatus.Caption = "Compacting error"

   'lblStatus.Caption = "Compact Error, Restart Computer"
   MsgBox "Error occured while trying to compact database Restart your Computer and try again", vbExclamation
   'MsgBox "Error - " & err.Description
   Exit Sub
End If

End Sub

Private Sub Form_Activate()
lblDBSize = "Current Database size: " & FormatNumber((dbaSize / 1024) / 1024, 3) & "MB."

End Sub

Private Sub Form_Load()
Dim fs, d, S
Dim drvpath As String
Dim DiskfSpace As Long

Me.Height = 4590
Me.Width = 6660
Me.Move (frmMain.Width / 2) - (Me.Width / 2), (frmMain.Height / 2) - ((Me.Height / 2) + 700)
drvpath = dbPath

Set fs = CreateObject("Scripting.FileSystemObject")
Set d = fs.GetDrive(fs.GetDriveName(drvpath))
DiskfSpace = d.freeSpace / 1024 / 1024

S = "Drive " & Left(App.Path, 1) & " has "
lblFreeSpace = S & FormatNumber(DiskfSpace, 0) & "MB free"
On Error GoTo err
dbaSize = FileLen(dbPath)
If DiskfSpace * 1024 * 1024 < dbaSize Then
  lblNewSize = "Not enough space to compact database clear some space on drive " & Left(App.Path, 1)
  cmdCompactDba.Enabled = False
End If
err:
Exit Sub
End Sub

Private Sub lblDriveSize_Click()

End Sub
