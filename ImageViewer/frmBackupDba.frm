VERSION 5.00
Begin VB.Form frmBackupDba 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup Database"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6660
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   3255
      Left            =   143
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtPath 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   5655
      End
      Begin VB.CommandButton cmdPath 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   4
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Close"
         Height          =   375
         Left            =   3443
         TabIndex        =   2
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton cmdBackupDba 
         Caption         =   "Backup Database"
         Height          =   375
         Left            =   1643
         TabIndex        =   1
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1650
         TabIndex        =   8
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Backup Destination"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lblDBSize 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   375
         Left            =   923
         TabIndex        =   6
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Backup Database"
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
         Height          =   495
         Left            =   623
         TabIndex        =   3
         Top             =   240
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmBackupDba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbaSize As Long

Private Sub cmdBackupDba_Click()
If SecLevel >= SecSldLvl(16) Then
 If txtPath.Text <> "" Then
  DoBackup dbPath, txtPath.Text
 ElseIf txtPath.Text = "" Then
  MsgBox "Must Supply a Destination Folder for the Backup", , "Backup Database"
 End If
ElseIf SecLevel < SecSldLvl(16) Then
 MsgBox "Security Level to low to perform this action", vbExclamation, "Security Warning"
End If

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdPath_Click()
Dim strTemp As String
strTemp = fBrowseForFolder(Me.hwnd, "Select backup path")
If strTemp <> "" Then
    txtPath = strTemp
End If
End Sub

Private Sub Form_Load()
Me.Height = 3960
Me.Width = 6750
Me.Move (frmMain.Width / 2) - (Me.Width / 2), (frmMain.Height / 2) - ((Me.Height / 2) + 700)
'SetRegion Me.hwnd, Me.Picture, RGB(255, 0, 255)
dbaSize = FileLen(dbPath)
lblDBSize = " Current Database Size : " & Format((dbaSize / 1024) / 1024, "standard") & "MB."
End Sub

