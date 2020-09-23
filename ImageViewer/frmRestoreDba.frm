VERSION 5.00
Begin VB.Form frmRestoreDba 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Restore Database"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6630
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   3255
      Left            =   128
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   3428
         TabIndex        =   7
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton cmdRestoreDba 
         Caption         =   "Restore Database"
         Height          =   375
         Left            =   1628
         TabIndex        =   6
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CommandButton cmdPath 
         Caption         =   "..."
         Height          =   375
         Left            =   5880
         TabIndex        =   5
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtPath 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   5535
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
         Left            =   1515
         TabIndex        =   9
         Top             =   2400
         Width           =   3600
      End
      Begin VB.Label lblSelectedDba 
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
         Left            =   555
         TabIndex        =   8
         Top             =   2040
         Width           =   5520
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Restore From"
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
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   2655
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
         Height          =   255
         Left            =   495
         TabIndex        =   2
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Restore Database"
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
         Left            =   1095
         TabIndex        =   1
         Top             =   240
         Width           =   4440
      End
   End
End
Attribute VB_Name = "frmRestoreDba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPath_Click()
On Error GoTo Erro
Dim strTemp As String
Dim dbaSize2 As Long
strTemp = fBrowseForFolder(Me.hwnd, "Restore From")
If strTemp <> "" Then
    txtPath = strTemp
    dbaSize2 = FileLen(txtPath & "\ImageDBA.MDB")
    lblSelectedDba = "Selected Backup Database is : " & Format((dbaSize2 / 1024) / 1024, "standard") & "MB."
    cmdRestoreDba.Enabled = True
End If
Erro:
    Select Case err.Number
       Case 53 'File Not Found
          lblSelectedDba = "No Backup at this location"
          cmdRestoreDba.Enabled = False
    End Select

End Sub

Private Sub cmdRestoreDba_Click()
Dim NewPath As String
Dim StringPos As Long
If SecLevel >= SecSldLvl(17) Then
 If MsgBox("Restoring database from location " & txtPath & " will replace existing database files.Do you want to Contunue", vbYesNo) = vbYes Then
  StringPos = InStr(1, dbPath, "\ImageDBA.mdb", vbTextCompare)
  NewPath = Left(dbPath, StringPos - 1)
  DoRestore txtPath.Text, NewPath
  If NoDba = True Then
   MsgBox "Database Restored Click Ok to Exit Program"
   frmRestoreDba.Hide
   Unload frmRestoreDba
  End If
  Else
   lblStatus.Caption = "Database Restore Canceled"
  End If
ElseIf SecLevel < SecSldLvl(17) Then
 MsgBox "Security Level to low to perform this action", vbExclamation, "Security Warning"
End If
End Sub

Private Sub Form_Load()
Dim dbaSize As Long
Me.Height = 3960
Me.Width = 6750
Me.Move (frmMain.Width / 2) - (Me.Width / 2), (frmMain.Height / 2) - ((Me.Height / 2) + 700)
dbaSize = FileLen(dbPath)
lblDBSize = " Current Database Size : " & Format((dbaSize / 1024) / 1024, "standard") & "MB."

End Sub
