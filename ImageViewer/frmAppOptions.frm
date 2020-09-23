VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fleet Control Settings"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6660
   Begin VB.Frame Frame1 
      Caption         =   "Compression Level"
      Height          =   3135
      Left            =   150
      TabIndex        =   5
      Top             =   960
      Width           =   6375
      Begin MSComctlLib.Slider sldCompresslvl 
         Height          =   2775
         Left            =   2025
         TabIndex        =   6
         Top             =   240
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   4895
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   1
         Min             =   1
         Max             =   6
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label Label7 
         Caption         =   "Medium Compression (Normal)"
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "High Compression (Slow)"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Light Compression (Fast)"
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Low Compression (Fastest)"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "No Compression"
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Highest Compression (Slowest)"
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdChangeDbLocation 
      Caption         =   "..."
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtDBPath 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   "Database Location :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmAppOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NewDBPath As String
Private Sub cmdChangeDbLocation_Click()
Dim strPath As String
Dim FileFound As Boolean
'If SecLevel >= SecSldLvl(13) Then
 strPath = fBrowseForFolder(frmMain.hwnd, "Select Database Path")
 txtDBPath.Text = strPath & "\ImageDBA.mdb"
If FileEx(txtDBPath.Text) <> True Then
 MsgBox "Database Could no be Found", vbCritical, "Database Error"
 txtDBPath.Text = dbPath
 cmdOk.Enabled = False
ElseIf FileEx(txtDBPath.Text) = True Then
 NewDBPath = txtDBPath
 cmdOk.Enabled = True
End If
'ElseIf SecLevel < SecSldLvl(13) Then
' MsgBox "Security Level to low to perform this action", vbExclamation, "Security Warning"
'End If

End Sub

Private Sub cmdClose_Click()
Unload Me
Me.Hide
End Sub

Private Sub cmdOk_Click()
SaveAppSet
Unload Me
End Sub

Private Sub Form_Load()
Me.Height = 5010
Me.Width = 6750
Me.Move (frmMain.Width / 2) - (Me.Width / 2), (frmMain.Height / 2) - ((Me.Height / 2) + 770)
LoadAppSet
End Sub

Private Sub SaveAppSet()
Select Case sldCompresslvl.Value
    Case 1
       CompLevel = 9
    Case 2
       CompLevel = 6
    Case 3
       CompLevel = -1
    Case 4
       CompLevel = 3
    Case 5
       CompLevel = 1
    Case 6
       CompLevel = 0
End Select

CreateKey ("HKEY_CURRENT_USER\Software\ImageViewer")
SetStringValue "HKEY_CURRENT_USER\Software\ImageViewer", "DBLocation", dbPath
SetStringValue "HKEY_CURRENT_USER\Software\ImageViewer", "CompLevel", Str(CompLevel)
End Sub

Private Sub LoadAppSet()
LoadDBPath
txtDBPath.Text = dbPath
Select Case CompLevel
    Case 9
       sldCompresslvl.Value = 1
    Case 6
       sldCompresslvl.Value = 2
    Case -1
       sldCompresslvl.Value = 3
    Case 3
       sldCompresslvl.Value = 4
    Case 1
       sldCompresslvl.Value = 5
    Case 0
       sldCompresslvl.Value = 6
End Select
End Sub

