VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExportAllImg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export All Images"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5220
   Begin MSComctlLib.ProgressBar prgStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   3810
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdExportAll 
      Caption         =   "Export"
      Enabled         =   0   'False
      Height          =   375
      Left            =   983
      TabIndex        =   12
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdGetPath 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   2175
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Export Image Type"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox cboImgSubMinType 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1080
         Width           =   2895
      End
      Begin VB.ComboBox cboImgSubType 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   2895
      End
      Begin VB.ComboBox cboImgType 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Image Sub Min Type :"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1110
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Image Sub Type :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   750
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Image Type :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   390
         Width           =   1215
      End
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   4935
   End
   Begin VB.Label Label5 
      Caption         =   "Export To :"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblImgNum 
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Number of Images to Export:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   2175
   End
End
Attribute VB_Name = "frmExportAllImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sPath As String
Dim ExpImgRs As ADODB.Recordset
Dim NumOfPics As Integer

Private Sub cboImgSubType_GotFocus()
cboImgSubMinType.Clear
End Sub

Private Sub cboImgSubType_LostFocus()
LoadImgSubMinType
End Sub

Private Sub cboImgType_GotFocus()
cboImgSubType.Clear

End Sub

Private Sub cboImgType_LostFocus()
LoadImgSubTypes
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdExportAll_Click()
Dim FilePath As String
Dim FileName As String
Dim j As Integer

If cboImgType.Text <> "" And cboImgSubType.Text = "" And cboImgSubMinType.Text = "" Then
 FileName = cboImgType.Text
ElseIf cboImgSubType.Text <> "" And cboImgSubMinType.Text = "" Then
 FileName = cboImgSubType.Text
ElseIf cboImgType.Text <> "" And cboImgSubType.Text <> "" And cboImgSubMinType.Text <> "" Then
 FileName = cboImgSubMinType.Text
End If
If ExpImgRs.EOF = True Or ExpImgRs.BOF = True Then
 ExpImgRs.MoveFirst
End If
prgStatus.Min = 0
prgStatus.Max = NumOfPics
lblStatus.Caption = "Exporting Started"
DoEvents
For j = 1 To NumOfPics
 With ExpImgRs
    FilePath = sPath & "\" & FileName & j & "." & !ImgExt
    ReadFromDba ExpImgRs.Fields("Image"), FilePath, ExpImgRs!comp
 End With
 prgStatus.Value = prgStatus.Value + 1
ExpImgRs.MoveNext
Next
lblStatus.Caption = "Export Completed"
DoEvents
End Sub

Private Sub cmdGetPath_Click()
sPath = fBrowseForFolder(Me.hwnd, "Select folder to display as thumbnails")
If sPath <> "" Then
 sPath = sPath & "\"
 txtPath.Text = sPath
 cmdExportAll.Enabled = True
GetImageCount
lblImgNum.Caption = NumOfPics
End If

End Sub

Private Sub Form_Load()
Me.Height = 4455
Me.Width = 5370
Me.Move (frmMain.Width / 2) - (Me.Width / 2), (frmMain.Height / 2) - ((Me.Height / 2) + 380)
LoadImgTypes
End Sub

Private Sub LoadImgTypes()
Dim ImgType As ADODB.Recordset
Set ImgType = New ADODB.Recordset
ImgType.Open "Select imagetype from imagetype where seclvl <=" & SecLevel & " group by imagetype", DB, adOpenStatic, adLockOptimistic
If ImgType.RecordCount <> 0 Then
  ImgType.MoveFirst
  Do While Not ImgType.EOF
   cboImgType.AddItem (ImgType("ImageType"))
   ImgType.MoveNext
  Loop
End If
End Sub

Private Sub LoadImgSubTypes()
Dim ImgSubType As ADODB.Recordset
Dim SqlStr As String
Set ImgSubType = New ADODB.Recordset
SqlStr = "Select distinct SubType from imagefile where imagetype ='" & cboImgType.Text & "' group by subtype"
ImgSubType.Open SqlStr, DB, adOpenStatic, adLockOptimistic
If ImgSubType.RecordCount <> 0 Then
 ImgSubType.MoveFirst
  Do While Not ImgSubType.EOF
   cboImgSubType.AddItem (ImgSubType("SubType"))
   ImgSubType.MoveNext
  Loop
End If
End Sub

Private Sub LoadImgSubMinType()
Dim ImgSubMinType As ADODB.Recordset
Dim SqlStr As String
Set ImgSubMinType = New ADODB.Recordset
SqlStr = "Select distinct SubMinType from imagefile where subtype ='" & cboImgSubType.Text & "' group by submintype"
ImgSubMinType.Open SqlStr, DB, adOpenStatic, adLockOptimistic
If ImgSubMinType.RecordCount <> 0 Then
 ImgSubMinType.MoveFirst
  Do While Not ImgSubMinType.EOF
   cboImgSubMinType.AddItem (ImgSubMinType("SubminType"))
   ImgSubMinType.MoveNext
  Loop
End If
End Sub

Private Sub GetImageCount()
Dim SqlStr As String
Set ExpImgRs = New ADODB.Recordset
SqlStr = "Select * from imagefile where imagetype = '" & cboImgType.Text & _
"' and subtype = '" & cboImgSubType.Text & "' and submintype ='" & cboImgSubMinType.Text & "'"
ExpImgRs.Open SqlStr, DB, adOpenStatic, adLockOptimistic

If ExpImgRs.RecordCount <> 0 Then
 NumOfPics = ExpImgRs.RecordCount
 

End If


End Sub
