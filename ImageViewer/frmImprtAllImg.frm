VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImportAllImg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import All Images"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5250
   Begin VB.OptionButton optComp 
      Caption         =   "Compress all Imported Images"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   2535
   End
   Begin MSComctlLib.ProgressBar prgStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   3840
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
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
      Left            =   4733
      TabIndex        =   11
      Top             =   495
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   173
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Import As Image Type"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   5055
      Begin VB.ComboBox cboImgType 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
      Begin VB.ComboBox cboImgSubType 
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Top             =   720
         Width           =   2895
      End
      Begin VB.ComboBox cboImgSubMinType 
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Image Type :"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Image Sub Type :"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   750
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Image Sub Min Type :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1110
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdImportAll 
      Caption         =   "Import"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblStatus 
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label lblImgNum 
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Total Images at Location :"
      Height          =   255
      Left            =   173
      TabIndex        =   12
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Import From :"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmImportAllImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sPath As String
Dim Pics() As String
Dim NumOfPics As Integer
Dim strPath As String


Private Sub cboImgType_Click()
cboImgSubType.Clear
LoadImgSubTypes
If cboImgType.Text <> "" Then
cmdImportAll.Enabled = True
End If
End Sub

Private Sub cboImgType_GotFocus()
cboImgSubType.Clear
End Sub

Private Sub cboImgType_LostFocus()
LoadImgSubTypes
If cboImgType.Text <> "" Then
cmdImportAll.Enabled = True
End If
End Sub

Private Sub cmdClose_Click()
Unload frmImportAllImg
End Sub

Private Sub cmdGetPath_Click()
sPath = fBrowseForFolder(Me.hwnd, "Select folder to display as thumbnails")
If sPath <> "" Then
frmLoading.Show
 sPath = sPath & "\"
 NumOfPics = ScanFolder(sPath)
 lblImgNum = NumOfPics
 txtPath.Text = sPath
 prgStatus.Min = 0
 prgStatus.Max = NumOfPics
 cboImgType.SetFocus
 LoadImgTypes
Unload frmLoading
End If
End Sub

Private Sub cmdImportAll_Click()
prgStatus.Value = 0
If cboImgType.Text <> "" Then
lblStatus.Caption = "Importing Started"
DoEvents
AddImportsToDBA
DoEvents
lblStatus.Caption = "Import Completed"
End If
End Sub
Private Sub AddImportsToDBA()
Dim NewImgRs As ADODB.Recordset
Dim GetImgID As ADODB.Recordset
Dim r As Integer
Dim PicData As String
Dim FileExt As String
Set NewImgRs = New ADODB.Recordset
NewImgRs.Open "Select * from ImageFile", DB, adOpenStatic, adLockOptimistic

For r = 1 To NumOfPics
 Set GetImgID = New ADODB.Recordset
 GetImgID.Open "Select Max(ImageID) as ImgID from imagefile", DB, adOpenStatic, adLockOptimistic
 PicData = ReadFromFile(sPath & Pics(r), optComp.Value)
 FileExt = Right(Pics(r), 3)
  With NewImgRs
   .AddNew
   If IsNull(GetImgID!ImgID) Then
    !imageid = 0
   Else
    !imageid = GetImgID!ImgID + 1
   End If
   !Image.AppendChunk PicData
   !imagetype = cboImgType.Text
   !subtype = ChgNull(cboImgSubType.Text)
   !subminType = ChgNull(cboImgSubMinType.Text)
   !ImgExt = UCase(FileExt)
   !comp = optComp.Value
   .Update
  End With
  'DoEvents
  prgStatus.Value = prgStatus.Value + 1
  GetImgID.Close
Next
End Sub
Private Sub Form_Load()
Me.Height = 4455
Me.Width = 5370
Me.Move (frmMain.Width / 2) - (Me.Width / 2), (frmMain.Height / 2) - ((Me.Height / 2) + 380)

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

Public Function ScanFolder(sPath As String) As Integer
On Error Resume Next
Dim fs, f, fc, f1
Dim PicNum As Long
Dim ChkExt As Integer
strPath = sPath
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFolder(sPath)
Set fc = f.Files
NumOfPics = CountImages(sPath)
ScanFolder = NumOfPics
If NumOfPics = 1 Then
MsgBox "There was only picture in the folder so it has been opened as normal", vbOKOnly, "Slideshow error"
frmMain.OpenImage (sPath & OnlyFile)
Unload Me
Else
frmLoading.prgLoading.Min = 0
frmLoading.prgLoading.Max = NumOfPics
'NumOfPics = NumOfPics - 1
frmLoading.Caption = "Finding Files, please wait..."
ReDim Pics(NumOfPics) As String
For Each f1 In fc
   DoEvents
   frmLoading.prgLoading.Value = frmLoading.prgLoading.Value + 1
   For ChkExt = 1 To 7
      If InStr(1, f1.Name, SuppExt(ChkExt).strData, vbTextCompare) Then
      PicNum = PicNum + 1
      Pics(PicNum) = f1.Name
      End If
    Next ChkExt
Next
End If
End Function

