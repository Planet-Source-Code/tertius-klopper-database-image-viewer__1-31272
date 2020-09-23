VERSION 5.00
Begin VB.Form frmModifyImage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Image in Database"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   11370
   Begin VB.ComboBox cboImgsubMinType 
      Height          =   315
      Left            =   8160
      TabIndex        =   2
      Top             =   5520
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Current Settings"
      Height          =   3015
      Left            =   8160
      TabIndex        =   8
      Top             =   240
      Width           =   3015
      Begin VB.Label lblFileExt 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Image Extention :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label lblImgSubMinType 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Image Sub Min Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label lblImgSubType 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Image Sub Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblImgType 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Image Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblImgID 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Image ID :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   9720
      TabIndex        =   4
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveChanges 
      Caption         =   "Save Changes"
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   6000
      Width           =   1455
   End
   Begin VB.ComboBox cboImgSubType 
      Height          =   315
      Left            =   8160
      TabIndex        =   1
      Top             =   4680
      Width           =   3015
   End
   Begin VB.ComboBox cboImgType 
      Height          =   315
      Left            =   8160
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Frame frmImage 
      Caption         =   "Image"
      Height          =   7695
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7935
      Begin VB.Image imgDBA 
         Height          =   7335
         Left            =   120
         Top             =   240
         Width           =   7695
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Image Sub Min Type"
      Height          =   255
      Left            =   8160
      TabIndex        =   19
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Image Sub Type"
      Height          =   255
      Left            =   8160
      TabIndex        =   7
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Image Type :"
      Height          =   255
      Left            =   8160
      TabIndex        =   6
      Top             =   3480
      Width           =   3015
   End
End
Attribute VB_Name = "frmModifyImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ImgID As Integer
Dim ChangeImgRS As ADODB.Recordset
Dim m_strPicData As String
Dim m_lngPicLen As Long


Private Sub cboImgType_LostFocus()
If cboImgType.Text <> "" Then
LoadImgSubTypes
ElseIf cboImgType.Text = "" Then
cboImgType.SetFocus
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSaveChanges_Click()
ChangeImgRS!imagetype = cboImgType.Text

If cboImgSubType.Text = "" Then
  ChangeImgRS!subtype = ""
  ChangeImgRS!subminType = ""
Else
    ChangeImgRS!subtype = cboImgSubType.Text
End If

ChangeImgRS!subminType = ChgNull(cboImgSubMinType.Text)
ChangeImgRS.Update
Unload Me
End Sub

Private Sub Form_Load()
Me.Height = 8295
Me.Width = 11475
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


Private Sub ThumbSize(iwidth As Integer, iheight As Integer, iNewW As Integer, iNewH As Integer)
Dim T As Integer
iNewW = iwidth: iNewH = iheight
T = 1
Do While iNewW > 7935 Or iNewH > 7695
 iNewW = Int(iwidth / T)
 iNewH = Int(iheight / T)
 T = T + 1
Loop
End Sub
Public Sub LookUpImage(iImgID As Integer)

Set ChangeImgRS = New ADODB.Recordset
ImgID = iImgID
ChangeImgRS.Open "Select * from imagefile where imageid =" & ImgID, DB, adOpenStatic, adLockOptimistic
If ChangeImgRS.RecordCount <> 0 Then
 LoadImage ReadFromDba(ChangeImgRS!Image, "", ChangeImgRS!comp)
 lblImgID.Caption = ChangeImgRS!imageid
 lblImgType.Caption = ChangeImgRS!imagetype
 lblImgSubType.Caption = ChangeImgRS!subtype
 lblImgSubMinType.Caption = ""
 lblFileExt.Caption = ChangeImgRS!ImgExt
 cboImgType.Text = ChangeImgRS!imagetype
 cboImgSubType.Text = ChangeImgRS!subtype
 cboImgSubMinType.Text = ChangeImgRS!subminType
End If

End Sub

Private Sub LoadImage(FileName As String)
Dim imgWidth As Integer
Dim imgHeight As Integer
Dim NewH As Integer
Dim NewW As Integer
imgDBA.Visible = False
imgDBA.Stretch = False
imgDBA.Picture = LoadPicture(FileName)
imgWidth = imgDBA.Width
imgHeight = imgDBA.Height
ThumbSize imgWidth, imgHeight, NewW, NewH
imgDBA.Move (frmImage.Width - NewW) / 2, (frmImage.Height - NewH) / 2, NewW, NewH
imgDBA.Picture = LoadPicture()
imgDBA.Visible = True
imgDBA.Stretch = True
imgDBA.Picture = LoadPicture(FileName)
frmModifyImage.Show
End Sub


