VERSION 5.00
Begin VB.Form frmSelectImageTypes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Image Types"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4785
   Begin VB.ComboBox cboImgSubMinType 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox cboImgSubType 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.ComboBox cboImgType 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Image Sub Min Type :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   870
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Image Sub Types :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Image Types :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "frmSelectImageTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DOpt As String


Private Sub cboImgSubType_GotFocus()
cboImgsubMinType.Clear
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

Private Sub cmdClose_Click()
frmMain.mnuDBThumNails.Enabled = True
Unload Me
End Sub
Private Sub cmdCreate_Click()
If DOpt = "Slideshow" Then
With frmDBSlideShow
  .CreateSlideShow cboImgType.Text, cboImgSubType.Text, cboImgsubMinType.Text
End With
ElseIf DOpt = "Thumbnails" Then
With frmDBThumbnails
   .CreateThumbs cboImgType.Text, cboImgSubType.Text, cboImgsubMinType.Text
   '.Show
End With
End If
End Sub

Private Sub Form_Load()
Me.Height = 2160
Me.Width = 4875
Me.Move (frmMain.Width / 2) - (Me.Width / 2), (frmMain.Height / 2) - ((Me.Height / 2) + 380)
LoadImgTypes
End Sub

Property Let DisplayOpt(Opt As String)
DOpt = Opt
End Property

Property Get DisplayOpt() As String
DisplayOpt = DOpt
End Property

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

Private Sub LoadImgSubMinType()
Dim ImgSubMinType As ADODB.Recordset
Dim SqlStr As String
Set ImgSubMinType = New ADODB.Recordset

SqlStr = "Select distinct SubMinType from imagefile where subtype ='" & cboImgSubType.Text & "' group by submintype"

ImgSubMinType.Open SqlStr, DB, adOpenStatic, adLockOptimistic

If ImgSubMinType.RecordCount <> 0 Then
 ImgSubMinType.MoveFirst
  Do While Not ImgSubMinType.EOF
   cboImgsubMinType.AddItem (ImgSubMinType("SubminType"))
   ImgSubMinType.MoveNext
  Loop
End If
End Sub
