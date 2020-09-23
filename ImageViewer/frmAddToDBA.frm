VERSION 5.00
Begin VB.Form frmAddToDBA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Image"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   11385
   Begin VB.OptionButton optCompressed 
      Caption         =   "Compress Image"
      Height          =   255
      Left            =   8160
      TabIndex        =   9
      Top             =   2640
      Width           =   3135
   End
   Begin VB.ComboBox cboImgSubMinType 
      Height          =   315
      Left            =   8160
      TabIndex        =   2
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   9720
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveImg 
      Caption         =   "Save Image"
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.ComboBox cboImgSubType 
      Height          =   315
      Left            =   8160
      TabIndex        =   1
      Top             =   1440
      Width           =   3135
   End
   Begin VB.ComboBox cboImgType 
      Height          =   315
      Left            =   8160
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Frame frmImage 
      Caption         =   "Image to Save in Database"
      Height          =   7695
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7935
      Begin VB.Image imgDBA 
         Height          =   7335
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   7695
      End
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   8160
      TabIndex        =   8
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   8160
      TabIndex        =   7
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   8160
      TabIndex        =   6
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmAddToDBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboImgSubMinType_KeyPress(KeyAscii As Integer)
KeyAscii = ValidateInput(KeyAscii, Text_Input, True)
End Sub

Private Sub cboImgSubType_KeyPress(KeyAscii As Integer)
KeyAscii = ValidateInput(KeyAscii, Text_Input, True)

End Sub

Private Sub cboImgSubType_LostFocus()
LoadImgSubMinType

End Sub

Private Sub cboImgType_LostFocus()
LoadImgSubTypes
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSaveImg_Click()
Dim ImgID As Integer
Dim SaveImg As ADODB.Recordset
Dim GetImgID As ADODB.Recordset
Dim FileType As String
Dim PicData As String
Dim FileExt As String

Set SaveImg = New ADODB.Recordset
Set GetImgID = New ADODB.Recordset
GetImgID.Open "Select Max(ImageID) as ImgID from imagefile", DB, adOpenStatic, adLockOptimistic
SaveImg.Open "Select * from ImageFile", DB, adOpenStatic, adLockOptimistic
ImgID = 0
If IsNull(GetImgID!ImgID) = True Then
 ImgID = 1
ElseIf IsNull(GetImgID!ImgID) = False Then
ImgID = GetImgID!ImgID + 1
End If


PicData = ReadFromFile(FileName, optCompressed.Value)
FileExt = Right(FileName, 3)


If cboImgType.Text <> "" Then
With SaveImg
   .AddNew
   !imageid = ImgID
   !Image = PicData
   '!Image.AppendChunk PicData
   !imagetype = cboImgType.Text
   !subtype = ChgNull(cboImgSubType.Text)
   !subminType = ChgNull(cboImgSubMinType.Text)
   !ImgExt = UCase(FileExt)
   !comp = optCompressed.Value
   .Update
End With
Unload frmAddToDBA
ElseIf cboImgType.Text = "" Then
 MsgBox "Must Supply a Image Type"
End If
End Sub
Private Sub Form_Activate()
LoadImgTypes
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'Find why Capttal L cases this function to activate

    If (KeyAscii = vbKeyReturn) Then 'Or (KeyAscii = vbKeySeparator) Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub Form_Load()
Me.Height = 8295
Me.Width = 11475
Me.Move (frmMain.Width / 2) - (Me.Width / 2), (frmMain.Height / 2) - ((Me.Height / 2) + 380)

End Sub

Public Sub AddtoDba(FilePath As String)
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

frmAddToDBA.Show
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

