VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmThumbs 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThumbs.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThumbs.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThumbs.frx":08A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThumbs.frx":1184
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   23
      Left            =   7920
      TabIndex        =   24
      Top             =   5040
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   23
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   22
      Left            =   6360
      TabIndex        =   23
      Top             =   5040
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   22
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   21
      Left            =   4800
      TabIndex        =   22
      Top             =   5040
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   21
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   20
      Left            =   3240
      TabIndex        =   21
      Top             =   5040
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   20
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   19
      Left            =   1680
      TabIndex        =   20
      Top             =   5040
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   19
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   18
      Left            =   120
      TabIndex        =   19
      Top             =   5040
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   18
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   17
      Left            =   7920
      TabIndex        =   18
      Top             =   3480
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   17
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   16
      Left            =   6360
      TabIndex        =   17
      Top             =   3480
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   16
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   15
      Left            =   4800
      TabIndex        =   16
      Top             =   3480
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   15
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   14
      Left            =   3240
      TabIndex        =   15
      Top             =   3480
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   14
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   13
      Left            =   1680
      TabIndex        =   14
      Top             =   3480
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   13
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   12
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   12
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   11
      Left            =   7920
      TabIndex        =   12
      Top             =   1920
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   11
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   10
      Left            =   6360
      TabIndex        =   11
      Top             =   1920
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   10
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   9
      Left            =   4800
      TabIndex        =   10
      Top             =   1920
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   9
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   8
      Left            =   3240
      TabIndex        =   9
      Top             =   1920
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   8
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   7
      Left            =   1680
      TabIndex        =   8
      Top             =   1920
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   7
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   6
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   5
      Left            =   7920
      TabIndex        =   6
      Top             =   360
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   5
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   4
      Left            =   6360
      TabIndex        =   5
      Top             =   360
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   4
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   3
      Left            =   4800
      TabIndex        =   4
      Top             =   360
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   3
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   2
      Left            =   3240
      TabIndex        =   3
      Top             =   360
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   2
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   360
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   1
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Frame frmImage 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1515
      Begin VB.Image imgThumbs 
         Height          =   1500
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1500
      End
   End
   Begin MSComctlLib.Toolbar tlbThumbs 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   582
      ButtonWidth     =   2540
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Previous Page"
            Key             =   "back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Next Page"
            Key             =   "nextpage"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Slide Show"
            Key             =   "slideshow"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "close"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPageNum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Page 0 of 0"
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
      TabIndex        =   25
      Top             =   6600
      Width           =   9255
   End
   Begin VB.Image imgTemp 
      Height          =   1980
      Left            =   3360
      Top             =   2400
      Visible         =   0   'False
      Width           =   1620
   End
End
Attribute VB_Name = "frmThumbs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strPath As String
Dim First As Integer
Dim Page As Integer
Dim Pages As Integer
Dim NumOfPics As Integer
Dim Pics() As String

Private Sub Form_Load()
First = 0
Page = 0
Pages = 0
frmThumbs.Height = 7020
frmThumbs.Width = 9660
frmThumbs.Move (frmMain.Width / 2) - (frmThumbs.Width / 2), (frmMain.Height / 2) - ((frmThumbs.Height / 2) + 360)
Pages = NumOfPics / 24
If NumOfPics <= 24 Then
 tlbThumbs.Buttons("nextpage").Enabled = False
End If
If Pages * 24 <= NumOfPics Then
 Pages = Pages + 1
End If
Page = Page + 1
lblPageNum.Caption = "Pages : " & Page & " of " & Pages
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
frmLoading.Caption = "Creating thumbnails, please wait..."
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

Public Sub DisplayThumbs()
  Dim ThumbNum As Integer
  For ThumbNum = 0 To 23
   First = First + 1
   DisplayImg ThumbNum, First
  If First > NumOfPics Then Exit Sub
  Next ThumbNum
End Sub

Private Sub DisplayImg(imgIndex As Integer, PicNum As Integer)
Dim imgWidth As Integer
Dim imgHeight As Integer
Dim NewH As Integer
Dim NewW As Integer

On Error GoTo openerror
   
   If VerifyFile(strPath & Pics(PicNum)) Then
      imgTemp.Picture = LoadPicture(strPath & Pics(PicNum))
      imgWidth = imgTemp.Width 'imgThumb(imgIndex).Width
      imgHeight = imgTemp.Height 'imgThumb(imgIndex).Height
      ThumbSize imgWidth, imgHeight, NewW, NewH
      imgThumbs(imgIndex).Move (frmImage(imgIndex).Width - NewW) / 2, (frmImage(imgIndex).Height - NewH) / 2, NewW, NewH
      imgThumbs(imgIndex).Picture = LoadPicture(strPath & Pics(PicNum))
      imgThumbs(imgIndex).ToolTipText = "Click to open Image " & Pics(PicNum)
      imgThumbs(imgIndex).Tag = strPath & Pics(PicNum)
      DoEvents
      Exit Sub
   Else
      imgThumbs(imgIndex).Picture = LoadPicture()
      imgThumbs(imgIndex).Tag = ""
      Exit Sub
   End If
openerror:
 imgThumbs(imgIndex).Picture = LoadPicture()
 imgThumbs(imgIndex).Tag = ""
End Sub

Private Sub imgThumbs_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
 If imgThumbs(Index).Tag <> "" Then frmMain.OpenImage imgThumbs(Index).Tag
ElseIf Button = 2 Then
 If imgThumbs(Index).Tag <> "" Then FileName = imgThumbs(Index).Tag
 PopupMenu frmMain.mnuDba
End If
End Sub



Private Sub tlbThumbs_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "back"
         BackAPage
    Case "nextpage"
         NextPage
    Case "slideshow"
         frmLoading.Show
         With frmSlideShow
          NumOfPics = .ScanFolder(strPath)
           If NumOfPics > 1 Then
             .NextPic
             .CenterImage
             .WindowState = vbMaximized
             .Show
             .StartSlideShow True
           End If
         End With
        Unload frmLoading
    Case "close"
     Unload Me
     frmMain.mnuThumbNails.Enabled = True
End Select
End Sub

Private Sub BackAPage()
Dim J As Integer
Dim ImgNum As Integer
Dim NewImgNum As Integer
ImgNum = First
ImgNum = ImgNum - 24
For J = 0 To 23
  NewImgNum = (ImgNum - 23) + J
  DisplayImg J, NewImgNum
Next
First = ImgNum
tlbThumbs.Buttons("nextpage").Enabled = True
If First <= 24 Then
 tlbThumbs.Buttons("back").Enabled = False
End If
Page = Page - 1
lblPageNum.Caption = "Pages : " & Page & " of " & Pages
End Sub

Private Sub NextPage()
Dim ImgNum As Integer
Dim J As Integer
ImgNum = First
For J = 0 To 23
   ImgNum = ImgNum + 1
   DisplayImg J, ImgNum
Next J
First = ImgNum
tlbThumbs.Buttons("back").Enabled = True
If ImgNum >= NumOfPics Then
 tlbThumbs.Buttons("nextpage").Enabled = False
End If
Page = Page + 1
lblPageNum.Caption = "Pages : " & Page & " of " & Pages
End Sub

Private Sub ThumbSize(iwidth As Integer, iheight As Integer, iNewW As Integer, iNewH As Integer)
Dim T As Integer
iNewW = iwidth: iNewH = iheight
T = 1
Do While iNewW > 1500 Or iNewH > 1500
 iNewW = Int(iwidth / T)
 iNewH = Int(iheight / T)
 T = T + 1
Loop
End Sub

