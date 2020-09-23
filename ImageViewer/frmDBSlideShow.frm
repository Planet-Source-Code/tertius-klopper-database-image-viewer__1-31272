VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDBSlideShow 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrSlideShow 
      Interval        =   5000
      Left            =   360
      Top             =   1320
   End
   Begin MSComctlLib.Toolbar tbrSlideShowOpt 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   635
      ButtonWidth     =   1852
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Key             =   "forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Start"
            Key             =   "start"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "stop"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Slower"
            Key             =   "slower"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Faster"
            Key             =   "faster"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Hide Info"
            Key             =   "hide"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   8
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "close"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBSlideShow.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBSlideShow.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBSlideShow.frx":08A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBSlideShow.frx":1184
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBSlideShow.frx":1A60
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBSlideShow.frx":1EB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBSlideShow.frx":2308
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBSlideShow.frx":2BE4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   9705
   End
   Begin VB.Image imgImage 
      Height          =   3615
      Left            =   2535
      Top             =   743
      Width           =   3975
   End
End
Attribute VB_Name = "frmDBSlideShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DBAImgRS As ADODB.Recordset
'Public ImgID As Integer
Public ImgType As String
Public ImgSubType As String
Public ImgSubMinType As String
Dim strPath As String
Dim NumOfPics As Integer
Dim SqlStr As String
Dim CurrentPic As Integer
Dim zoom As Integer
Private iwidth As Double
Private iheight As Double

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case vbKeyLeft
      StopSlideShow
      PrevPic
      DisplayPic
   Case vbKeyRight
      StopSlideShow
      NextPic
      DisplayPic
   Case vbKeyUp
     If tmrSlideShow.Interval < 65000 Then
      tmrSlideShow.Interval = tmrSlideShow.Interval + 100
      lblInfo.Caption = "Slide Speed : " & Round(tmrSlideShow.Interval / 1000, 1) & " sec"
     End If
   Case vbKeyDown
     If tmrSlideShow.Interval > 100 Then
      tmrSlideShow.Interval = tmrSlideShow.Interval - 100
      lblInfo.Caption = "Slide Speed : " & Round(tmrSlideShow.Interval / 1000, 1) & " sec"
     End If
   Case vbKeyReturn
    If tbrSlideShowOpt.Buttons("start").Enabled = True Then
      StartSlideShow
    ElseIf tbrSlideShowOpt.Buttons("start").Enabled = False Then
      StopSlideShow
    End If
   Case vbKeyEscape
    Unload Me
End Select

End Sub

Private Sub Form_Load()
Me.WindowState = vbMaximized
End Sub
Public Sub CreateSlideShow(iImgType As String, iImgSubType As String, iImgSubMinType As String)
ImgType = iImgType
ImgSubType = iImgSubType
ImgSubMinType = iImgSubMinType
Set DBAImgRS = New ADODB.Recordset
SqlStr = "Select * from imagefile where imagetype = '" & ImgType & _
"' and subtype = '" & ImgSubType & "' and submintype ='" & ImgSubMinType & "'"
DBAImgRS.Open SqlStr, DB, adOpenStatic, adLockOptimistic
If DBAImgRS.RecordCount <> 0 Then
NumOfPics = DBAImgRS.RecordCount
Me.Show
DBAImgRS.MoveFirst
Unload frmSelectImageTypes
StartSlideShow True
ElseIf DBAImgRS.RecordCount = 0 Then
MsgBox "No Image for Type " & ImgType
frmMain.mnuDBThumNails.Enabled = True
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y > tbrSlideShowOpt.Height * 2 Then
   tbrSlideShowOpt.Visible = False
   lblInfo.Top = 0
Else
   tbrSlideShowOpt.Visible = True
   lblInfo.Top = tbrSlideShowOpt.Height
End If
End Sub

Private Sub Form_Resize()
lblInfo.Move 0, 0, Me.Width
lblInfo.Caption = "Slide Speed : " & Round(tmrSlideShow.Interval / 1000, 1) & " sec"

End Sub

Private Sub tbrSlideShowOpt_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
  Case "back"
    StopSlideShow
    PrevPic
    DisplayPic
  Case "forward"
    StopSlideShow
    NextPic
    DisplayPic
  Case "start"
   StartSlideShow
  Case "stop"
   StopSlideShow
  Case "slower"
   If tmrSlideShow.Interval < 65000 Then
     tmrSlideShow.Interval = tmrSlideShow.Interval + 100
     lblInfo.Caption = "Slide Speed : " & Round(tmrSlideShow.Interval / 1000, 1) & " sec"
   End If
  Case "faster"
    If tmrSlideShow.Interval > 100 Then
     tmrSlideShow.Interval = tmrSlideShow.Interval - 100
     lblInfo.Caption = "Slide Speed : " & Round(tmrSlideShow.Interval / 1000, 1) & " sec"
    End If
  Case "hide"
    If tbrSlideShowOpt.Buttons("hide").Caption = "Hide Info" Then
      tbrSlideShowOpt.Buttons("hide").Caption = "Show Info"
      lblInfo.Visible = False
    ElseIf tbrSlideShowOpt.Buttons("hide").Caption = "Show Info" Then
      tbrSlideShowOpt.Buttons("hide").Caption = "Hide Info"
      lblInfo.Visible = True
    End If

  Case "close"
   Unload Me
  End Select

End Sub
Public Sub CenterImage()
 zoom = 0
If (imgImage.Width <> iwidth) Or (imgImage.Height <> iheight) Then
    imgImage.Width = iwidth
    imgImage.Height = iheight
End If
    imgImage.Move (Me.Width / 2) - (imgImage.Width / 2), (Me.Height / 2) - (imgImage.Height / 2)
End Sub
Public Sub StartSlideShow(Optional Restart As Boolean)
tbrSlideShowOpt.Visible = False
DisplayPic Restart
tmrSlideShow.Enabled = True
tbrSlideShowOpt.Buttons("start").Enabled = False
tbrSlideShowOpt.Buttons("stop").Enabled = True
End Sub
Public Sub StopSlideShow()
tmrSlideShow.Enabled = False
tbrSlideShowOpt.Buttons("start").Enabled = True
tbrSlideShowOpt.Buttons("stop").Enabled = False
End Sub
Public Sub NextPic()
    If CurrentPic < NumOfPics Then
      CurrentPic = CurrentPic + 1
      DBAImgRS.MoveNext
    Else
      CurrentPic = 1
      DBAImgRS.MoveFirst
    End If
End Sub
        
Public Sub PrevPic()
    If CurrentPic = 1 Then
      CurrentPic = NumOfPics
      DBAImgRS.MoveLast
    Else
      CurrentPic = CurrentPic - 1
      DBAImgRS.MovePrevious
    End If
End Sub

Public Sub DisplayPic(Optional Restart As Boolean)
On Error GoTo openerror
If Restart Then CurrentPic = 1
' If VerifyFile(strPath & Pics(CurrentPic)) Then
   zoom = 0
   imgImage.Visible = False
   imgImage.Move 0, 0
   imgImage.Stretch = False
   imgImage.Picture = LoadPicture(ReadFromDba(DBAImgRS!Image, "", DBAImgRS!comp))
'   imgImage.Picture = LoadPicture(strPath & Pics(CurrentPic))
    While imgImage.Width > 12000
        imgImage.Width = imgImage.Width * 0.9
        imgImage.Height = imgImage.Height * 0.9
        imgImage.Stretch = True
    Wend
    While imgImage.Height > 9000
       imgImage.Height = imgImage.Height * 0.9
       imgImage.Width = imgImage.Width * 0.9
       imgImage.Stretch = True
    Wend
    imgImage.Visible = True
    iwidth = imgImage.Width
    iheight = imgImage.Height
    CenterImage
    lblInfo.Caption = "Pic : " & CurrentPic & " of " & NumOfPics
' Else
'   MsgBox "File Open Error"
'   Exit Sub
' End If
openerror:
'Open Error
End Sub

'Private Function OpenDBAImage() As String
'Dim TempFile As String
'Dim FileExt As String
'    m_lngPicLen = DBAImgRS.Fields("Image").ActualSize
'    If m_lngPicLen > 0 Then
'        m_strPicData = DBAImgRS.Fields("Image").GetChunk(m_lngPicLen)
'        FileExt = Right(m_strPicData, 3)
'        TempFile = App.Path & "\tmp.bmp"
'        Open TempFile For Binary As #1
'            Put #1, , m_strPicData
'        Close #1
'        OpenDBAImage = TempFile
'    Else
'        OpenDBAImage = ""
'    End If
'End Function


Private Sub tmrSlideShow_Timer()
NextPic
DisplayPic
End Sub

