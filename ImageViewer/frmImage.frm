VERSION 5.00
Begin VB.Form frmImages 
   BackColor       =   &H00000000&
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2775
   ScaleWidth      =   4440
   WindowState     =   2  'Maximized
   Begin VB.Image imgImage 
      Height          =   2775
      Left            =   120
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private zoom As Integer
Private iwidth As Double
Private iheight As Double
Public Sub LoadImage(FileName As String)
  On Error GoTo openerror:
    If VerifyFile(FileName) Then
      imgImage.Visible = False
      imgImage.Move 0, 0
      imgImage.Picture = LoadPicture(FileName)
      While imgImage.Width > frmWidth '12000
        imgImage.Width = imgImage.Width * 0.9
        imgImage.Height = imgImage.Height * 0.9
        imgImage.Stretch = True
      Wend
      While imgImage.Height > frmHeight '9000
        imgImage.Height = imgImage.Height * 0.9
        imgImage.Width = imgImage.Width * 0.9
        imgImage.Stretch = True
      Wend
      imgImage.Visible = True
      iwidth = imgImage.Width
      iheight = imgImage.Height
      CenterImage
      Exit Sub
    Else
      MsgBox "File open error", vbOKOnly, "File could not be verified"
      Me.Tag = "error"
      Exit Sub
    End If
openerror:
    MsgBox "File open error", vbOKOnly, "Eyebrowse"
    Me.Tag = "error"
    Exit Sub
End Sub
Public Sub CenterImage()
 zoom = 0
If (imgImage.Width <> iwidth) Or (imgImage.Height <> iheight) Then
    imgImage.Width = iwidth
    imgImage.Height = iheight
End If
    imgImage.Move (Me.Width / 2) - (imgImage.Width / 2), (Me.Height / 2) - ((imgImage.Height / 2) + 185) '185 is for menu
End Sub
Private Sub Form_Resize()
CenterImage
End Sub

Private Sub imgImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
     If zoom < 5 Then
       zoom = zoom + 1
       imgImage.Visible = False
       imgImage.Stretch = True
       imgImage.Width = imgImage.Width * 2
       imgImage.Height = imgImage.Height * 2
       imgImage.Move (Me.Width / 2) - (X * 2), (Me.Width / 2) - (Y * 2)
       imgImage.Visible = True
     End If
  ElseIf Button = vbRightButton Then
     If zoom > 0 Then
       zoom = zoom - 1
       imgImage.Visible = False
       imgImage.Stretch = True
       imgImage.Width = imgImage.Width / 2
       imgImage.Height = imgImage.Height / 2
       If zoom > 0 Then
         imgImage.Move (Me.Width / 2) - (X / 2), (Me.Width / 2) - (Y / 2)
       Else
         CenterImage
       End If
         imgImage.Visible = True
     End If
  End If
End Sub

